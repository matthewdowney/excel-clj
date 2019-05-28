(ns
  ^{:doc "A key-value tree for Excel (accounting) data. The format is
    [Label [Children]] for nodes and [Label {:column :value}] for leaves.

    For any tree, t, the value function returns the sum of the {:column :value}
    attributes under the root.
         (let [t [:everything
                  [[:child-1 {:usd 10M :mxn 10M}]
                   [:child-2 {:usd 5M :mxn -3M}]]]]
           (value t))
    ; => {:usd 15M :mxn 7M}

    For some example code, see the functions `balance-sheet-example` or
    `tree-table-example` in this namespace."
    :author "Matthew Downey"} excel-clj.tree
  (:require [clojure.string :as string]))

;;; Utilities for vector math

(defn sum-maps
  ([m1 m2] ;; nil == 0
   (let [all-keys (into #{} (concat (keys m1) (keys m2)))]
     (into {} (map (fn [k] [k (+ (or (m1 k) 0) (or (m2 k) 0))])) all-keys)))
  ([m1 m2 & ms]
   (reduce sum-maps {} (into [m1 m2] ms))))

(defn negate-map
  [m]
  (into {} (map (fn [[k v]] [k (* -1M (or v 0))])) m))

(defn subtract-maps
  ([m1 m2]
   (sum-maps m1 (negate-map m2)))
  ([m1 m2 & ms]
   (reduce subtract-maps {} (into [m1 m2] ms))))

;;; Basic tree API

(def leaf?
  "Is the node a leaf?"
  (comp map? second))

(def label
  "The label for a node."
  first)

(defn children
  "The node's children, or [] in the case of a leaf."
  [node]
  (if (leaf? node)
    []
    (second node)))

(defn value
  "If the node is a leaf, returns the map of values for the leaf. Otherwise
  returns the sum of that value map for all children under the node."
  [node]
  (if-not (leaf? node)
    (loop [children' (children node) summed {}]
      (if-let [nxt (first children')]
        (if (leaf? nxt)
          (recur (rest children') (sum-maps summed (second nxt)))
          (recur (concat (rest children') (children nxt)) summed))
        summed))
    (second node)))

(defn force-map
  "Returns the argument if it's a map, otherwise calls `value` on the arg."
  [tree-or-map]
  (if (map? tree-or-map)
    tree-or-map
    (value tree-or-map)))

(defmacro axf
  "Sort of a composition of f with xf, except that xf is applied to _each_
  argument of f. I.e. f's arguments are transformed by the function xf.

  E.g. 'composing' compare with Math/abs to compare absolute values of the
  numbers:
    (defn abs-compare [n n']
      (axf compare Math/abs n n'))

  Or comparing two java.util.Date objects:
    ;; This won't work...
    (< (Date.) (Date.))
    ; => ClassCastException java.util.Date cannot be cast to java.lang.Number

    ;; But this will
    (axf < .getTime (Date.) (Date.))
    ; => false"
  [f xf & args]
  (let [xformed (map #(-> `(~xf ~%)) args)]
    `(~f ~@xformed)))

(defmacro add-trees [& trees]
  `(axf sum-maps force-map ~@trees))

(defmacro subtract-trees [& trees]
  `(axf subtract-maps force-map ~@trees))

(defmacro tree-math
  "Any calls to + or - within form are modified to work on trees and tree
  values."
  [form]
  (clojure.walk/postwalk-replace
    {'+ `add-trees, '- `subtract-trees}
    form))

;;; Utilities for constructing & walking trees

(defn walk
  "Map f across all [label attrs] and [label [child]] nodes."
  ([f tree]
   (walk f (complement leaf?) children tree))
  ([f branch? children root]
   (let [walk (fn walk [node]
                (if (branch? node)
                  (f node (mapv walk (children node)))
                  (f node [])))]
     (walk root))))

(defn ->tree
  "Construct a tree given the same arguments as `tree-seq`.

  Use in conjunction with some mapping function over the tree to build a tree."
  [branch? children root]
  (walk (fn [node children] [node (vec children)]) branch? children root))

(comment
  "For example, create a file tree with nodes listing the :size of each file."
  (walk
    (fn [f xs]
      (if-not (seq xs)
        [(.getName f) {:size (.length f)}]
        [(.getName f) xs]))
    #(.isDirectory %) #(.listFiles %) (clojure.java.io/file ".")))

(defn negate-tree
  "Negate all of the numbers in a tree."
  [tree]
  (walk
    (fn [node children]
      (if-not (seq children)
        [(label node) (negate-map (value node))]
        [(label node) children]))
    tree))

(defn shallow
  "'Shallow' the tree one level by getting rid of the root and combining its
  children. Doesn't modify a leaf."
  [tree]
  (if-not (leaf? tree)
    (let [merged-label (string/join " & " (map label (children tree)))
          merged-children (mapcat
                            (fn [node]
                              (if (leaf? node) [node] (children node)))
                            (children tree))]
      [merged-label (vec merged-children)])
    tree))

(defn merge-trees
  "Merge the children of the provided trees under a single root."
  [root-label & trees]
  (assoc (shallow ["Merged" trees]) 0 root-label))

;;; Coerce a tree format to a tabular format

(def ^:dynamic *min-leaf-depth*
  "For formatting purposes, set this to artificially make a leaf sit at a depth
  of at least this number."
  2)

(defn tree->raw-table
  "Render the trees into a 'table': a coll of maps, each map containing ::depth
  (distance from the root), ::label (the first part of the [Label [Children]]
  node vector), and ::total? attributes plus the (attrs) for the node/leaf.
  
  The intention being that you can then use the information in the namespace
  qualified attributes to render the tree items into a line format that suits
  you."
  [trees & {:keys [sum-totals?] :or {sum-totals? true}}]
  ;; All the attr keys plus ::depth (root is the max depth) and ::label
  (loop [ts trees, rendered []]
    (if-let [t (first ts)]
      (if (map? t) ;; It's an already processed line, just add it and move on
        (recur (rest ts) (conj rendered t))
        (let [t-depth (nth t 2 0)]
          (if (leaf? t) ;; It's a leaf, so display with all of its attributes
            (let [line (merge
                         {::depth (max *min-leaf-depth* t-depth)
                          ::label (label t)}
                         (value t))]
              (recur (rest ts) (conj rendered line)))
            ;; It's a node, so add a header, display its children, then a total
            (let [;; w/ depth
                  children' (mapv #(assoc % 2 (inc t-depth)) (children t))
                  total (when sum-totals?
                          (merge {::depth t-depth ::label "" ::total? true} (value t)))
                  ;; Only one child and it's a leaf, no need for a total even if
                  ;; its enabled
                  show-total? (not (and (= (count (children t)) 1)
                                        (leaf? (first (children t)))))
                  header {::depth t-depth ::label (label t)}
                  next-lines (into
                               (cond-> children' (and total show-total?) (conj total))
                               (rest ts))]
              (recur next-lines (conj rendered header))))))
      rendered)))

(defn raw-table->rendered
  "Given table items with a qualified ::depth and ::label keys, render a table
  string indenting labels with ::depth and keeping other keys as column labels."
  [table-items & {:keys [indent-width] :or {indent-width 2}}]
  (let [indent-str (apply str (repeat indent-width " "))]
    (letfn [(fmt [line-item]
              (-> line-item
                  (dissoc ::depth ::label ::total?)
                  (assoc "" (str (apply str (repeat (::depth line-item) indent-str))
                                 (::label line-item)))))]
      (map fmt table-items))))

(defn print-table
  "Display tabular data in a way that preserves label indentation in a way the
  clojure.pprint/print-table does not."
  ([xs]
   (print-table xs {}))
  ([xs {:keys [ks empty-str pad-width]}]
   (let [ks (or ks (sequence (comp (mapcat keys) (distinct)) xs))
         empty-str (or empty-str "-")
         pad-width (or pad-width 2)]
     (let [len (fn [k]
                 (let [len' #(or (some-> (% k) str count) 0)]
                   (+ pad-width (transduce (map len') (completing max) 0 (conj xs {k k})))))
           header (into {} (map (juxt identity identity)) ks)
           ks' (mapv (juxt identity len) ks)]
       (doseq [x (cons header xs)]
         (doseq [[k l] ks']
           (print (format (str "%-" l "s") (get x k empty-str))))
         (println ""))))))

(defn headers
  "Return a vector of headers in the tree, with any headers given in first-hs
  at the beginning and and in last-hs in order."
  [tree first-hs last-hs]
  (let [all-specified (into first-hs last-hs)
        all-headers (set (keys (value tree)))
        other-headers (apply disj all-headers all-specified)]
    (vec (filter all-headers (concat first-hs other-headers last-hs)))))

(def mock-balance-sheet
  (vector
    ["Assets"
     [["Current Assets"
       [["Cash" {2018 100M, 2017 85M}]
        ["Accounts Receivable" {2018 5M, 2017 45M}]]]
      ["Investments" {2018 100M, 2017 10M}]
      ["Other" {2018 12M, 2017 8M}]]]
    ["Liabilities & Stockholders' Equity"
     [["Liabilities"
       [["Current Liabilities"
         [["Notes payable" {2018 5M, 2017 8M}]
          ["Accounts payable" {2018 10M, 2017 10M}]]]
        ["Long-term liabilities" {2018 100M, 2017 50M}]]]
      ["Equity"
       [["Common Stock" {2018 102M, 2017 80M}]]]]]))

(defn balance-sheet-example []
  ;; Render the tree as a table
  (-> mock-balance-sheet tree->raw-table raw-table->rendered print-table)

  ;; Do addition or subtraction with trees using the tree-math macro
  (let [[assets [_ [liabilities equity]]] mock-balance-sheet]
    (println "Assets - Liabilities =" (tree-math (- assets liabilities)))
    (println "Equity =" (value equity))
    (println)
    (println "Equity + Liabilities =" (tree-math (+ equity liabilities)))
    (println "Assets =" (value assets))))
; =>
;                                     2018 2017
; Assets                              -    -
;   Current Assets                    -    -
;     Cash                            100  85
;     Accounts Receivable             5    45
;                                     105  130
;   Investments                       100  10
;   Other                             12   8
;                                     217  148
; Liabilities & Stockholders' Equity  -    -
;   Liabilities                       -    -
;     Current Liabilities             -    -
;       Notes payable                 5    8
;       Accounts payable              10   10
;                                     15   18
;     Long-term liabilities           100  50
;                                     115  68
;   Equity                            -    -
;     Common Stock                    102  80
;                                     102  80
;                                     217  148
;
; Assets - Liabilities = {2018 102M, 2017 80M}
; Equity = {2018 102M, 2017 80M}
;
; Equity + Liabilities = {2018 217M, 2017 148M}
; Assets = {2018 217M, 2017 148M}

(comment
  ;; Or you can visualize with ztellman/rhizome
  ;; Keep in mind that this requires $ apt-get install graphviz
  (use '(rhizome viz))
  (rhizome.viz/view-tree
    (complement leaf?) children (second mock-balance-sheet)
    :edge->descriptor (fn [x y] (when (leaf? y) {:label (label y)}))
    :node->descriptor #(->{:label (if (leaf? %) (value %) [(label %) (value %)])})))

;;; Coerce a tabular format to a tree format

(defn ordered-group-by
  "Like `group-by`, but returns a [k [v]] seq and doesn't rearrange values except
  to include them in a group. Probably less performant because it has to search
  the built up seq to find the proper key-value store."
  [f xs]
  (letfn [(update-or-add [xs pred update default]
            (loop [xs' [], xs xs]
              (if-let [x (first xs)]
                (if (pred x)
                  (into xs' (cons (update x) (rest xs)))
                  (recur (conj xs' x) (rest xs)))
                (conj xs' default))))
          (assign-to-group [groups x]
            (let [group (f x)]
              (update-or-add
                groups
                #(= (first %) group)
                #(update % 1 conj x)
                [group [x]])))]
    (reduce assign-to-group [] xs)))

(defn table->trees
  "Collapse a tabular collection of maps into a collection of trees, where the
  label at each level of the tree is given by each of `node-fns` and the columns
  displayed are the result of `format-leaf`, which returns a tabular map.

  See the (comment ...) block under this method declaration for an example."
  [tabular format-leaf & node-fns]
  (letfn [(inner-build
            ([root items]
             (vector
               root
               (if (= (count items) 1)
                 (format-leaf (first items))
                 (map #(->["" (format-leaf %)]) items))))
            ([root items below-root & subsequent]
             (vector
               root
               (->> (ordered-group-by below-root items)
                    (mapv (fn [[next-root next-items]]
                            (apply inner-build next-root next-items subsequent)))))))]
    (second (apply inner-build "" tabular node-fns))))

(defn tree-table-example []
  (-> (table->trees
        ;; The table we'll convert to a tree
        [{:from "MXN" :to "AUD" :on "BrokerA" :return (rand)}
         {:from "MXN" :to "USD" :on "BrokerB" :return (rand)}
         {:from "MXN" :to "JPY" :on "BrokerB" :return (rand)}
         {:from "USD" :to "AUD" :on "BrokerA" :return (rand)}]

        ;; The data fields we want to look at
        #(-> {"Return"            (format "%.2f%%" (:return %))
              "Trade Description" (format "%s -> %s" (:from %) (:to %))})

        ;; The top level label -- split by above/below 50% return
        #(if (> (:return %) 0.5) "High Return" "Some Return")

        ;; Then split by which currency we start with
        #(str "Trading " (:from %))

        ;; Finally, by broker
        :on)
      (tree->raw-table :sum-totals? false)
      (raw-table->rendered :indent-width 5)
      (print-table {:empty-str "" :pad-width 5})))

; =>                       Return     Trade Description
;       High Return
;       Trading MXN
;              BrokerA     0.70%      MXN -> AUD
;              BrokerB
;                          0.68%      MXN -> USD
;                          0.93%      MXN -> JPY
;     Some Return
;         Trading USD
;              BrokerA     0.20%      USD -> AUD"
