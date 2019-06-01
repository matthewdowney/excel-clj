(ns
  ^{:doc "A key-value tree for Excel (accounting) data. The format is
    [Label [Children]] for nodes and [Label {:column :value}] for leaves.

    For some example code, check out the various (comment ...) blocks in this
    namespace."
    :author "Matthew Downey"} excel-clj.tree
  (:require [clojure.string :as string]
            [clojure.walk :as cwalk]))

;;; Utilities for vector math

(defn sum-maps
  "Similar to (merge-with + ...) but treats nil keys as 0 values."
  ([m1 m2] ;; nil == 0
   (let [all-keys (into #{} (concat (keys m1) (keys m2)))]
     (into {} (map (fn [k] [k (+ (or (m1 k) 0) (or (m2 k) 0))])) all-keys)))
  ([m1 m2 & ms]
   (reduce sum-maps {} (into [m1 m2] ms))))

(defn negate-map
  [m]
  (into {} (map (fn [[k v]] [k (* -1 (or v 0))])) m))

(defn subtract-maps
  "Very important difference from (merge-with - ...):

    (merge-with - {:foo 10} {:foo 5 :bar 5})
    ; => {:foo 5, :bar 5}

    (subtract-maps {:foo 10} {:foo 5 :bar 5})
    ; => {:foo 5, :bar -5}
  "
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
  "Aggregate all of the leaf maps in `tree` by reducing over them with
  `reducing-fn` (defaults to summing maps together). If given a single
  map, returns the map."
  ([tree]
   (value tree sum-maps))
  ([tree reducing-fn]
   (cond
     (map? tree) tree
     (leaf? tree) (second tree)
     :else
     (transduce
       (comp (filter leaf?) (map second))
       (completing reducing-fn)
       {}
       (tree-seq (complement leaf?) children tree)))))

(defmacro math
  "Any calls to + or - within form are modified to work on trees and tree
  values (maps of numbers)."
  [form]
  (cwalk/postwalk
    (fn [form]
      (if (and (sequential? form) (#{'+ '-} (first form)))
        (let [replace-with ({'+ `sum-maps '- `subtract-maps} (first form))]
          (cons replace-with (map (fn [tree-expr] (list `value tree-expr)) (rest form))))
        form))
    form))

;;; Utilities for constructing & walking / modifying trees

(defn walk
  "Map f across all [label attrs] and [label [child]] nodes, depth-first.

  Use with the same `branch?` and `children` functions that you'd give to
  `tree-seq` in order to build a tree of the format used by this namespace."
  ([f tree]
   (walk f (complement leaf?) children tree))
  ([f branch? children root]
   (let [walk (fn walk [node]
                (if (branch? node)
                  (f node (mapv walk (children node)))
                  (f node [])))]
     (walk root))))

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

(defn accounting-table
  "Render a coll of trees into a coll of tabular maps, where leaf values are
  listed on the same line and aggregated below into a total (default aggregation
  is addition).

  Each item in the coll is a map with ::depth, ::label, ::header?, and ::total?
  attributes, in addition to the attributes in the leaves.

  If an `:aggregate-with` function is provided, total lines are constructed by
  reducing that function over sub-leaves. Defaults to a reducing with `sum-maps`."
  [trees & {:keys [aggregate-with min-leaf-depth] :or {aggregate-with sum-maps
                                                       min-leaf-depth 2}}]
  (->>
    trees
    (mapcat
      (fn [tree]
        (walk
          (fn [node children]
            (if-let [children (seq (flatten children))]
              (concat
                ;; First we show the header
                [{::depth 0 ::label (label node) ::header? true}]
                ;; Then the children & their values
                (mapv #(update % ::depth inc) children)
                ;; And finally an aggregation if there are multiple header children
                ;; or any leaf children
                (let [fchild (first children)
                      siblings (get (group-by :depth children) (:depth fchild))]
                  (when (or (>= (count siblings) 2) (not (::header? fchild)))
                    [(merge {::depth 0 ::label "" ::total? true} (value node aggregate-with))])))
              ;; A leaf just has its label & value attrs. The depth is inc'd by each
              ;; parent back to the root, so it does not stay at 0.
              (merge {::depth 0 ::label (label node)} (value node aggregate-with))))
          tree)))
    (map
      (fn [table-row]
        ;; not a leaf
        (if (or (::header? table-row) (::total? table-row) (not (::depth table-row)))
          table-row
          (update table-row ::depth max min-leaf-depth))))))

(defn unaggregated-table
  "Similar to account-table, but makes no attempt to aggregate non-leaf headers,
  and accepts a coll of trees."
  [trees]
  (mapcat
    (fn [tree]
      (walk
        (fn [node children]
          (if-let [children (seq (flatten children))]
            (into [{::depth 0 ::label (label node) ::header? true}] (map #(update % ::depth inc)) children)
            (merge {::depth 0 ::label (label node)} (value node))))
        tree))
    trees))

(defn render
  "Given a coll of table items with a qualified ::depth and ::label keys, return
  a table items indenting labels with ::depth and keeping other keys as column
  labels, removing namespace qualified keywords.

  (Used for printing in a string, rather than with Excel.)"
  [table-items & {:keys [indent-width] :or {indent-width 2}}]
  (let [indent-str (apply str (repeat indent-width " "))]
    (letfn [(fmt [line-item]
              (-> line-item
                  (dissoc ::depth ::label ::total? ::header?)
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

(comment
  ;; Render the tree as a table
  (-> mock-balance-sheet accounting-table render print-table)

  ;; Do addition or subtraction with trees using the tree-math macro
  (let [[assets [_ [liabilities equity]]] mock-balance-sheet]
    (println "Assets - Liabilities =" (math (- assets liabilities)))
    (println "Equity =" (value equity))
    (println)
    (println "Equity + Liabilities =" (math (+ equity liabilities)))
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

(comment
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
      (unaggregated-table)
      (render :indent-width 5)
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
