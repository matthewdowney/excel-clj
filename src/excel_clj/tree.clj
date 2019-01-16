(ns
  ^{:doc "A tree for Excel (accounting) data. The format is [Label [Children]]
    for nodes and [Label {:column :value}] for leaves.

    For any tree, t, the value function returns the sum of the {:column :value}
    attributes under the root.
         (let [t [:everything
                  [[:child-1 {:usd 10M :mxn 10M}]
                   [:child-2 {:usd 5M :mxn -3M}]]]]
           (value t))
    ; => {:usd 15M :mxn 7M}

    See the example at the bottom of the ns for some code to render a balance
    sheet."
    :author "Matthew Downey"} excel-clj.tree
  (:require
    [clojure.string :as string]))

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

(defn seq-tree
  "The opposite of core/tree-seq: construct a tree given the root and traversal
  functions."
  ([branch? children root]
   (seq-tree branch? children root identity))
  ([branch? children root leaf-factory]
   (letfn [(traverse [node]
             (if-not (branch? node)
               (leaf-factory node)
               (fn []
                 [node (mapv #(trampoline traverse %) (children node))])))]
     (trampoline traverse root))))

(defn force-map [tree-or-map]
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

;;; Modify a tree by mapping over its nodes and reconstructing

(defn map-nodes
  "Map f across all [label attrs] and [label [child]] nodes."
  [tree f]
  (letfn [(map' [node]
            (if (leaf? node)
              (f node)
              (fn []
                (let [recurred (mapv #(trampoline map' %) (children node))
                      children' (or
                                  (not-empty (vec (filter some? recurred)))
                                  {})]
                  (f [(label node) children'])))))]
    (trampoline map' tree)))

(defn map-leaves
  "Map f across all leaf nodes."
  [tree f]
  (map-nodes tree (fn [node] (cond-> node (leaf? node) f))))

(defn map-leaf-vals
  "Map f across the value map of all leaves."
  [tree f]
  (map-leaves tree (fn [[label attrs]] [label (f attrs)])))

(defn negate-tree
  "Negate all of the numbers in a tree."
  [tree]
  (map-leaf-vals tree negate-map))

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

(defn render-table
  "Render the trees into a 'table': a coll of maps, each map containing :depth
  (distance from the root) and :label (the first part of the [Label [Children]]
  node vector) attributes plus the (attrs) for the node/leaf."
  [& trees]
  ;; All the attr keys plus :depth (root is the max depth) and :label
  (loop [ts trees rendered []]
    (if-let [t (first ts)]
      (if (map? t) ;; It's an already processed line, just add it an move on
        (recur (rest ts) (conj rendered t))
        (let [t-depth (nth t 2 0)]
          (if (leaf? t) ;; It's a leaf, so display with all of its attributes
            (let [line (merge
                         {:depth (max *min-leaf-depth* t-depth)
                          :label (label t)}
                         (value t))]
              (recur (rest ts) (conj rendered line)))
            ;; It's a node, so add a header, display its children, then a total
            (let [;; w/ depth
                  children' (mapv #(assoc % 2 (inc t-depth)) (children t))
                  total  (merge {:depth t-depth :label ""} (value t))
                  ;; Only one child and it's a leaf, no need for a total
                  show-total? (not (and (= (count (children t)) 1)
                                        (leaf? (first (children t)))))
                  header {:depth t-depth :label (label t)}
                  next-lines (into
                               (cond-> children' show-total? (conj total))
                               (rest ts))]
              (recur next-lines (conj rendered header))))))
      rendered)))

(defn headers
  "Return a vector of headers in the tree, with any headers given in first-hs
  at the beginning and and in last-hs in order."
  [tree first-hs last-hs]
  (let [all-specified (into first-hs last-hs)
        all-headers (set (keys (value tree)))
        other-headers (apply disj all-headers all-specified)]
    (vec (filter all-headers (concat first-hs other-headers last-hs)))))

(defn tbl
  "Display tabular data in a way that preserves label indentation in a way the
  clojure.pprint/print-table does not."
  ([xs]
   (tbl
     (sequence (comp (mapcat keys) (distinct)) xs)
     xs))
  ([ks xs]
   (let [pad 2
         max' (completing #(max %1 %2))
         len (fn [k]
               (let [len' #(or (some-> (% k) str count) 0)]
                 (+ pad (transduce (map len') max' 0 xs))))
         header (into {} (map (juxt identity identity)) ks)
         ks' (mapv (juxt identity len) ks)]
     (doseq [x (cons header xs)]
       (doseq [[k l] ks']
         (print (format (str "%-" l "s") (get x k "-"))))
       (println "")))))

(defn render-tbl-string
  [& trees]
  (letfn [(fmt [{:keys [depth label] :as line}]
            (-> line
                (dissoc :depth :label)
                (assoc "" (str (apply str (repeat depth "  ")) label))))]
    (with-out-str
      (tbl (map fmt (apply render-table trees))))))

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

(defn example []
  ;; Render the tree as a table
  (println (apply render-tbl-string mock-balance-sheet))

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

