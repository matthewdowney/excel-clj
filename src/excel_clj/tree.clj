(ns excel-clj.tree
  "Trees are maps, leaves are maps of something->(not a map).

  Use ordered maps (like array-map) to enforce order."
  {:author "Matthew Downey"}
  (:require [clojure.walk :as walk])
  (:import (clojure.lang Named)))


(defn leaf?
  "A leaf is any map whose values are not maps."
  [x]
  (and (map? x) (every? (complement map?) (vals x))))


(defn fold-kvs
  "Fold the `tree` leaves together into one combined leaf calling
  `(f k (get leaf-1 k) (get leaf-2 k))`.

  The function `f` is called for the _union_ of all keys for both leaves,
  so one of the values may be `nil`."
  [f tree]
  (->> tree
       (tree-seq (complement leaf?) vals)
       (filter leaf?)
       (reduce
         (fn [combined leaf]
           (let [all-keys (into (set (keys combined)) (keys leaf))]
             (reduce
               (fn [x k] (update x k #(f k % (get leaf k))))
               combined
               all-keys))))))


(defn fold
  "Fold the `tree` leaves together into one combined leaf calling
  `(f (get leaf-1 k nil-value) (get leaf-2 k nil-value))`.

  E.g. `(fold + tree)` would sum all of the `{label number}` leaves in tree,
  equivalent to `(apply merge-with + all-leaves)`.

  However, `(fold - tree)` is not `(apply merge-with - all-leaves)`. They
  differ because `merge-with` only uses its function in case of collision;
  `(merge-with - {:x 1} {:y 1})` is `{:x 1, :y 1}`. The result with `fold`
  would be `{:x 1, :y -1}`."
  ([f tree]
   (fold f 0 tree))
  ([f nil-value tree]
   (fold-kvs (fn [k x y] (f (or x nil-value) (or y nil-value))) tree)))


(comment
  (fold + {:bitstamp {:btc 1 :xrp 35000}
           :bitmex {:margin {:btc 2}
                    :short-xrp {:btc 1 :xrp -35000}}})
  ;=> {:btc 4, :xrp 0}

  (fold - {:capital {:btc 1} :debt {:btc 1 :mxn 1}})
  ;=> {:btc 0, :mxn -1}
  )


(defn tree
  "Build a tree from the same arguments you would use for `tree-seq`, plus
  `k` and `v` functions for node keys and leaf value maps, respectively."
  [branch? children root k v]
  (let [build (fn build [node]
                (if-let [children (when (branch? node)
                                    (seq (children node)))]
                  {(k node) (apply merge (map build children))}
                  {(k node) (v node)}))]
    (build root)))


(comment
  "E.g. to build a file tree..."
  (let [dir? #(.isDirectory %)
        listfs #(.listFiles %)
        name #(.getName %)
        size (fn [f] {:size (.length f)})]
    (tree dir? listfs (clojure.java.io/file ".") name size))

  "...and then get the total size"
  (fold + *1)
  ;=> {:size 19096768}
  )


(defn negate
  "Invert the sign of every leaf number for a `tree` with leaves of x->number."
  [tree]
  (walk/postwalk
    (fn [x]
      (if (leaf? x)
        (zipmap (keys x) (map - (vals x)))
        x))
    tree))


(def ^{:private true :dynamic true} *depth* nil)
(defn- ->str [x] (if (instance? Named x) (name x) (str x)))
(defn table
  "Given `(fn f [parent-key node depth] => row-map)`, convert `tree` into a
  table of `[row]`.

  If no `f` is provided, the default implementation creates a pivot table with
  no aggregation of groups and a :tree/indent in each row corresponding to the
  depth of the node.

  Pass `(combined-header)` or `(combined-footer)` as `f` to aggregate sub-trees
  according to custom logic (summing by default)."
  ([tree]
   (table
     (fn render [parent node depth]
       (let [row (fold (fn [_ _] nil) node)]
         (cons
           (assoc row "" (->str parent) :tree/indent depth)
           (when-not (leaf? node) (table render node)))))
     tree))
  ([f tree]
   (into [] (mapcat (fn [[k t]] (table f k t))) tree))
  ([f k tree]
   (binding [*depth* (inc (or *depth* -1))]
     (f k tree *depth*))))


(defn combined-header
  "To build a table where each branch node is a row with values equal to its
  combined leaves."
  ([] (combined-header (partial fold +)))
  ([combine-with]
   (fn render [parent node depth]
     (cons
       (assoc
         (combine-with node)
         "" (->str parent)
         :tree/indent depth)
       (when-not (leaf? node) (table render node))))))


(defn combined-footer
  "To build a table where each branch node is followed by its children and then
  a blank-labeled total row at the same :tree/indent as the header with a value
  equal to its combined leaves."
  ([] (combined-footer (partial fold +)))
  ([combine-with]
   (fn render [parent node depth]
     (if-not (leaf? node)
       (let [combined (combine-with node)
             empty-row (zipmap (keys combined) (repeat nil))]
         (concat
           [(assoc empty-row "" (->str parent) :tree/indent depth)] ; header
           (table render node) ; children
           [(assoc combined "" "" :tree/indent depth)])) ; total
       [(assoc node "" (->str parent) :tree/indent depth)]))))


(defn indent
  "Increase the :tree/indent of each table row by `n` (default 1)."
  ([table-rows] (indent table-rows 1))
  ([table-rows n] (map #(update % :tree/indent (fnil + 0) n) table-rows)))


(defn with-table-header
  "Prepend a table header with the given label & indent the following rows."
  [label table-rows]
  (let [[x & xs :as indented] (indent table-rows)
        nil-values (fn [m] (zipmap (keys m) (repeat nil)))]
    (cons
      (-> x
          (dissoc :tree/indent)
          nil-values
          (assoc "" label :tree/indent (dec (:tree/indent x))))
      indented)))


(defn- table-cell [k row width]
  (format (str "%-" width "s") (or (get row k) "-")))


(defn- table-column-widths [ks rows indent-with]
  (let [indent (count indent-with)]
    (reduce
      (fn [k->width row]
        (let [indent-width (fn [k]
                             (if (= k (first ks))
                               (* (get row :tree/indent 0) indent)
                               0))
              k->rwidth (map
                          #(+ (count (table-cell % row 1)) (indent-width %))
                          (keys k->width))]
          (zipmap
            (keys k->width)
            (map max (vals k->width) k->rwidth))))
      (zipmap ks (map (comp count pr-str) ks))
      rows)))


(defn print-table
  "Pretty print a tree with the same signature as `clojure.pprint/print-table`,
  indenting rows according to a :tree/indent attribute.

  E.g. (print-table (table tree))"
  ([rows]
   (let [ks (-> (keys (first rows))
                (set)
                (disj "" :tree/indent))
         labeled? (contains? (set (keys (first rows))) "")]
     (print-table (into (if labeled? [""] []) ks) rows)))
  ([ks rows]
   (let [indent "  "
         k->max (table-column-widths ks rows indent)]
     (doseq [row (cons (zipmap ks ks) rows)
             :let [n-indents (get row :tree/indent 0)]]
       (dotimes [_ n-indents] (print indent))
       (doseq [k ks
               :let [width (get k->max k)
                     indent (if (= k (first ks))
                              (* n-indents (count indent))
                              0)]]
         (print (table-cell k row (- width indent)) " "))
       (println)))))


;;; For example


(def mock-balance-sheet
  {"Assets"
   {"Current Assets" {"Cash"                {2018 100M, 2017 85M}
                      "Accounts Receivable" {2018 5M, 2017 45M}}
    "Investments"    {2018 100M, 2017 10M}
    "Other"          {2018 12M, 2017 8M}}

   "Liabilities & Stockholders' Equity"
   {"Liabilities" {"Current Liabilities"
                    {"Notes payable"    {2018 5M, 2017 8M}
                     "Accounts payable" {2018 10M, 2017 10M}}
                   "Long-term liabilities" {2018 100M, 2017 50M}}
    "Equity"      {"Common Stock" {2018 102M, 2017 80M}}}})


(comment

  ;; Render as tables
  (print-table (table mock-balance-sheet))
  (print-table (table (combined-footer) mock-balance-sheet))
  (print-table (table (combined-header) mock-balance-sheet))

  ;; Do some math to subtract liabilities from assets
  (def assets (get mock-balance-sheet "Assets"))

  (def liabilities
    (get-in
      mock-balance-sheet
      ["Liabilities & Stockholders' Equity" "Liabilities"]))

  (fold + assets)
  ;=> {2018 217M, 2017 148M}

  (fold + liabilities)
  ;=> {2018 115M, 2017 68M}

  ; this should give us the equity amount
  (fold - {:assets (fold + assets) :liabilities (fold + liabilities)})
  ;=> {2018 102M, 2017 80M}

  ;; Print a more complex table illustrating that math
  (def tbl
    (let [blank-line [{"" ""}]]
      (concat
        (with-table-header "Assets" (table assets))
        blank-line
        (with-table-header "Liabilities" (table liabilities))
        blank-line
        (table {"Assets Less Liabilities"
                (fold - {:assets (fold + assets)
                         :liabilities (fold + liabilities)})}))))

  (print-table tbl))
