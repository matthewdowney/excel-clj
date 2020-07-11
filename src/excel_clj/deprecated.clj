(ns ^:deprecated excel-clj.deprecated
  "To provide some minimal backwards compatibility with v1.x"
  (:require [excel-clj.cell :as cell]
            [excel-clj.tree :as tree]
            [clojure.string :as string]
            [taoensso.encore :as enc]))


(defn- best-guess-row-format
  "Try to guess appropriate formatting based on column name and cell value."
  [row-data column]
  (let [column' (string/lower-case column)
        val (get row-data column)]
    (cond
      (and (string? val) (> (count val) 75))
      {:wrap-text true}

      (or (string/includes? column' "percent") (string/includes? column' "%"))
      {:data-format :percent}

      (string/includes? column' "date")
      {:data-format :ymd :alignment :left}

      (decimal? val)
      {:data-format :accounting}

      :else nil)))


(def ^:private default-header-style
  (constantly
    {:border-bottom :thin :font {:bold true}}))


(defn ^:deprecated table
  "Build a sheet grid from the provided collection of tabular data, where each
  item has the format {Column Name, Cell Value}.
  If provided
    headers      is an ordered coll of column names
    header-style is a function header-name => style map for the header.
    data-style   is a function that takes (datum-map, column name) and returns
                 a style specification or nil for the default style."
  [tabular-data & {:keys [headers header-style data-style]
                   :or {data-style (constantly {})}}]
  (let [;; add the headers either in the order they're provided or in the order
        ;; of (seq) on the first datum
        headers (let [direction (if (> (count (last tabular-data))
                                       (count (first tabular-data)))
                                  reverse identity)
                      hs (or headers (sequence (comp (mapcat keys) (distinct))
                                               (direction tabular-data)))]
                  (assert (not-empty hs) "Table headers are not empty.")
                  hs)
        ;; A little hack to keep track of which numbers excel will right
        ;; justify, and therefore which headers to right justify by default
        numeric? (volatile! #{})
        data-cell (fn [col-name row]
                    (let [style (enc/nested-merge
                                  (or (data-style row col-name) {})
                                  (best-guess-row-format row col-name))]
                      (when (or (= (:data-format style) :accounting)
                                (number? (get row col-name "")))
                        (vswap! numeric? conj col-name))
                      (cell/style (get row col-name) style)))
        getters (map (fn [col-name] #(data-cell col-name %)) headers)
        header-style (or header-style
                         ;; Add right alignment if it's an accounting column
                         (fn [name]
                           (cond-> (default-header-style name)
                                   (@numeric? name)
                                   (assoc :alignment :right))))]
    (cons
      (map (fn [x] (cell/style x (header-style x))) headers)
      (map (apply juxt getters) tabular-data))))


(def default-tree-formatters
  {0 {:font {:bold true} :border-bottom :medium}
   1 {:font {:bold true}}
   2 {:indention 2}
   3 {:font {:italic true} :alignment :right}})


(def default-tree-total-formatters
  {0 {:font {:bold true} :border-top :medium}
   1 {:border-top :thin :border-bottom :thin}})


(defn old->new-tree [[title tree]]
  (let [branch? (complement (fn [x] (and (vector? x) (map? (second x)))))
        children #(when (vector? %) (second %))]
    (tree/tree branch? children tree first second)))


(defn ^:deprecated tree
  "Build a sheet grid from the provided tree of data
    [Tree Title [[Category Label [Children]] ... [Category Label [Children]]]]
  with leaves of the shape [Category Label {:column :value}].
  E.g. The assets section of a balance sheet might be represented by the tree
  [:balance-sheet
    [:assets
     [[:current-assets
       [[:cash {2018 100M, 2017 90M}]
        [:inventory {2018 1500M, 2017 1200M}]]]
      [:investments {2018 50M, 2017 45M}]]]]
  If provided, the formatters argument is a function that takes the integer
  depth of a category (increases with nesting) and returns a cell format for
  the row, and total-formatters is the same for rows that are totals."
  [core-table with-title t & {:keys [headers formatters total-formatters
                                     min-leaf-depth data-format]
                              :or {formatters default-tree-formatters
                                   total-formatters default-tree-total-formatters
                                   min-leaf-depth 2
                                   data-format :accounting}}]
  (let [title (first t)
        t (old->new-tree t)
        fmts (into (sorted-map) formatters)
        total-fmts (into (sorted-map) total-formatters)
        get' (fn [m k] (or (get m k) (val (last m))))]
    (with-title title
      (core-table
        (into [""] (remove #{""}) (or headers (keys (tree/fold + t))))
        (tree/table
          ;; Insert total rows below nodes with children
          (fn render [parent node depth]
            (if-not (tree/leaf? node)
              (let [combined (tree/fold + node)
                    empty-row (zipmap (keys combined) (repeat nil))]
                (concat
                  ; header
                  [(cell/style
                     (assoc empty-row "" (name parent))
                     (get' fmts depth))]
                  ; children
                  (tree/table render node)
                  ; total row
                  (when (> (count node) 1)
                    [(cell/style (assoc combined "" "") (get' total-fmts depth))])))
              ; leaf
              [(cell/style (assoc node "" (name parent))
                           (get' fmts (max min-leaf-depth depth)))]))
          t)))))
