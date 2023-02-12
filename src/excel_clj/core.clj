(ns excel-clj.core
  "Utilities for declarative creation of Excel (.xlsx) spreadsheets,
  with higher level abstractions over Apache POI (https://poi.apache.org/).

  The highest level data abstraction used to create excel spreadsheets is a
  tree, followed by a table, and finally the most basic abstraction is a grid.

  The tree and table functions convert tree formatted or tabular data into a
  grid of [[cell]].

  See the (comment) form with examples at the bottom of this namespace."
  {:author "Matthew Downey"}
  (:require [clojure.pprint :as pprint]
            [clojure.string :as string]

            [excel-clj.cell :refer [data dims style wrapped?]]
            [excel-clj.deprecated :as deprecated]
            [excel-clj.file :as file]
            [excel-clj.tree :as tree]

            [taoensso.tufte :as tufte])
  (:import (clojure.lang Named)
           (java.util Date)))


(set! *warn-on-reflection* true)


;;; Build grids of [[cell]] out of Clojure's data structures


(defn- name' [x]
  (if (instance? Named x)
    (name x)
    (str x)))


(defn best-guess-cell-format
  "Try to guess appropriate formatting based on column name and cell value."
  [val column-name]
  (let [column' (string/lower-case (name' column-name))]
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

(defn table-grid
  "Build a lazy sheet grid from `rows`.

  Applies default styles to cells which are not already styled, but preserves
  any existing styles.

  Additionally, expands any rows which are wrapped with style data to apply the
  style to each cell of the row. See the comment form below this function
  definition for examples.

  This fn has the same shape as clojure.pprint/print-table."
  ([rows]
   (table-grid (keys (data (first rows))) rows))
  ([ks rows]
   (assert (seq ks) "Columns are not empty.")
   (let [col-style {:border-bottom :thin :font {:bold true}}
         >row (fn [row-style row-data]
                (mapv
                  (fn [key]
                    (let [cell (get row-data key)]
                      (style
                        (if (wrapped? cell)
                          cell
                          (style cell (best-guess-cell-format cell key)))
                        row-style)))
                  ks))]
     (cons
       (mapv #(style (data %) col-style) ks)
       (for [row rows]
         (tufte/p :gen-row (>row (style row) (data row))))))))


(comment
  "Table examples"

  (defn tdata [n-rows]
    (for [i (range n-rows)]
      {"N" i
       "N^2" (* i i)
       "N as %" (/ i 100)}))

  (file/quick-open!
    {"My Table" (table-grid (tdata 100)) ;; Write a table

     ;; Write a table that highlights rows where N has a whole square root
     "Highlight Table" (let [highlight {:fill-pattern :solid-foreground
                                        :fill-foreground-color :yellow}
                             square? (fn [n]
                                       (when (pos? n)
                                         (let [sqrt (Math/sqrt n)]
                                           (zero? (rem sqrt (int sqrt))))))]
                         (table-grid
                           (for [row (tdata 100)]
                             (if (square? (row "N"))
                               (style row highlight)
                               row))))

     ;; Write a table with a merged top row
     "Titled Table" (cons
                      [(-> "My Big Title"
                           (dims {:width 3})
                           (style {:alignment :center}))]
                      (table-grid (tdata 100)))}))


(defn- tree->rows [t]
  (let [total-fmts (sorted-map
                     0 {:font {:bold true} :border-top :medium}
                     1 {:border-top :thin :border-bottom :thin})
        fmts (sorted-map
               0 {:font {:bold true} :border-bottom :medium}
               1 {:font {:bold true}}
               2 {:indention 2}
               3 {:font {:italic true} :alignment :right})
        num-format {:data-format :accounting}

        get' (fn [m k] (or (get m k) (val (last m))))
        style-data (fn [row style-map]
                     (let [label-key ""]
                       (->> row
                            (map (fn [[k v]]
                                   (if-not (= k label-key)
                                     [k (-> v
                                            (style num-format)
                                            (style style-map))]
                                     [k v])))
                            (into {}))))]
    (tree/table
      ;; Insert total rows below nodes with children
      (fn render [parent node depth]
        (if-not (tree/leaf? node)
          (let [combined (tree/fold + node)
                empty-row (zipmap (keys combined) (repeat nil))]
            (concat
              ; header
              [(style (assoc empty-row "" (name' parent)) (get' fmts depth))]
              ; children
              (tree/table render node)
              ; total row
              (when (> (count node) 1)
                [(style-data (assoc combined "" "") (get' total-fmts depth))])))
          ; leaf
          [(style-data (assoc node "" (name' parent)) (get' fmts (max depth 2)))]))
      t)))


(defn tree-grid
  "Build a lazy sheet grid from `tree`, whose leaves are shaped key->number.

  E.g. (tree-grid {:assets {:cash {:usd 100 :eur 100}}})

  See the comment form below this definition for examples."
  ([tree]
   (let [ks (into [""] (keys (tree/fold + tree)))]
     (tree-grid ks tree)))
  ([ks tree]
   (let [ks (into [""] (remove #{""}) ks)] ;; force the "" col to come first
     (table-grid ks (tree->rows tree)))))


(comment

  "Example: Trees using the 'tree' helper with default formatting."
  (let [assets {"Current" {:cash {:usd 100 :eur 100}
                           :inventory {:usd 500}}
                "Other" {:loans {:bank {:usd 500}
                                 :margin {:usd 1000 :eur 30000}}}}
        liabilities {"Current" {:accounts-payable {:usd 50 :eur 0}}}]
    (file/quick-open!
      {"Just Assets"
       (tree-grid {"Assets" assets})

       "Both in One Tree"
       (tree-grid
         {"Accounts"
          {"Assets" assets
           ;; Because they're in one tree, assets will sum with liabilities,
           ;; so we should invert the sign on the liabilities to get a
           ;; meaningful sum
           "Liabilities" (tree/negate liabilities)}})

       "Both in Two Trees"
       (let [diff (tree/fold
                    - {:assets-sum (tree/fold + assets)
                       :liabilities-sum (tree/fold - liabilities)})
             no-header rest]
         (concat
           (tree-grid {"Assets" assets})
           [[""]]
           (no-header (tree-grid {"Liabilities" liabilities}))
           [[""]]
           (no-header (tree-grid {"Assets Less Liabilities" diff}))))}))

  "Example: Trees using `excel-clj.tree/table` and then using the `table`
  helper."
  (let [table-data
        (->> (tree/table tree/mock-balance-sheet)
             (map
               (fn [row]
                 (let [spaces (apply str (repeat (:tree/indent row) "  "))]
                   (-> row
                       (dissoc :tree/indent)
                       (update "" #(str spaces %)))))))]
    (file/quick-open! {"Defaults" (table-grid ["" 2018 2017] table-data)})))


;;; Helpers to manipulate [[cell]] data structures


(defn with-title
  "Prepend a centered `title` row to the `grid` with the same width as the
  first row of the grid."
  [title [row & _ :as grid]]
  (let [width (count row)]
    (cons
      [(-> title (dims {:width width}) (style {:alignment :center}))]
      grid)))


(defn transpose
  "Transpose a grid."
  [grid]
  (apply mapv vector grid))


(defn juxtapose
  "Put grids side by side (whereas `concat` works vertically, this works
  horizontally).

  Optionally, supply some number of blank `padding` columns between the two
  grids.

  Finds the maximum row width in the left-most grid and pads all of its rows
  to that length before sticking them together."
  ([left-grid right-grid]
   (juxtapose left-grid right-grid 0))
  ([left-grid right-grid padding]
   (let [;; First pad the height of both grids
         height (max (count left-grid) (count right-grid))
         empty-row []
         pad-height (fn [xs]
                      (concat xs (repeat (- height (count xs)) empty-row)))

         ;; Then pad the width of the left grid so that it's uniform
         row-width (fn [row] (apply + (map (comp :width dims) row)))
         max-row-width (apply max (map row-width left-grid))
         pad-to (fn [width row]
                  (let [cells-needed (- width (row-width row))]
                    (into row (repeat cells-needed ""))))
         padded-left-grid (map
                            (partial pad-to (+ max-row-width padding))
                            (pad-height left-grid))]
     (map into padded-left-grid (pad-height right-grid)))))


(comment
  "Example: juxtaposing two grids with different widths and heights"
  (let [squares (-> (table-grid (for [i (range 10)] {"X" i "X^2" (* i i)}))
                    (vec)
                    (update 5 into [(dims "<- This one is 4^2" {:width 2})])
                    (update 6 into ["^ Juxt should make room for that cell"]))
        cubes (table-grid (for [i (range 20)] {"X" i "X^3" (* i i i)}))]
    (file/quick-open!
      {"Juxtapose" (juxtapose squares cubes)}))

  "Example: A multiplication table"
  (let [highlight {:fill-pattern :solid-foreground
                   :fill-foreground-color :yellow}

        grid (for [x (range 1 11)]
               (for [y (range 1 11)]
                 (cond-> (* x y) (= x y) (style highlight))))

        cols (map #(style % {:font {:bold true}}) (range 1 11))

        grid (concat [cols] grid)
        grid (juxtapose (transpose [(cons nil cols)]) grid)]
    (file/quick-open!
      {"Transpose & Juxtapose"
       (with-title "Multiplication Table" grid)})))


;;; File interaction


(defn write!
  "Write the `workbook` to the given `path` and return a file object pointing
  at the written file.

  The workbook is a key value collection of (sheet-name grid), either as map or
  an association list (if ordering is important)."
  ([workbook path] (file/write! workbook path))
  ([workbook path {:keys [streaming? auto-size-cols?]
                   :or   {streaming? true}
                   :as   ops}]
   (file/write! workbook path ops)))


(defn append!
  "Merge the `workbook` with the one saved at `from-path`, write it to the
  given `path`, and return a file object pointing at the written file.

  The workbook is a key value collection of (sheet-name grid), either as map or
  an association list (if ordering is important).

  The 'merge' logic overwrites sheets of the same name in the workbook at
  `from-path`, so this function is only capable of appending sheets to a
  workbook, not appending cells to a sheet."
  ([workbook from-path path] (file/append! workbook from-path path))
  ([workbook from-path path {:keys [streaming? auto-size-cols?]
                             :or   {streaming? true}
                             :as   ops}]
   (file/append! workbook from-path path ops)))


(defn write-stream!
  "Like `write!`, but for a stream."
  ([workbook stream]
   (file/write-stream! workbook stream))
  ([workbook stream {:keys [streaming? auto-size-cols?]
                     :or   {streaming? true}
                     :as   ops}]
   (file/write-stream! workbook stream ops)))


(defn append-stream!
  "Like `append!`, but for streams."
  ([workbook from-stream stream]
   (file/append-stream! workbook from-stream stream))
  ([workbook from-stream stream {:keys [streaming? auto-size-cols?]
                                 :or   {streaming? true}
                                 :as   ops}]
   (file/append-stream! workbook from-stream stream ops)))


(defn write-pdf!
  "Write the workbook to the given filename and return a file object pointing
  at the written file.

  Requires OpenOffice. See https://github.com/sbraconnier/jodconverter.

  The workbook is a key value collection of (sheet-name grid), either as map or
  an association list (if ordering is important)."
  [workbook path]
  (file/write-pdf! workbook path))


(defn quick-open!
  "Write a workbook to a temp file & open it. Useful for quick repl viewing."
  [workbook]
  (file/quick-open! workbook))


(defn quick-open-pdf!
  "Write a workbook to a temp file as a pdf & open it. Useful for quick repl
  viewing."
  [workbook]
  (file/quick-open-pdf! workbook))


;; Convenience macro to redirect print-table / print-tree to excel


(defonce ^:private var->excel-rebinding (atom {}))


(defn declare-excelable!
  "Redefine some function's `var` to generate Excel output when enclosed in an
  `excel` macro.

  The `fn` returns a grid (optionally with :excel/sheet-name metadata)."
  [var fn]
  (swap! var->excel-rebinding assoc var fn))


(declare-excelable! #'pprint/print-table
  (fn ;; This fn has the same signature as the var it's redefining
    ([rows] (vary-meta (table-grid rows) merge (meta rows)))
    ([ks rows] (vary-meta (table-grid ks rows) merge (meta rows)))))


(declare-excelable! #'tree/print-table
  (fn this
    ([rows]
     (this
       (into [""] (remove #{"" :tree/indent}) (keys (data (first rows))))
       rows))
    ([ks rows]
     (vary-meta
       (table-grid ks
         (map
           (fn [{:keys [tree/indent] :as row}]
             (update row "" #(str (apply str (repeat (or indent 0) " ")) %)))
           rows))
       merge (meta rows)))))


(defn -build-excel-rebindings [wb-atom var->excel-rebinding]
  (letfn [(conj-page [sheets contents]
            (let [sheet-name (or (:excel/sheet-name (meta contents))
                               (str "Sheet" (inc (count sheets))))]
              (conj sheets [sheet-name contents])))
          (conj-page! [contents] (swap! wb-atom conj-page contents))]
    (update-vals var->excel-rebinding
      (fn [grid-fn] (comp conj-page! grid-fn)))))


(defmacro excel
  "Build an Excel workbook with whatever data is emitted during the execution
  of `body` from functions on which `declare-excelable!` has been called.

  If the first argument is a compile-time map, it may contain a :hook function
  to be called with the final workbook. If no hook is passed, it defaults to
  `quick-open!`.

  (Compatible by default for `clojure.pprint/print-table` and
  `excel-clj.tree/print-table`.)

  Returns the return value of `body`."
  [& body]
  (let [[opts body] (if (map? (first body))
                      [(first body) (rest body)]
                      [{} body])
        hook (or (:hook opts) quick-open!)]
    `(let [wb# (atom [])]
       (with-redefs-fn (-build-excel-rebindings wb# ~(deref var->excel-rebinding))
         (fn []
           (let [ret# (do ~@body)]
             (~hook (apply array-map (mapcat identity @wb#)))
             ret#))))))


(comment
  ;; For example
  (excel
    (do
      ;; Print a table to one sheet
      (pprint/print-table (map (fn [i] {"Ch" (char i) "i" i}) (range 33 43)))
      ;; And a tree to another
      (let [tbl (tree/table (tree/combined-header) tree/mock-balance-sheet)]
        (tree/print-table tbl))
      :ok)))


;; Some v1.X backwards compatibility


(def ^:deprecated tree
  "Deprecated in favor of `tree-grid`."
  (partial deprecated/tree table-grid with-title))


(def ^:deprecated table
  "Deprecated in favor of `table-grid`."
  deprecated/table)


(def ^:deprecated quick-open
  "Deprecated in favor of `quick-open!`."
  quick-open!)


(comment
  "Example: Using deprecated `tree` and `table` functions"
  (quick-open!
    {"tree"  (tree
               ["Mock Balance Sheet for the year ending Dec 31st, 2018"
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
                   [["Common Stock" {2018 102M, 2017 80M}]]]]]])
     "table" (table (for [n (range 100)] {"X" n "X^2" (* n n)}))}))


;;; Performance tests for order-of-magnitude checks


(comment

  (defmacro time' [& body]
    `(let [start# (System/currentTimeMillis)]
       (do ~@body)
       [(- (System/currentTimeMillis) start#) :ms]))

  (defn example-table [n-rows]
    (for [i (range n-rows)]
      {"N" i
       "N^2" (* i i)
       "N as %" (/ i 100)}))

  (defn do-test
    ([n-rows]
     (do-test n-rows nil))
    ([n-rows ops]
     (let [n (long n-rows)]
       (println "Writing" n "rows...")
       {n (time'
            (if ops
              (file/write! {"Sheet 1" (example-table n)} "test.xlsx" ops)
              (file/write! {"Sheet 1" (example-table n)} "test.xlsx")))})))

  ;;; (1) Performance with auto-sizing of columns

  (let [ops {:auto-size-cols? true}]
    (->> [1e2 1e3 1e4 1e5]
         (map #(do-test % ops))
         (apply merge)))
  ;=> {100    [88 :ms]
  ;    1000   [106 :ms]
  ;    10000  [830 :ms]
  ;    100000 [8036 :ms]}

  ;;; (2) Performance WITHOUT auto-sizing of columns

  (let [ops {:auto-size-cols? false}]
    (->> [1e2 1e3 1e4 1e5]
         (map #(do-test % ops))
         (apply merge)))
  ;=> {100    [30 :ms]
  ;    1000   [41 :ms]
  ;    10000  [183 :ms]
  ;    100000 [1290 :ms]}

  (tufte/add-basic-println-handler! {})
  (tufte/profile {} (do-test 100000 {:auto-size-cols? false}))

  ;; Hence by default, we turn off auto-sizing after 10,000 rows

  ;;; (3) Performance with default settings

  (->> [1e2 1e3 1e4 1e5]
       (map do-test)
       (apply merge))
  ;=> {100    [74 :ms]
  ;    1000   [178 :ms]
  ;    10000  [145 :ms]
  ;    100000 [1249 :ms]}
  )


;;; Final examples


(def example-workbook-data
  {"Tree Sheet"
   (let [title "Mock Balance Sheet Ending Dec 31st, 2020"]
     (with-title (style title {:alignment :center})
                 (tree-grid tree/mock-balance-sheet)))

   "Tabular Sheet"
   (table-grid
     [{"Date" "2018-01-01" "% Return" 0.05M "USD" 1500.5005M}
      {"Date" "2018-02-01" "% Return" 0.04M "USD" 1300.20M}
      {"Date" "2018-03-01" "% Return" 0.07M "USD" 2100.66666666M}])

   "Freeform Grid Sheet"
   [["First" "Second" (dims "Wide" {:width 2}) (dims "Wider" {:width 3})]
    ["First Column Value" "Second Column Value"]
    ["This" "Row" "Has" "Its" "Own" (style "Format" {:font {:bold true}})]]})


(defn example []
  (quick-open! example-workbook-data))


(def example-template-data
  ;; Some mocked tabular uptime data to inject into the template
  (let [start-ts (inst-ms #inst"2020-05-01")
        one-hour (* 1000 60 60)]
    (for [i (range 99)]
      {"Date"                 (Date. ^long (+ start-ts (* i one-hour)))
       "Webserver Uptime"     (- 1.0 (rand 0.25))
       "REST API Uptime"      (- 1.0 (rand 0.25))
       "WebSocket API Uptime" (- 1.0 (rand 0.25))})))


(comment
  "Example: Creating a workbooks different kinds of worksheets"
  (example)

  "Example: Creating a workbook by filling in a template.

  The template here has a 'raw' sheet, which contains uptime data for 3 time
  series, and a 'Summary' sheet, wich uses formulas + the raw data to compute
  and plot. We're going to overwrite the 'raw' sheet to fill in the template."
  (let [template (clojure.java.io/resource "uptime-template.xlsx")
        new-data {"raw" (table-grid example-template-data)}]
    (file/open (append! new-data template "filled-in-template.xlsx"))))
