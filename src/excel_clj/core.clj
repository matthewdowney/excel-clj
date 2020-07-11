(ns excel-clj.core
  "Utilities for declarative creation of Excel (.xlsx) spreadsheets,
  with higher level abstractions over Apache POI (https://poi.apache.org/).

  The highest level data abstraction used to create excel spreadsheets is a
  tree, followed by a table, and finally the most basic abstraction is a grid.

  The tree and table functions convert tree formatted or tabular data into a
  grid of [[cell]].

  Run the (example) function at the bottom of this namespace to see more."
  {:author "Matthew Downey"}
  (:require [excel-clj.cell :refer [style data dims wrapped?]]
            [excel-clj.file :as file]
            [excel-clj.tree :as tree]

            [clojure.string :as string]

            [taoensso.tufte :as tufte])
  (:import (clojure.lang Named)))


(set! *warn-on-reflection* true)


;;; Build grids of [[cell]] out of Clojure's data structures


(defn best-guess-cell-format
  "Try to guess appropriate formatting based on column name and cell value."
  [val column-name]
  (let [column' (string/lower-case
                  (if (instance? Named column-name)
                    (name column-name)
                    (str column-name)))]
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


(defn table
  "Build a lazy sheet grid from `rows`.

  Applies default styles to cells which are not already styled, but preserves
  any existing styles.

  Additionally, expands any rows which are wrapped with style data to apply the
  style to each cell of the row. See the comment block below this function
  definition for examples.

  This fn has the same shape as clojure.pprint/print-table."
  ([rows]
   (table (keys (data (first rows))) rows))
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
    {"My Table" (table (tdata 100)) ;; Write a table

     ;; Write a table that highlights rows where N has a whole square root
     "Highlight Table" (let [highlight {:fill-pattern :solid-foreground
                                        :fill-foreground-color :yellow}
                             square? (fn [n]
                                       (when (pos? n)
                                         (let [sqrt (Math/sqrt n)]
                                           (zero? (rem sqrt (int sqrt))))))]
                         (table
                           (for [row (tdata 100)]
                             (if (square? (row "N"))
                               (style row highlight)
                               row))))

     ;; Write a table with a merged top row
     "Titled Table" (cons
                      [(-> "My Big Title"
                           (dims {:width 3})
                           (style {:alignment :center}))]
                      (table (tdata 100)))}))


(defn- tree->rows [t]
  (let [total-fmts (sorted-map
                     0 {:font {:bold true} :border-top :medium}
                     1 {:border-top :thin :border-bottom :thin})
        fmts (sorted-map
               0 {:font {:bold true} :border-bottom :medium}
               1 {:font {:bold true}}
               2 {:indention 2}
               3 {:font {:italic true} :alignment :right})

        get' (fn [m k] (or (get m k) (val (last m))))]
    (tree/table
      ;; Insert total rows below nodes with children
      (fn render [parent node depth]
        (if-not (tree/leaf? node)
          (let [combined (tree/fold + node)
                empty-row (zipmap (keys combined) (repeat nil))]
            (concat
              ; header
              [(style (assoc empty-row "" (name parent)) (get' fmts depth))]
              ; children
              (tree/table render node)
              ; total row
              [(style (assoc combined "" "") (get' total-fmts depth))]))
          ; leaf
          [(style (assoc node "" (name parent)) (get' fmts depth))]))
      t)))


(defn tree
  [t]
  (let [ks (into [""] (keys (tree/fold + t)))]
    (table ks (tree->rows t))))


(defn with-title
  [title [row & _ :as rows]]
  (let [width (count row)]
    (cons
      [(dims title {:width width})]
      rows)))

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


(defn write-stream!
  "Like `write!`, but for a stream."
  ([workbook stream]
   (file/write-stream! workbook stream))
  ([workbook stream {:keys [streaming? auto-size-cols?]
                     :or   {streaming? true}
                     :as   ops}]
   (file/write-stream! workbook stream ops)))


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
       (tree tree/mock-balance-sheet)))

   "Tabular Sheet"
   (table
     [{"Date" "2018-01-01" "% Return" 0.05M "USD" 1500.5005M}
      {"Date" "2018-02-01" "% Return" 0.04M "USD" 1300.20M}
      {"Date" "2018-03-01" "% Return" 0.07M "USD" 2100.66666666M}])

   "Freeform Grid Sheet"
   [["First" "Second" (dims "Wide" {:width 2}) (dims "Wider" {:width 3})]
    ["First Column Value" "Second Column Value"]
    ["This" "Row" "Has" "Its" "Own" (style "Format" {:font {:bold true}})]]})


(defn example []
  (file/quick-open! example-workbook-data))
