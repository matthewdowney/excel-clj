(ns excel-clj.core
  "Utilities for declarative creation of Excel (.xlsx) spreadsheets,
  with higher level abstractions over Apache POI (https://poi.apache.org/).

  The highest level data abstraction used to create excel spreadsheets is a
  tree, followed by a table, and finally the most basic abstraction is a grid.

  The tree and table functions convert tree formatted or tabular data into a
  grid of [[cell]].

  Run the (example) function at the bottom of this namespace to see more."
  {:author "Matthew Downey"}
  (:require [excel-clj.tree :as tree]
            [excel-clj.style :as style]
            [clojure.string :as string]
            [clojure.java.io :as io]
            [taoensso.encore :as enc]
            [excel-clj.poi :as poi]
            [taoensso.tufte :as tufte])
  (:import (java.io File)
           (java.awt Desktop HeadlessException)
           (org.jodconverter.office DefaultOfficeManagerBuilder)
           (org.jodconverter OfficeDocumentConverter)
           (org.apache.poi.ss.usermodel Sheet)
           (org.apache.poi.xssf.streaming SXSSFSheet)))


(set! *warn-on-reflection* true)


;;; Model an Excel document 'cell'.

;;; A cell can be either a plain value (a string, java.util.Date, etc.) or such
;;; a value wrapped in a map which also includes style and dimension data.


(defn wrapped? [x] (::wrapped? x))
(defn wrapped
  "If `x` contains cell data wrapped in a map (with style & dimension data),
  return it as-is. Otherwise return a wrapped version."
  [x]
  (if (wrapped? x)
    x
    {::wrapped? true ::data x}))


(defn style
  "Get the style specification for `x`, or deep-merge its current style spec
  with the given `style-map`."
  ([x]
   (or (::style x) {}))
  ([x style-map]
   (let [style-map (enc/nested-merge (style x) style-map)]
     (assoc (wrapped x) ::style style-map))))


(defn dims
  "Get the {:width N, :height N} dimension map for `x`, or merge in the given
  `dims-map` of the same format."
  ([x]
   (or (::dims x) {:width 1 :height 1}))
  ([x dims-map]
   (let [dims-map (merge (dims x) dims-map)]
     (assoc (wrapped x) ::dims dims-map))))


(defn data
  "If `x` contains cell data wrapped in a map (with style & dimension data),
  return the wrapped cell value. Otherwise return as-is."
  [x]
  (if (wrapped? x)
    (::data x)
    x))


;;; Interface: Writing and opening Excel worksheets & PDFs


(defn- write-rows!
  "Write the rows via the poi/SheetWriter `sh`, returning the max row width."
  [sh rows-seq]
  (reduce
    (fn [n next-row]
      (let [width
            (count
              (for [cell next-row]
                (let [{:keys [width height]} (dims cell)]
                  (poi/write! sh (data cell) (style cell) width height))))]
        (poi/newline! sh)
        (max n width)))
    0
    rows-seq))


(defn- write [workbook poi-writer {:keys [streaming? auto-size-cols?] :as ops}]
  (doseq [[nm rows] workbook
          :let [sh (poi/sheet-writer poi-writer nm)
                auto-size? (or (true? auto-size-cols?)
                               (get auto-size-cols? nm))]]

    (when (and streaming? auto-size?)
      (.trackAllColumnsForAutoSizing ^SXSSFSheet (:sheet sh)))

    (let [n-cols (write-rows! sh rows)]
      (when auto-size?
        (dotimes [i n-cols]
          (.autoSizeColumn ^Sheet (:sheet sh) i))))))


(defn default-ops
  "Decide if sheet columns should be autosized by default based on how many
  rows there are.

  This check is careful to preserve the laziness of grids as much as possible."
  [workbook]
  (reduce
    (fn [ops [sheet-name sheet-grid]]
      (if (>= (bounded-count 10000 sheet-grid) 10000)
        (assoc-in ops [:auto-size-cols? sheet-name] false)
        (assoc-in ops [:auto-size-cols? sheet-name] true)))
    {:streaming? true :auto-size-cols? {}}
    workbook))


(defn force-extension [path ext]
  (let [path (.getCanonicalPath (io/file path))]
    (if (string/ends-with? path ext)
      path
      (let [sep (re-pattern (string/re-quote-replacement File/separator))
            parts (string/split path sep)]
        (str
          (string/join
            File/separator (if (> (count parts) 1) (butlast parts) parts))
          "." ext)))))


(defn write!
  "Write the `workbook` to the given `path` and return a file object pointing
  at the written file.

  The workbook is a key value collection of (sheet-name grid), either as map or
  an association list (if ordering is important)."
  ([workbook path]
   (write! workbook path (default-ops workbook)))
  ([workbook path {:keys [streaming? auto-size-cols?]
                   :or   {streaming? true}
                   :as   ops}]
   (let [f (io/file (force-extension (str path) ".xlsx"))]
     (with-open [w (poi/writer f streaming?)]
       (write workbook w (assoc ops :streaming? streaming?)))
     f)))


(defn write-stream!
  "Like `write!`, but for a stream."
  ([workbook stream]
   (write-stream! workbook stream (default-ops workbook)))
  ([workbook stream {:keys [streaming? auto-size-cols?]
                     :or   {streaming? true}
                     :as   ops}]
   (with-open [w (poi/stream-writer stream)]
     (write workbook w (assoc ops :streaming? streaming?)))))


(defn temp
  "Return a (string) path to a temp file with the given extension."
  [ext]
  (-> (File/createTempFile "generated-sheet" ext) .getCanonicalPath))


(defn convert-pdf!
  "Convert the `from-document`, either a File or a path to any office document,
  to pdf format and write the pdf to the given pdf-path.

  Requires OpenOffice. See https://github.com/sbraconnier/jodconverter.

  Returns a File pointing at the PDF."
  [from-document pdf-path]
  (let [path (force-extension pdf-path "pdf")
        office-manager (.build (DefaultOfficeManagerBuilder.))]
    (.start office-manager)
    (try
      (let [document-converter (OfficeDocumentConverter. office-manager)]
        (.convert document-converter (io/file from-document) (io/file path)))
      (finally
        (.stop office-manager)))
    (io/file path)))


(defn write-pdf!
  "Write the workbook to the given filename and return a file object pointing
  at the written file.

  Requires OpenOffice. See https://github.com/sbraconnier/jodconverter.

  The workbook is a key value collection of (sheet-name grid), either as map or
  an association list (if ordering is important)."
  [workbook path]
  (let [temp-path (temp ".xlsx")
        pdf-file (convert-pdf! (write! workbook temp-path) path)]
    (.delete (io/file temp-path))
    pdf-file))


(defn open
  "Open the given file path with the default program."
  [file-path]
  (try
    (let [f (io/file file-path)]
      (.open (Desktop/getDesktop) f)
      f)
    (catch HeadlessException e
      (throw (ex-info "There's no desktop." {:opening file-path} e)))))


(defn quick-open
  "Write a workbook to a temp file & open it. Useful for quick repl viewing."
  [workbook]
  (open (write! workbook (temp ".xlsx"))))


(defn quick-open-pdf
  "Write a workbook to a temp file as a pdf & open it. Useful for quick repl
  viewing."
  [workbook]
  (open (write-pdf! workbook (temp ".pdf"))))


(comment
  "Grid examples"

  ;;; Writing a simple grid
  (let [grid [["A" "B" "C"]
              [1 2 3]]]
    (quick-open {"Sheet 1" grid}))


  ;;; Writing a grid with a bit more formatting
  (let [t (java.util.Calendar/getInstance)
        grid [["String" "Abc"]
              ["Numbers" 100M 1.234 1234 12345N]
              ["Date (not styled, styled)" t (style t {:data-format :ymd})]]

        header-style {:border-bottom :thin :font {:bold true}}
        header-rows [[(-> "Type"
                          (style header-style)
                          (dims {:height 2})
                          (style {:vertical-alignment :center}))
                      (-> "Examples"
                          (style header-style)
                          (dims {:width 4})
                          (style {:alignment :center :border-bottom :none}))]
                     (mapv #(style % {:font {:italic true}
                                      :alignment :center
                                      :border-bottom :thin})
                           [nil 1 2 3 4])]
        excel-file (quick-open {"Sheet 1" (concat header-rows grid)})]

    (try
      (open (convert-pdf! excel-file (temp ".pdf")))
      (catch Exception e
        (println "(Couldn't open a PDF on this platform.)")))))


;;; Main interface: build Excel worksheets out of Clojure's data structures.


(defn best-guess-cell-format
  "Try to guess appropriate formatting based on column name and cell value."
  [val column-name]
  (let [column' (string/lower-case (name column-name))]
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
  any existing styles. Additionally, expands any rows which are wrapped with
  style data to apply the style to each cell of the row.

  See the comment block below this function definition for examples.

  This fn has the same shape as clojure.pprint/print-table."
  ([rows]
   (table (keys (first rows)) rows))
  ([ks rows]
   (assert (seq ks) "Columns are not empty.")
   (let [col-style {:border-bottom :thin :font {:bold true}}]
     (cons
       (mapv #(style (data %) col-style) ks)
       (for [row rows]
         (tufte/p :gen-row
           (let [row-style (style row)
                 row (data row)]
             (mapv
               (fn [key]
                 (let [cell (get row key)]
                   (style
                     (if (wrapped? cell)
                       cell
                       (style cell (best-guess-cell-format cell key)))
                     row-style)))
               ks))))))))


(comment
  "Table examples"

  (defn tdata [n-rows]
    (for [i (range n-rows)]
      {"N" i
       "N^2" (* i i)
       "N as %" (/ i 100)}))

  (quick-open
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


;; TODO: (defn tree [])


(defn tree
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
    [t & {:keys [headers formatters total-formatters min-leaf-depth data-format]
          :or {formatters style/default-tree-formatters
               total-formatters style/default-tree-total-formatters
               min-leaf-depth 2
               data-format :accounting}}]
    (try
      (let [tabular (tree/accounting-table (second t) :min-leaf-depth min-leaf-depth)
            fmt-or-max (fn [fs n]
                         (or (get fs n) (second (apply max-key first fs))))
            all-colls (or headers
                          (sequence
                            (comp
                              (mapcat keys)
                              (filter (complement qualified-keyword?))
                              (distinct))
                            tabular))
            header-style {:font {:bold true} :alignment :right}]
        (concat
          ;; Title
          [[(-> (first t)
                (style {:alignment :center})
                (dims {:width (inc (count all-colls))}))]]

          ;; Headers
          [(into [""] (map #(style % header-style)) all-colls)]

          ;; Line items
          (for [line tabular]
            (let [total? (::tree/total? line)
                  style' (or
                           (fmt-or-max
                             (if total? total-formatters formatters)
                             (::tree/depth line))
                           {})
                  style' (enc/nested-merge style' {:data-format data-format})]
              (into [(style (::tree/label line) (if total? {} style'))]
                    (map #(style (get line %) style') all-colls))))))
      (catch Exception e
        (throw (ex-info "Failed to render tree" {:tree t} e)))))

;;; Performance tests for order-of-magnitude checks


(comment

  (defmacro time' [& body]
    `(let [start# (System/currentTimeMillis)]
       (do ~@body)
       [(- (System/currentTimeMillis) start#) :ms]))

  (defn do-test
    ([n-rows]
     (do-test n-rows nil))
    ([n-rows ops]
     (let [n (long n-rows)]
       (println "Writing" n "rows...")
       {n (time'
            (if ops
              (write! {"Sheet 1" (example-table n)} "test.xlsx" ops)
              (write! {"Sheet 1" (example-table n)} "test.xlsx")))})))

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


;;; Examples (note that others exist in comment blocks throughout the ns)


(def example-workbook-data
  {"Tree Sheet"
   (tree
     ["Mock Balance Sheet for the year ending Dec 31st, 2018"
      tree/mock-balance-sheet])

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
  (quick-open example-workbook-data))


(comment
  ;; This should open an Excel workbook
  (example)

  ;; This will both open an example excel sheet and write & open a test pdf file
  ;; with the same contents. On platforms without OpenOffice the convert-pdf!
  ;; call will most likely fail.
  (open (convert-pdf! (example) (temp ".pdf")))

  ;; Expose ordering / styling issues in v1.2.X
  (quick-open {"Test" (table
                        (for [x (range 20000)]
                          {"N" x, "N^2" (* x x), "N^3" (* x x x)}))})

  ;; Ballpark performance test
  (dotimes [_ 5]
    (time (write! [["Test" (table
                             (for [x (range 100000)]
                               {"N" x "N^2" (* x x) "N^3" (* x x x)}))]]
                  "test.xlsx")))
  ; "Elapsed time: 1238.133651 msecs"
  ; "Elapsed time: 1190.656899 msecs"
  ; "Elapsed time: 1195.8068 msecs"
  ; "Elapsed time: 1182.257177 msecs"
  ; "Elapsed time: 1166.420738 msecs"
  ;=> nil
  )
