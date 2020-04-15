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
            [excel-clj.prototype :as pt]
            [clojure.string :as string]
            [clojure.java.io :as io])
  (:import (java.io File)
           (java.awt Desktop HeadlessException)
           (org.jodconverter.office DefaultOfficeManagerBuilder)
           (org.jodconverter OfficeDocumentConverter)))

(set! *warn-on-reflection* true)

(def ^{:dynamic true :deprecated true} *max-col-width*
  "Deprecated -- no longer has any effect."
  15000)

(def ^{:dynamic true :deprecated true} *n-threads*
  "Deprecated -- no longer has any effect."
  (+ 2 (.. Runtime getRuntime availableProcessors)))

;;; Higher-level code to specify grids in terms of clojure data structures,
;;; organized as either a table or a tree

(defn table
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
                    (let [style (style/merge-all
                                  (or (data-style row col-name) {})
                                  (style/best-guess-row-format row col-name))]
                      (when (or (= (:data-format style) :accounting)
                                (number? (get row col-name "")))
                        (vswap! numeric? conj col-name))
                      {:value (get row col-name)
                       :style style}))
        getters (map (fn [col-name] #(data-cell col-name %)) headers)
        header-style (or header-style
                         ;; Add right alignment if it's an accounting column
                         (fn [name]
                           (cond-> (style/default-header-style name)
                                   (@numeric? name)
                                   (assoc :alignment :right))))]
    (cons
      (map (fn [x] {:value x :style (header-style x)}) headers)
      (map (apply juxt getters) tabular-data))))

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
        [[{:value (first t) :style {:alignment :center}
           :width (inc (count all-colls))}]]

        ;; Headers
        [(into [""] (map #(->{:value % :style header-style})) all-colls)]

        ;; Line items
        (for [line tabular]
          (let [total? (::tree/total? line)
                format (or
                         (fmt-or-max
                           (if total? total-formatters formatters)
                           (::tree/depth line))
                         {})
                style (style/merge-all format {:data-format data-format})]
            (into [{:value (::tree/label line) :style (if total? {} style)}]
                  (map #(->{:value (get line %) :style style})) all-colls)))))
    (catch Exception e
      (throw (ex-info "Failed to render tree" {:tree t} e)))))

(defn with-title
  "Write a title above the given grid with a width equal to the widest row."
  [grid title]
  (let [width (transduce (map count) (completing max) 0M grid)]
    (concat
      [[{:value title :width width :style {:alignment :center}}]]
      grid)))

;;; Utilities to write & open workbooks as XLSX or PDF files

(defn- force-extension [path ext]
  (let [path (.getCanonicalPath (io/file path))]
    (if (.endsWith path ext)
      path
      (let [parts (string/split path (re-pattern File/separator))]
        (str
          (string/join
            File/separator (if (> (count parts) 1) (butlast parts) parts))
          "." ext)))))

(defn- temp
  "Return a (string) path to a temp file with the given extension."
  [ext]
  (-> (File/createTempFile "generated-sheet" ext) .getCanonicalPath))

(defn write!
  "Write the workbook to the given filename and return a file object pointing
  at the written file.

  The workbook is a key value collection of (sheet-name grid), either as map or
  an association list (if ordering is important)."
  [workbook path]
  (let [convert-cell (fn [{:keys [value style width height]
                           :or   {width 1 height 1}
                           :as   cell-data}]
                       (if-not (map? cell-data)
                         (pt/wrapped cell-data)
                         (-> (pt/wrapped value)
                             (pt/style style)
                             (pt/dims {:width width :height height}))))
        convert-row (fn [row] (map convert-cell row))]
    (pt/write!
      (map (fn [[sheet grid]] [sheet (map convert-row grid)]) workbook)
      path)))

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
     [["First" "Second" {:value "Wide" :width 2} {:value "Wider" :width 3}]
      ["First Column Value" "Second Column Value"]
      ["This" "Row" "Has" "Its" "Own"
       {:value "Format" :style {:font {:bold true}}}]]})

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
                        (for [x (range 10000)]
                          {"N" x, "N^2" (* x x), "N^3" (* x x x)}))})

  ;; Ballpark performance test
  (dotimes [_ 5]
    (time
      (write!
        [["Test"
          (table
            (for [x (range 100000)]
              {"N" x "N^2" (* x x) "N^3" (* x x x)}))]]
        "test.xlsx")))

  )
