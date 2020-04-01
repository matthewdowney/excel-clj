(ns
  ^{:doc "Utilities for declarative creation of Excel (.xlsx) spreadsheets,
  with higher level abstractions over Apache POI (https://poi.apache.org/).

  The highest level data abstraction used to create excel spreadsheets is a
  tree, followed by a table, and finally the most basic abstraction is a grid.

  The tree and table functions convert tree formatted or tabular data into a
  grid of [[cell]].

  Run the (example) function at the bottom of this namespace to see more."
    :author "Matthew Downey"} excel-clj.core
  (:require [excel-clj.tree :as tree]
            [excel-clj.style :as style]
            [clojure.string :as string]
            [clojure.java.io :as io]
            [taoensso.tufte :as tufte :refer (defnp p profiled profile)])
  (:import (org.apache.poi.ss.usermodel Cell RichTextString)
           (org.apache.poi.xssf.usermodel XSSFWorkbook XSSFSheet XSSFRow XSSFCell)
           (java.io File)
           (java.awt Desktop HeadlessException)
           (java.util Calendar Date)
           (org.apache.poi.ss.util CellRangeAddress)
           (org.jodconverter.office DefaultOfficeManagerBuilder)
           (org.jodconverter OfficeDocumentConverter)))

(set! *warn-on-reflection* true)

;;; Low level code to write to & style sheets; you probably shouldn't have to
;;; touch this to make use of the API, but might choose to when adding or
;;; extending functionality

(defmacro ^:private if-type
  "For situations where there are overloads of a Java method that accept
  multiple types and you want to either call the method with a correct type
  hint (avoiding reflection) or do something else.

  In the `if-true` form, the given `sym` becomes type hinted with the type in
  `types` where (instance? type sym). Otherwise the `if-false` form is run."
  [[sym types] if-true if-false]
  (let [typed-sym (gensym)]
    (letfn [(with-hint [type]
              (let [using-hinted
                    ;; Replace uses of the un-hinted symbol if-true form with
                    ;; the generated symbol, to which we're about to add a hint
                    (clojure.walk/postwalk-replace {sym typed-sym} if-true)]
                ;; Let the generated sym with a hint, e.g. (let [^Float x ...])
                `(let [~(with-meta typed-sym {:tag type}) ~sym]
                   ~using-hinted)))
            (condition [type] (list `(instance? ~type ~sym) (with-hint type)))]
      `(cond
         ~@(mapcat condition types)
         :else ~if-false))))

;; Example of the use of if-type
(comment
  (let [test-fn #(time (reduce + (map % (repeat 1000000 "asdf"))))
        reflection (fn [x] (.length x))
        len-hinted (fn [^String x] (.length x))
        if-type' (fn [x] (if-type [x [String]]
                           (.length x)
                           ;; So we know it executes the if-true path
                           (throw (RuntimeException.))))]
    (println "Running...")
    (print "With manual type hinting =>" (with-out-str (test-fn len-hinted)))
    (print "With if-type hinting     =>" (with-out-str (test-fn if-type')))
    (print "With reflection          => ")
    (flush)
    (print (with-out-str (test-fn reflection)))))

(defn- write-cell!
  "Write the given data to the mutable cell object, coercing its type if
  necessary."
  [^Cell cell data]
  ;; These types are allowed natively
  (if-type [data [Boolean Calendar String Date Double RichTextString]]
           (doto cell (.setCellValue data))

           ;; Apache POI requires that numbers be doubles
           (if (number? data)
             (doto cell (.setCellValue (double data)))

             ;; Otherwise stringify it
             (let [to-write (or (some-> data pr-str) "")]
               (doto cell (.setCellValue ^String to-write))))))

(def ^:dynamic *max-col-width*
  "Sometimes POI's auto sizing isn't super intelligent, so set a sanity-max on
  the column width."
  15000)

(def ^:dynamic *n-threads*
  "Allow a custom number of threads used during writing."
  (+ 2 (.. Runtime getRuntime availableProcessors)))

(defmacro ^:private doparallel [[sym coll] & body]
  "Performance hack for writing the POI cells.
  Like (dotimes [x xs] ...) but parallel."
  `(let [n# *n-threads*
         equal-chunks# (loop [num# n#, parts# [], coll# ~coll, c# (count ~coll)]
                         (if (<= num# 0)
                           parts#
                           (let [t# (quot (+ c# num# -1) num#)]
                             (recur (dec num#) (conj parts# (take t# coll#))
                                    (drop t# coll#) (- c# t#)))))
         workers#
         (doall
           (for [chunk# equal-chunks#]
             (future
               (doseq [~sym chunk#]
                 ~@body))))]
     (doseq [w# workers#]
       (deref w#))))

(defn- ^XSSFSheet write-grid!
  "Modify the given workbook by adding a sheet with the given name built from
  the provided grid.

  The grid is a collection of rows, where each cell is either a plain, non-map
  value or a map of {:value ..., :style ..., :width ...}, with :value being the
  contents of the cell, :style being an optional map of style data, and :width
  being an optional cell width dictating how many horizontal slots the cell
  takes up (creates merged cells).

  Returns the sheet object."
  [^XSSFWorkbook workbook ^String sheet-name grid]
  (let [^XSSFSheet sh (.createSheet workbook sheet-name)
        build-style' (memoize ;; Immutable styles can share mutable objects :)
                       (fn [style-map]
                         (->> (style/merge-all style/default-style (or style-map {}))
                              (style/build-style workbook))))
        layout (volatile! {})]
    (try

      ;; N.B. So this code got uglier due to performance. Writing the cells
      ;; takes many seconds for a large sheet (~50,000 rows) and we can improve
      ;; the process a bit by doing the cell creation sequentially and the cell
      ;; writing in parallel (on test data set reduced from ~19s to ~14s).

      ;; Unfortunately much of the time is spent writing to disk (~8s).

      ;; We have to do this part sequentially because POI doesn't use
      ;; thread-safe data structures
      (doseq [[row-idx row-data] (map-indexed vector grid)]
        (let [row (p :create-row (.createRow sh (int row-idx)))]
          (loop [col-idx 0 cells row-data]
            (when-let [cell-data (first cells)]
              ;; (1) Build the cell
              (let [cell (p :create-cell (.createCell ^XSSFRow row col-idx))
                    width (if (map? cell-data) (get cell-data :width 1) 1)]

                ;; (2) Merge if necessary into adjacent cells
                (when (> width 1)
                  (.addMergedRegion
                    sh (CellRangeAddress.
                         row-idx row-idx col-idx (dec (+ col-idx width)))))

                ;; (3) Save the cell
                (vswap! layout assoc-in [row-idx col-idx] cell)
                (recur (+ col-idx ^long width) (rest cells)))))))

      ;; We can do this part in parallel at least, since the cells are all
      ;; different objects
      (let [layout @layout]
        (doparallel [row (map-indexed vector grid)]
          (let [[row-idx row-data] row]
            (loop [col-idx 0, cells row-data]
              (when-let [cell-data (first cells)]
                ;; (1) Find the cell
                (let [width (if (map? cell-data) (get cell-data :width 1) 1)
                      ^XSSFCell cell (get (get layout row-idx) col-idx)]

                ;; (2) Write the cell data
                (p :write-cell
                   (write-cell! cell (cond-> cell-data (map? cell-data) :value)))

                ;; (3) Set the cell style
                (let [style (build-style'
                              (if (map? cell-data) (:style cell-data) {}))]
                  (p :set-cell-style
                     (.setCellStyle cell style)))

                (recur (+ col-idx ^long width) (rest cells))))))))
      (catch Exception e
        (-> "Failed to write grid!"
            (ex-info {:sheet-name sheet-name :grid grid} e)
            (throw))))

    (dotimes [i (transduce (map count) (completing max) 0 grid)]

      ;; Only auto-size small tables because it takes forever (~10s on a large
      ;; grid)
      (when (< (count grid) 2000)
        (p :auto-size (.autoSizeColumn sh i)))

      (when (> (.getColumnWidth sh i) *max-col-width*)
        (.setColumnWidth sh i *max-col-width*)))

    (p :set-print-settings
       (.setFitToPage sh true)
       (.setFitWidth (.getPrintSetup sh) 1))
    sh))

(defn- workbook!
  "Create a new Apache POI XSSFWorkbook workbook object."
  []
  (XSSFWorkbook.))

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
        rows (mapv (apply juxt getters) tabular-data)
        header-style (or header-style
                         ;; Add right alignment if it's an accounting column
                         (fn [name]
                           (cond-> (style/default-header-style name)
                                   (@numeric? name)
                                   (assoc :alignment :right))))]
    (into
      [(mapv #(->{:value % :style (header-style %)}) headers)]
      rows)))

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
  (let [path' (force-extension path "xlsx")
        ;; Create the mutable, POI workbook object
        ^XSSFWorkbook wb
        (reduce
          (fn [wb [sheet-name grid]] (doto wb (write-grid! sheet-name grid)))
          (workbook!)
          (seq workbook))]
    (p :write-to-disk
       (with-open [fos (io/output-stream (io/file (str path')))]
         (.write wb fos)))
    (io/file path')))

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

(defn example []
  (quick-open
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
     [["First Column" "Second Column" {:value "A few merged" :width 3}]
      ["First Column Value" "Second Column Value"]
      ["This" "Row" "Has" "Its" "Own"
       {:value "Format" :style {:font {:bold true}}}]]}))

(comment
  ;; This will both open an example excel sheet and write & open a test pdf file
  ;; with the same contents. On platforms without OpenOffice the convert-pdf!
  ;; call will most likely fail.
  (open (convert-pdf! (example) (temp ".pdf"))))
