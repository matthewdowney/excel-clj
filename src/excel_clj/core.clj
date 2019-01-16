(ns
  ^{:doc "Utilities for declarative creation of Excel (.xlsx) spreadsheets,
  with higher level abstractions over Apache POI (https://poi.apache.org/).

  The highest level data abstraction used to create excel spreadsheets is a
  tree, followed by a table, and finally the most basic abstraction is a grid.

  The tree and table functions convert tree formatted or tabular data into a
  grid of [[cell]]. A workbook is a key value collection of
  {page-name page-grid}.

  The table is probably the most intuitive abstraction, give it a run:
    (quick-view
      (table
        [{\"Date\" \"2018-01-01\" \"% Return\" 0.05M \"USD\" 1500.5005M}
         {\"Date\" \"2018-02-01\" \"% Return\" 0.04M \"USD\" 1300.20M}
         {\"Date\" \"2018-03-01\" \"% Return\" 0.07M \"USD\" 2100.66666666M}]))

  Run the (example) function at the bottom of this namespace to see more.

  ## Values & Styles

  The basic unit of spreadsheet data is the cell, which can either be a plain
  value or a value with style data.

    {:value 0.2345
     :style {:data-format :percent,
             :font {:bold true :font-height-in-points 10}}}

  The goal of the style map is to reuse all of the functionality built in to the
  underlying Apache POI objects, but with immutable data structures. Advantages
  include using `nested-merge` to combine styles plus all of the things you
  already expect of maps.

  Style map data are representative of nested calls to the corresponding setter
  methods in the Apache POI framework starting with a CellStyle object. That is,
  the above example is roughly interpreted as:

    ;; A CellStyle POI object created under the hood during rendering
    (let [cell-style ...]

      ;; The style map attributes are converted to camel cased setters
      (doto cell-style
        (.setDataFormat :percent)
        (.setFont
          (doto <create a font>
            (.setBold true)
            (.setFontHeightInPoints 10)))))

  The two nontrivial challenges are
    - creating nested objects, e.g. (.setFont cell-style <font>) needs to be
      called with a POI Font object; and
    - translating keywords like :percent to POI objects.

  Both are solved with the coerce-to-obj multimethod, which specifies how to
  coerce different attributes to POI objects. It has the signature
    (workbook, attribute, value) => Object
  and dispatches on the attribute (a keyword).

  We coerce key value pairs to objects from the bottom of the style map upwards,
  meaning that by the time coerce-to-obj is being invoked for some attribute,
  any nested attributes in the value have already been coerced.

  A more nuanced representation of how the style map 'expands':

    ;; {:data-format :percent, :font {:bold true :font-height-in-points 10}}}
    ;; expands to
    (let [cell-style ..., workbook ...] ;; POI objects created during rendering
      (doto cell-style
        (.setDataFormat (coerce-to-obj workbook :data-format :percent))
        ;; The {:bold true :font-height-in-points 10} expands recursively
        (.setFont
          (coerce-to-obj
            workbook :font {:bold true :font-height-in-points 10}))))"
    :author "Matthew Downey"} excel-clj.core
  (:require [excel-clj.tree :as tree]
            [clojure.string :as string]
            [clojure.java.io :as io])
  (:import (org.apache.poi.ss.usermodel DataFormat Cell RichTextString
                                        BorderStyle HorizontalAlignment
                                        FontUnderline)
           (org.apache.poi.xssf.usermodel XSSFWorkbook XSSFSheet XSSFColor
                                          DefaultIndexedColorMap)
           (java.io FileOutputStream File)
           (java.awt Desktop HeadlessException)
           (java.util Calendar Date)
           (org.apache.poi.ss.util CellRangeAddress)
           (org.jodconverter.office DefaultOfficeManagerBuilder)
           (org.jodconverter OfficeDocumentConverter)))

;;; Code to allow specification of Excel CellStyle objects as nested maps. You
;;; might touch this code to add an implementation of `coerce-to-obj` for some
;;; cell style attribute.

(defn do-set!
  "Set an attribute on a Java object & return the object. E.g.
    (let [attr :font-height-in-points]
      (do-set! some-obj attr 14))
    ;; Equivalent to (doto some-obj (.setFontHeightInPoints 14))"
  [obj attr val]
  (let [cap (fn [coll] (map string/capitalize coll))
        camel (fn [kw]
                (str "set" (-> (name kw) (string/split #"\W") cap string/join)))
        setter (eval
                 (read-string (format "(memfn %s arg)" (-> attr camel))))]
    (doto obj
      (setter val))))

(defn do-set-all!
  [base-object attributes]
  (reduce
    (fn [left [attr val]] (do-set! left attr val))
    base-object attributes))

(defmulti coerce-to-obj
          "For some keyword attribute of a CellStyle object, attempt to coerce clojure
          data (either a keyword or a map) to the Java object the setter is expecting.

          This allows nesting of style specification maps
            {:font {:bold true, :color :yellow}}
          so that when it's time to generate a CellStyle object, we can say that we
          know how to go from an attribute map to a Font object for :font attributes,
          from a keyword to a Color object for :color attributes, etc."
          (fn [^XSSFWorkbook workbook attr-keyword value]
            attr-keyword))

;; Coercions from simple map lookups

(defmacro coerce-from-map
  ([attr-keyword coercion-map]
   `(coerce-from-map ~attr-keyword ~coercion-map (fn [a# b# val#] val#)))
  ([attr-keyword coercion-map otherwise]
   `(defmethod coerce-to-obj ~attr-keyword
      [wb# akw# val#]
      (if (keyword? val#)
        (or
          (get ~coercion-map val#)
          (-> ~(str "No " attr-keyword " registered.")
              (ex-info {:given val# :have (keys ~coercion-map)})
              (throw)))
        (~otherwise wb# akw# val#)))))

(def alignments
  {:general          HorizontalAlignment/GENERAL
   :left             HorizontalAlignment/LEFT
   :center           HorizontalAlignment/CENTER
   :right            HorizontalAlignment/RIGHT
   :fill             HorizontalAlignment/FILL
   :justify          HorizontalAlignment/JUSTIFY
   :center-selection HorizontalAlignment/CENTER_SELECTION
   :distributed      HorizontalAlignment/DISTRIBUTED})

(def underlines
  {:single            FontUnderline/SINGLE
   :single-accounting FontUnderline/SINGLE_ACCOUNTING
   :double            FontUnderline/DOUBLE
   :double-accounting FontUnderline/DOUBLE_ACCOUNTING
   :none              FontUnderline/NONE})

(def borders
  {:none                BorderStyle/NONE
   :thin                BorderStyle/THIN
   :medium              BorderStyle/MEDIUM
   :dashed              BorderStyle/DASHED
   :dotted              BorderStyle/DOTTED
   :thick               BorderStyle/THICK
   :double              BorderStyle/DOUBLE
   :hair                BorderStyle/HAIR
   :medium_dashed       BorderStyle/MEDIUM_DASHED
   :dash_dot            BorderStyle/DASH_DOT
   :medium_dash_dot     BorderStyle/MEDIUM_DASH_DOT
   :dash_dot_dot        BorderStyle/DASH_DOT_DOT
   :medium_dash_dot_dot BorderStyle/MEDIUM_DASH_DOT_DOT
   :slanted_dash_dot    BorderStyle/SLANTED_DASH_DOT})

(def data-formats
  {:accounting "_($* #,##0.00_);_($* (#,##0.00);_($* \"-\"??_);_(@_)"
   :ymd "yyyy-MM-dd"
   :percent "0.00%"})

(defn ^XSSFColor rgb-color
  "Create an XSSFColor object from the given r g b values."
  [r g b]
  (XSSFColor. (byte-array [r g b]) (DefaultIndexedColorMap.)))

(def colors
  {:gray   (rgb-color 232 232 232)
   :blue   (rgb-color 204 255 255)
   :green  (rgb-color 221 255 204)
   :yellow (rgb-color 250 255 204)})

(coerce-from-map :alignment alignments)
(coerce-from-map :underline underlines)
(coerce-from-map :border-top borders)
(coerce-from-map :border-left borders)
(coerce-from-map :border-right borders)
(coerce-from-map :border-bottom borders)

(coerce-from-map :color colors
                 ;; If there's nothing in the map ...
                 (fn [_ _ color]
                   (if (and (coll? color) (= (count color) 3))
                     (apply rgb-color color)
                     (-> "Can only create colors from rgb three-tuples or keywords."
                         (ex-info {:given color})
                         (throw)))))

(defmethod coerce-to-obj :font
  [wb _ font-attrs]
  (do-set-all! (.createFont wb) font-attrs))

(defmethod coerce-to-obj :data-format
  [wb _ format]
  (if (instance? DataFormat format)
    format
    (if-let [format' (cond->> format (keyword? format) (get data-formats))]
      (let [ch (.getCreationHelper wb)]
        (.getFormat ^DataFormat (.createDataFormat ch) (str format')))
      (-> "Can't coerce to data format."
          (ex-info {:given format :have (keys data-formats)})
          (throw)))))

(defmethod coerce-to-obj :default
  [_ _ x] x)

(defn coerce-nested-to-obj
  "Given an attribute map, start at the most nested layer and work upwards,
  attempting to coerce each attribute to an object."
  [wb attributes]
  (letfn [(coerce-nested [av-pairs rebuilt]
            (if-let [[a v] (first av-pairs)]
              (if-not (map? v)
                (let [coerced (coerce-to-obj wb a v)]
                  (recur (rest av-pairs) (assoc rebuilt a coerced)))
                #(coerce-nested
                   (rest av-pairs)
                   (->>
                     (coerce-to-obj wb a (trampoline coerce-nested (seq v) {}))
                     (assoc rebuilt a))))
              rebuilt))]
    (trampoline coerce-nested attributes {})))

(defn build-style
  "Create a CellStyle from the given attrs using the given workbook
  CellStyle attributes are anything that can be set with
  .setCamelCasedAttribute on a CellStyle object, including
    {:data-format  string or keyword
     :font         { ... font attrs ... }
     :wrap-text    boolean
     :hidden       boolean
     :alignment    org.apache.poi.ss.usermodel.HorizontalAlignment
     :border-[bottom|left|right|top] org.apache.poi.ss.usermodel.BorderStyle}

  Any of the attributes can be java objects. Alternatively, if a `coerce-to-obj`
  implementation is provided for some attribute (e.g. :font), the attribute can
  be specified as data."
  [workbook attrs]
  (let [attrs' (coerce-nested-to-obj workbook attrs)]
    (try
      (do-set-all! (.createCellStyle workbook) attrs')
      (catch Exception e
        (-> "Failed to create cell style."
            (ex-info {:raw-attributes attrs :built-attributes attrs'} e)
            (throw))))))

;;; Low level code to write to & style sheets; you probably shouldn't have to
;;; touch this to make use of the API, but might choose to when adding or
;;; extending functionality

(def default-style
  {:font {:font-height-in-points 10 :font-name "Arial"}})

(defn nested-merge
  "Like merge except nested maps are also merged. E.g.
    (merge-nested {:foo {:a :b}} {:foo {:c :d}})
      ; => {:foo {:a :b, :c :d}}"
  [& maps]
  (letfn [(merge? [left right]
            (if (and (map? left) (map? right))
              (merge-2 left right)
              right))
          (merge-entry [m e]
            (let [k (key e) v (val e)]
              (if (contains? m k)
                #(assoc m k (merge? (get m k) v))
                (assoc m k v))))
          (merge-2 [m1 m2]
            (reduce #(trampoline merge-entry %1 %2) (or m1 {}) (seq m2)))]
    (when (some identity maps)
      (reduce merge-2 maps))))

(defn merge-style
  "Merge cell's current style with the provided style map, preserving any style
  that does not conflict."
  [cell style]
  (update
    (if (map? cell) cell {:value cell})
    :style (fn [s] (if-not s style (nested-merge s style)))))

(defmacro if-type
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

(defn write-cell!
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
             (doto cell (.setCellValue ^String (or (some-> data pr-str) ""))))))

(def ^:dynamic *max-width*
  15000)

(defn ^XSSFSheet write-grid!
  "Modify the given workbook by adding a sheet with the given name built from
  the provided grid.

  The grid is a collection of rows, where each cell is either a plain, non-map
  value or a map of {:value ..., :style ..., :width ...}, with :value being the
  contents of the cell, :style being an optional map of style data, and :width
  being an optional cell width dictating how many horizontal slots the cell
  takes up (creates merged cells).

  Returns the sheet object."
  [workbook sheet-name grid]
  (let [^XSSFSheet sh (.createSheet workbook sheet-name)
        build-style' (memoize ;; Immutable styles can share mutable objects :)
                       (fn [style-map]
                         (->> (nested-merge default-style (or style-map {}))
                              (build-style workbook))))]
    (try
      (doseq [[row-idx row-data] (map-indexed vector grid)]
        (let [row (.createRow sh (int row-idx))]
          (loop [col-idx 0 cells row-data]
            (when-let [cell-data (first cells)]
              (let [cell (.createCell row col-idx)
                    width (if (map? cell-data) (get cell-data :width 1) 1)]
                (write-cell! cell (cond-> cell-data (map? cell-data) :value))
                (.setCellStyle
                  cell
                  (build-style' (if (map? cell-data) (:style cell-data) {})))
                (when (> width 1)
                  (.addMergedRegion
                    sh (CellRangeAddress.
                         row-idx row-idx col-idx (dec (+ col-idx width)))))
                (recur (+ col-idx width) (rest cells)))))))
      (catch Exception e
        (-> "Failed to write grid!"
            (ex-info {:sheet-name sheet-name :grid grid} e)
            (throw))))

    (dotimes [i (transduce (map count) (completing max) 0 grid)]
      (.autoSizeColumn sh i)
      (when (> (.getColumnWidth sh i) *max-width*)
        (.setColumnWidth sh i *max-width*)))

    (.setFitToPage sh true)
    (.setFitWidth (.getPrintSetup sh) 1)
    sh))

(defn workbook!
  "Create a new Apache POI XSSFWorkbook workbook object."
  []
  (XSSFWorkbook.))

;;; Higher-level code to specify sheets in terms of clojure data structures.
;;;   - table has the shape [{"Column Name" "Cell Value"}] with optional styles
;;;   - tree has the shape ["Section"
;;;                         [["Subsec 1" {"Column" Numeric Value}]
;;;                          ["Subsec 2" {"Column" Numeric Value}]]]
;;;     (nested as deeply as you want)

(defn best-guess-row-format
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

(def default-header-style
  (constantly
    {:border-bottom BorderStyle/THIN :font {:bold true}}))

(defn table
  "Build a sheet grid from the provided collection of tabular data, where each
  item has the format {Column Name, Cell Value}.

  If provided
    headers      is an ordered coll of column names
    header-style is a function header-name => style map for the header.
    data-style   is a function that takes (datum-map, column name) and returns
                 a style specification or nil for the default style."
  [tabular-data & {:keys [headers header-style data-style]
                   :or {data-style best-guess-row-format}}]
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
                    (let [style (data-style row col-name)]
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
                           (cond-> (default-header-style name)
                                   (@numeric? name)
                                   (assoc :alignment :right))))]
    (into
      [(mapv #(->{:value % :style (header-style %)}) headers)]
      rows)))

(def default-tree-formatters
  {0 {:font {:bold true} :border-bottom :medium}
   1 {:font {:bold true}}
   2 {:indention 2}
   3 {:font {:italic true} :alignment :right}})

(def default-tree-total-formatters
  {0 {:font {:bold true} :border-top :medium}
   1 {:border-top :thin :border-bottom :thin}})

(defn formatter-or-max [formatters n]
  (or
    (get formatters n)
    (second (apply max-key first formatters))))

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
  ([t]
   (tree t nil))
  ([t headers]
   (tree t headers default-tree-formatters default-tree-total-formatters))
  ([t headers formatters total-formatters]
   (try
     (let [tabular (apply tree/render-table (second t))
           all-colls (or headers
                         (sequence
                           (comp
                             (mapcat keys)
                             (filter (complement #{:depth :label}))
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
           (let [total? (empty? (str (:label line)))
                 format (or
                          (formatter-or-max
                            (if total? total-formatters formatters)
                            (:depth line))
                          {})
                 style (nested-merge format {:data-format :accounting})]
             (into [{:value (:label line) :style (if total? {} style)}]
                   (map #(->{:value (get line %) :style style})) all-colls)))))
     (catch Exception e
       (throw (ex-info "Failed to render tree" {:tree t} e))))))

(defn with-title
  "Write a title above the given grid with a width equal to the widest row."
  [grid title]
  (let [width (transduce (map count) (completing max) 0M grid)]
    (concat
      [[{:value title :width width :style {:alignment :center}}]]
      grid)))

(defn- force-extension [path ext]
  (let [path (.getCanonicalPath (io/file path))]
    (if (.endsWith path ext)
      path
      (let [parts (string/split path (re-pattern File/separator))]
        (str
          (string/join
            File/separator (if (> (count parts) 1) (butlast parts) parts))
          "." ext)))))

(defn write!
  "Write the workbook to the given filename and return a file object pointing
  at the written file.

  The workbook is a key value collection of (sheet-name grid), either as map or
  an association list (if ordering is important)."
  [workbook path]
  (let [path' (force-extension path "xlsx")
        ;; Create the mutable, POI workbook objects
        wb (reduce
             (fn [wb [sheet-name grid]] (doto wb (write-grid! sheet-name grid)))
             (workbook!)
             (seq workbook))]
    (with-open [fos (FileOutputStream. (str path'))]
      (.write wb fos))
    (io/file path')))

;;; Convenience utilities

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

