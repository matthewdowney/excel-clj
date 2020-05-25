(ns excel-clj.poi
  "Exposes a low level cell writer that uses Apache POI.

  See the `example` and `performance-test` functions at the end of
  this ns + the adjacent (comment ...) forms for more detail."
  {:author "Matthew Downey"}
  (:require [clojure.java.io :as io]
            [taoensso.encore :as enc]
            [excel-clj.style :as style]
            [clojure.walk :as walk]
            [taoensso.tufte :as tufte])
  (:import (java.io Closeable BufferedInputStream InputStream)
           (org.apache.poi.ss.usermodel RichTextString Sheet Cell Row Workbook)
           (java.util Date Calendar)
           (org.apache.poi.ss.util CellRangeAddress)
           (org.apache.poi.xssf.streaming SXSSFWorkbook)
           (org.apache.poi.xssf.usermodel XSSFWorkbook)))


(set! *warn-on-reflection* true)


(defprotocol IWorkbookWriter
  (dissoc-sheet! [this sheet-name]
    "If there's a sheet with the given name, get rid of it.")
  (workbook* [this]
    "Get the underlying Apache POI Workbook object."))


(defprotocol IWorksheetWriter
  (write! [this value] [this value style width height]
    "Write a single cell.

    If provided, `style` is a map shaped as described in excel-clj.style.

    Width and height determine cell merging, e.g. a width of 2 describes a
    cell that is merged into the cell to the right.")

  (newline! [this]
    "Skip the writer to the next row in the worksheet.")

  (sheet* [this]
    "Get the underlying Apache POI XSSFSheet object."))


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
                    (walk/postwalk-replace {sym typed-sym} if-true)]
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
  (if-type
    [data [Boolean Calendar String Date Double RichTextString]]
    (doto cell (.setCellValue data))

    ;; Apache POI requires that numbers be doubles
    (if (number? data)
      (doto cell (.setCellValue (double data)))

      ;; Otherwise stringify it
      (let [to-write (or (some-> data pr-str) "")]
        (doto cell (.setCellValue ^String to-write))))))


(defn- ensure-row! [{:keys [^Sheet sheet row row-cursor]}]
  (if-let [r @row]
    r
    (let [^int idx (vswap! row-cursor inc)]
      (vreset! row (.createRow sheet idx)))))


(defrecord ^:private SheetWriter
  [cell-style-cache ^Sheet sheet row row-cursor col-cursor]
  IWorksheetWriter
  (write! [this value]
    (write! this value nil 1 1))

  (write! [this value style width height]
    (let [^Row poi-row (ensure-row! this)
          ^int cidx (vswap! col-cursor inc)
          poi-cell (.createCell poi-row cidx)]

      (when (or (> width 1) (> height 1))
        ;; If the width is > 1, move the cursor along so that the next write on
        ;; this row happens in the next free cell, skipping the merged area
        (vswap! col-cursor + (dec width))
        (let [ridx @row-cursor
              cra (CellRangeAddress.
                    ridx (dec (+ ridx height))
                    cidx (dec (+ cidx width)))]
          (.addMergedRegion sheet cra)))

      (tufte/p :write-cell
        (write-cell! poi-cell value))

      (when-let [cell-style (cell-style-cache style)]
        (tufte/p :style-cell
          (.setCellStyle poi-cell cell-style))))

    this)

  (newline! [this]
    (vreset! row nil)
    (vreset! col-cursor -1)
    this)

  (sheet* [this]
    sheet)

  Closeable
  (close [this]
    (tufte/p :set-print-settings
      (.setFitToPage sheet true)
      (.setFitWidth (.getPrintSetup sheet) 1))
    this))


(defrecord ^:private WorkbookWriter
  [^Workbook workbook stream-factory owns-created-stream?]
  IWorkbookWriter
  (workbook* [this]
    workbook)

  (dissoc-sheet! [this sheet-name]
    (when-let [sh (.getSheet workbook sheet-name)]
      (.removeSheetAt workbook (.getSheetIndex workbook sh))
      sh))

  Closeable
  (close [this]
    (tufte/p :write-to-disk
      (if owns-created-stream? ;; We have to close the stream
        (with-open [fos ^Closeable (stream-factory this)]
          (.write workbook fos)
          (.close workbook))
        (let [fos (stream-factory this)] ;; Client is responsible for stream
          (.write workbook fos)
          (.close workbook))))))


(defn ^Sheet create-sheet [^Workbook workbook sheet-name]
  (when-let [sh (.getSheet workbook sheet-name)]
    (.removeSheetAt workbook (.getSheetIndex workbook sh)))
  (.createSheet workbook sheet-name))


(defn ^SheetWriter sheet-writer
  "Create a writer for an individual sheet within the workbook."
  [workbook-writer sheet-name]
  (let [{:keys [^Workbook workbook]} workbook-writer
        cache (enc/memoize_
                (fn [style]
                  (let [style (enc/nested-merge style/default-style style)]
                    (style/build-style workbook style))))
        sheet (create-sheet workbook sheet-name)]

    (map->SheetWriter
      {:cell-style-cache cache
       :sheet            sheet
       :row              (volatile! nil)
       :row-cursor       (volatile! -1)
       :col-cursor       (volatile! -1)})))


(defn ^WorkbookWriter writer
  "Open a writer for Excel workbooks.

  See `stream-writer` for writing to your own streams (maybe you're writing
  as a web server response, to S3, or otherwise over TCP).

  If `streaming?` is true (default), uses Apache POI streaming implementations.

  N.B. The streaming version is an order of magnitude faster than the
   alternative, so override this default only if you have a good reason!"
  ([path]
   (writer path true))
  ([path streaming?]
   (map->WorkbookWriter
     {:workbook (if streaming? (SXSSFWorkbook.) (XSSFWorkbook.))
      :path path
      :stream-factory #(io/output-stream (io/file (:path %)))
      :owns-created-stream? true})))


(defn- ^XSSFWorkbook appendable [path]
  (XSSFWorkbook. (BufferedInputStream. (io/input-stream path))))


(defn ^WorkbookWriter appender
  "Like `writer`, but allows overwriting individual sheets within a template
  workbook."
  ([from-path to-path]
   (appender from-path to-path true))
  ([from-path to-path streaming?]
   (map->WorkbookWriter
     {:workbook (if streaming?
                  (SXSSFWorkbook. (appendable from-path))
                  (appendable from-path))
      :path to-path
      :stream-factory #(io/output-stream (io/file (:path %)))
      :owns-created-stream? true})))


(defn ^WorkbookWriter stream-writer
  "Open a stream writer for Excel workbooks.

  If `streaming?` is true (default), uses Apache POI streaming implementations.

  N.B. The streaming version is an order of magnitude faster than the
  alternative, so override this default only if you have a good reason!"
  ([stream]
   (stream-writer stream true))
  ([stream streaming?]
   (map->WorkbookWriter
     {:workbook (if streaming? (SXSSFWorkbook.) (XSSFWorkbook.))
      :stream-factory (constantly stream)
      :owns-created-stream? false})))


(defn ^WorkbookWriter stream-appender
  "Like `stream-writer`, but allows overwriting individual sheets within a
  template workbook."
  ([from-stream to-stream]
   (stream-appender from-stream to-stream true))
  ([^InputStream from-stream to-stream streaming?]
   (map->WorkbookWriter
     (let [wb (XSSFWorkbook. from-stream)]
       {:workbook (if streaming? (SXSSFWorkbook. wb) wb)
        :stream-factory (constantly to-stream)
        :owns-created-stream? false}))))


(defn example [file-to-write-to]
  (with-open [w (writer file-to-write-to)
              t (sheet-writer w "Test")]
    (let [header-style {:border-bottom :thin :font {:bold true}}]
      (write! t "First Col" header-style 1 1)
      (write! t "Second Col" header-style 1 1)
      (write! t "Third Col" header-style 1 1)

      (newline! t)
      (write! t "Cell")
      (write! t "Wide Red Cell" {:font {:color :red}} 2 1)

      (newline! t)
      (write! t "Tall Cell" nil 1 2)
      (write! t "Cell 2")
      (write! t "Cell 3")

      (newline! t)
      ;; This one won't be visible, because it's hidden behind the tall cell
      (write! t "1")
      (write! t "2")
      (write! t "3")

      (newline! t)
      (write! t "Wide" nil 2 1)
      (write! t "Wider" nil 3 1)
      (write! t "Much Wider" nil 5 1))))


(defn template-example [file-to-write-to]
  ; The template here has a 'raw' sheet, which contains uptime data for 3 time
  ; series, and a 'Summary' sheet, wich uses formulas + the raw data to compute
  ; and plot. We're going to overwrite the 'raw' sheet to fill in the template.
  (let [template "resources/uptime-template.xlsx"]
    (with-open [w (appender template file-to-write-to)
                ; the template sheet to overwrite completely
                sh (sheet-writer w "raw")]

      (doseq [header ["Date" ; use the same headers as in the template
                      "Webserver Uptime"
                      "REST API Uptime"
                      "WebSocket API Uptime"]]
        (write! sh header))

      (newline! sh)

      ; then write random uptime values in one hour intervals
      (let [start-ts (inst-ms #inst"2020-05-01")
            one-hour (* 1000 60 60)]
        (dotimes [i 99]
          (let [row-ts (+ start-ts (* i one-hour))
                ymd {:data-format :ymd :alignment :left}]
            (write! sh (Date. ^long row-ts) ymd 1 1))

          ; random uptime values
          (write! sh (- 1.0 (rand 0.25)))
          (write! sh (- 1.0 (rand 0.25)))
          (write! sh (- 1.0 (rand 0.25)))
          (newline! sh))))))


(defn performance-test
  "Write `n-rows` of data to `to-file` and see how long it takes."
  [to-file n-rows & {:keys [streaming?] :or {streaming? true}}]
  (let [start (System/currentTimeMillis)
        header-style {:border-bottom :thin :font {:bold true}}]
    (with-open [w (writer to-file streaming?)
                sh (sheet-writer w "Test")]

      (write! sh "Date" header-style 1 1)
      (write! sh "Milliseconds" header-style 1 1)
      (write! sh "Days Since Start of 2018" header-style 1 1)
      (println "Wrote headers after" (- (System/currentTimeMillis) start) "ms")

      (let [start-ms (inst-ms #inst"2018")
            day-ms (enc/ms :days 1)]
        (dotimes [i n-rows]
          (let [ms (+ start-ms (* day-ms i))]
            (newline! sh)
            (write! sh (Date. ^long ms))
            (write! sh ms)
            (write! sh i))))

      (println "Wrote rows after" (- (System/currentTimeMillis) start) "ms"))

    (let [total (- (System/currentTimeMillis) start)]
      (println "Wrote file after" total "ms")
      total)))


(comment
  "Writing cells very manually"
  (example "cells.xlsx")

  "Filling in a template with random data"
  (template-example "filled-in-template.xlsx")

  "Testing overall performance, plus looking at streaming vs not streaming."
  ;; To get more detailed profiling output
  (tufte/add-basic-println-handler! {})

  ;;; 200,000 rows with and without streaming
  (tufte/profile {} (performance-test "test.xlsx" 200000 :streaming? true))
  ;=> 2234

  (tufte/profile {} (performance-test "test.xlsx" 200000 :streaming? false) )
  ;=> 11187


  ;;; 300,000 rows with and without streaming
  (tufte/profile {} (performance-test "test.xlsx" 500000 :streaming? true))
  ;=> 5093

  (tufte/profile {} (performance-test "test.xlsx" 500000 :streaming? false))
  ; ... like a 2 minute delay and then OOM error (with my 8G of ram) ... haha
  )
