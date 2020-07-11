(ns excel-clj.file
  "Write Clojure grids of `[[cell]]` as Excel worksheets, convert Excel
  worksheets to PDFs, and read Excel worksheets.

  A cell can be either a plain value (a string, java.util.Date, etc.) or such
  a value wrapped in a map which also includes style and dimension data.

  Check out the (example) function at the bottom of this namespace for more."
  {:author "Matthew Downey"}
  (:require [excel-clj.cell :refer [dims data style]]
            [excel-clj.poi :as poi]

            [clojure.string :as string]
            [clojure.java.io :as io])
  (:import (org.apache.poi.xssf.streaming SXSSFSheet)
           (org.apache.poi.ss.usermodel Sheet)
           (java.io File)
           (org.jodconverter.office DefaultOfficeManagerBuilder)
           (org.jodconverter OfficeDocumentConverter)
           (java.awt Desktop HeadlessException)))


;;; Code to write [[cell]]


(defn write-rows!
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


(defn write*
  "For best performance, using {:streaming true, :auto-size-cols? false}."
  [workbook poi-writer {:keys [streaming? auto-size-cols?] :as ops}]
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


(defn write! ; see core/write!
  ([workbook path]
   (write! workbook path (default-ops workbook)))
  ([workbook path {:keys [streaming? auto-size-cols?]
                   :or   {streaming? true}
                   :as   ops}]
   (let [f (io/file (force-extension (str path) ".xlsx"))]
     (with-open [w (poi/writer f streaming?)]
       (write* workbook w (assoc ops :streaming? streaming?)))
     f)))


(defn write-stream! ; see core/write-stream!
  ([workbook stream]
   (write-stream! workbook stream (default-ops workbook)))
  ([workbook stream {:keys [streaming? auto-size-cols?]
                     :or   {streaming? true}
                     :as   ops}]
   (with-open [w (poi/stream-writer stream)]
     (write* workbook w (assoc ops :streaming? streaming?)))))


;;; Other file utilities


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


(defn write-pdf! [workbook path] ; see core/write-pdf!
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


(defn quick-open! [workbook]
  (open (write! workbook (temp ".xlsx"))))


(defn quick-open-pdf! [workbook]
  (open (write-pdf! workbook (temp ".pdf"))))


(defn example
  "Write & open a sheet composed of a simple grid."
  []
  (let [grid [["A" "B" "C"]
              [1 2 3]]]
    (quick-open! {"Sheet 1" grid})))


(defn example-plus
  "Write & open a sheet composed of a more involved grid."
  []
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
        excel-file (quick-open! {"Sheet 1" (concat header-rows grid)})]

    (try
      (open (convert-pdf! excel-file (temp ".pdf")))
      (catch Exception e
        (println "(Couldn't open a PDF on this platform.)")))))
