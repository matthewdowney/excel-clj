(ns excel-clj.prototype
  "Prototype features to be included in v2.0.0 -- everything subject to change."
  {:author "Matthew Downey"}
  (:require [excel-clj.poi :as poi]
            [clojure.java.io :as io]
            [clojure.string :as string]
            [taoensso.encore :as enc])
  (:import (java.io File)
           (org.jodconverter.office DefaultOfficeManagerBuilder)
           (org.jodconverter OfficeDocumentConverter)
           (java.awt Desktop HeadlessException)))


(set! *warn-on-reflection* true)


;;; Code to model the 'cell' an Excel document.
;;; A cell can be either a plain value (a string, java.util.Date, etc.) or a
;;; such a value wrapped inside of a map which also includes style and dimension
;;; data.


(defn wrapped
  "If `x` contains cell data wrapped in a map (with style & dimension data),
  return it as-is. Otherwise return a wrapped version."
  [x]
  (if (::wrapped? x)
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
  (if (::wrapped? x)
    (::data x)
    x))


;;; Code to build Excel worksheets out of Clojure's data structures


(comment
  "I'm not really sure if this stuff is helpful..."

  (defn- ensure-rows [sheet n] (into sheet (repeat (- n (count sheet)) [])))
  (defn- ensure-cols [row n] (into row (repeat (- n (count row)) nil)))


  (defn write
    "Write to the cell in the `sheet` grid at `(x, y)`."
    [sheet x y cell]
    (let [sheet-data (ensure-rows sheet y)
          row (-> (get sheet-data y [])
                  (ensure-cols x)
                  (assoc x cell))]
      (assoc sheet-data y row)))


  (defn write-row
    "Append `row` to the `sheet` grid."
    [sheet row]
    (conj (or sheet []) row)))


;; TODO: Table -> [[cell]]
;; TODO: Tree -> [[cell]]


;;; Code to convert [[cell]] to .xlsx documents, etc. -- IO stuff


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


(defn- write-rows!
  "Write the rows via the given sheet-writer, returning [the number of rows
  written, number of columns written]."
  [sheet-writer rows-seq]
  (reduce
    (fn [[rows cols] next-row]
      (doseq [cell next-row]
        (let [{:keys [width height]} (dims cell)]
          (poi/write! sheet-writer (data cell) (style cell) width height)))
      (poi/newline! sheet-writer)
      [(inc rows) (max cols (count next-row))])
    [0 0]
    rows-seq))


(defn write!
  "Write the `workbook` to the given `path` and return a file object pointing
  at the written file.

  The workbook is a key value collection of (sheet-name grid), either as map or
  an association list (if ordering is important)."
  [workbook path]
  (let [f (io/file (force-extension (str path) ".xlsx"))]
    (with-open [w (poi/writer f)]
      (doseq [[nm rows] workbook
              :let [sh (poi/sheet-writer w nm)
                    [rows-written cols-written] (write-rows! sh rows)]]
        ;; Only auto-size columns for small sheets, otherwise it takes forever
        (when (< rows-written 2000)
          (dotimes [i cols-written]
            (poi/autosize!! sh i)))))
    f))


(defn temp
  "Return a (string) path to a temp file with the given extension."
  [ext]
  (-> (File/createTempFile "generated-sheet" ext) .getCanonicalPath))


(defn- convert-pdf!
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


  ;; Ballpark performance test
  (dotimes [_ 5]
    (time
      (let [header-style {:border-bottom :thin :font {:bold true}}
            headers (map #(cell/style % header-style) ["N0" "N1" "N2"])]
        (write!
          [["Test" (cons headers (for [x (range 100000)] [x (str x) x]))]]
          "test.xlsx"))))
  ; "Elapsed time: 5484.565899 msecs"
  ; "Elapsed time: 5302.312954 msecs"
  ; "Elapsed time: 4656.451894 msecs"
  ; "Elapsed time: 4734.160618 msecs"
  ; "Elapsed time: 5753.986336 msecs"

  )
