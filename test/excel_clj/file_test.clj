(ns excel-clj.file-test
  (:require [clojure.test :refer :all]
            [excel-clj.file :as file]
            [clojure.java.io :as io]))

(deftest ^:office-integrations convert-to-pdf-test
  (let [input-file (clojure.java.io/resource "uptime-template.xlsx")
        temp-pdf-file (io/file (file/temp ".pdf"))]
    (try
      (testing "Convert XLSX file to PDF"
        (println "Writing example PDF...")
        (file/convert-pdf! input-file temp-pdf-file))
      (finally
        (io/delete-file temp-pdf-file)))))
