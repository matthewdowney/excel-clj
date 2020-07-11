(ns excel-clj.core-test
  (:require [excel-clj.cell :refer :all]
            [excel-clj.core :refer :all]
            [excel-clj.file :refer [temp]]

            [clojure.test :refer :all]
            [clojure.java.io :as io]))


(deftest table-test
  (let [td [{"Date" "2018-01-01" "% Return" 0.05M "USD" 1500.5005M}
            {"Date" "2018-02-01" "% Return" 0.04M "USD" 1300.20M}
            {"Date" "2018-03-01" "% Return" 0.07M "USD" 2100.66666666M}]
        generated (table td)]
    (testing "Generated grid has the expected shape for the tabular data"
      (is (= (mapv #(mapv data %) generated)
             [["Date" "% Return" "USD"]
              ["2018-01-01" 0.05M 1500.5005M]
              ["2018-02-01" 0.04M 1300.20M]
              ["2018-03-01" 0.07M 2100.66666666M]])))))


(deftest tree-test
  (let [data {"Title"
              {"Tree 1" {"Child" {2018 2, 2017 1}
                         "Another" {2018 3, 2017 1}}
               "Tree 2" {"Child" {2018 -2, 2017 -1}}}}]
    (testing "Renders tree into a grid with a title and total rows."
      (is (= (mapv #(mapv :value %) (tree data)))
          [["Title"]
           [nil 2018 2017]
           ["Tree 1" nil nil]
           ["Child" 2 1]
           ["Another" 3 1]
           ["" 5 2]
           ["Tree 2" nil nil]
           ["Child" -2 -1]]))))


(deftest example-test
  (let [temp-file (io/file (temp ".xlsx"))]
    (try
      (testing "Example code snippet writes successfully."
        (write! example-workbook-data temp-file))
      (finally
        (io/delete-file temp-file)))))


(deftest template-example-test
  (let [temp-file (io/file (temp ".xlsx"))]
    (try
      (testing "Example code snippet writes successfully."
        (let [template (clojure.java.io/resource "uptime-template.xlsx")
              new-data {"raw" (table example-template-data)}]
          (append! new-data template "filled-in-template.xlsx")))
      (finally
        (io/delete-file temp-file)))))
