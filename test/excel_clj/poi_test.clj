(ns excel-clj.poi-test
  (:require [clojure.test :refer :all]

            [excel-clj.poi :as poi]
            [excel-clj.file :as file]))

(deftest poi-writer-test
  (is (= (try (poi/example (file/temp ".xlsx")) :success
              (catch Exception e e))
         :success)
      "Example function writes successfully."))


(deftest performance-test
  (testing "Performance is reasonable"
    (println "Starting performance test -- writing to a temp file...")
    (dotimes [_ 3]
      (println "")
      (is (<= (poi/performance-test (file/temp ".xlsx") 50000) 1000)
          "It should be (much) faster than 1 second to write 50k rows."))
    (println "Performance test complete.")))
