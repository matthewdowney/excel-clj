(ns excel-clj.poi-test
  (:require [excel-clj.poi :as poi]
            [excel-clj.file :as file]

            [clojure.test :refer :all]
            [clojure.java.io :as io]))


(deftest poi-writer-test
  (is (= (try (poi/example (file/temp ".xlsx")) :success (catch Exception e e))
         :success)
      "Example function writes successfully."))


(deftest performance-test
  (testing "Performance is reasonable"
    (println "Starting performance test -- writing to a temp file...")
    (let [tmp (file/temp ".xlsx")]
      (println "Warming up...")
      (dotimes [_ 3] (poi/performance-test tmp 50000))
      (println "Testing...")
      (dotimes [_ 3]
        (println "")
        (is (<= (poi/performance-test tmp 50000) 1000)
            "It should be (much) faster than 1 second to write 50k rows."))
      (io/delete-file tmp))
    (println "Performance test complete.")))
