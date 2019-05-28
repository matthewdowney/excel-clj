(ns excel-clj.tree-test
  (:require [clojure.test :refer :all])
  (:require [excel-clj.tree :refer :all]))

(def ^:private cash-leaf
  ["Cash" {2018 100M, 2017 85M}])

(deftest leaf?-test
  (testing "Recognizes a leaf."
    (is (true? (leaf? cash-leaf))))
  (testing "Recognizes a non-leaf."
    (is (not (leaf? ["Assets" [cash-leaf]])))))

(deftest children-test
  (testing "Returns children of a node"
    (let [alt-leaf (assoc cash-leaf 0 "Kash")]
      (is
        (= (children ["Assets" [cash-leaf alt-leaf]]) [cash-leaf alt-leaf])))))

(deftest value-test
  (let [[assets liabilities-equity] mock-balance-sheet]
    (testing "Sums trees properly"
      (is (= {2018 217M, 2017 148M}, (value assets) (value liabilities-equity))))
    (testing "Sums a single leaf properly."
      (is (= {2018 100M, 2017 85M} (value cash-leaf))))))

(deftest math-test
  (testing "Addition & subtraction works on trees & maps"
    (let [[assets liabilities-equity] mock-balance-sheet]
      ;; Assets  & liabilities/equity cancel each other, leaving just the map
      (is (= (math (- assets (+ liabilities-equity {2018 1, 2017 2})))
             {2018 -1M, 2017 -2M})))))

(deftest negate-tree-test
  (testing "Negates the values in a tree."
    (let [[assets liabilities-equity] mock-balance-sheet]
      (is (= (math (- assets liabilities-equity))
             (math (+ assets (negate-tree liabilities-equity)))
             (math (+ (negate-tree assets) liabilities-equity)))))))

(deftest shallow-test
  (testing "Maintains tree structure while combining labels."
    (let [[assets liabilities-equity] mock-balance-sheet
          shallowed (shallow ["" [assets (negate-tree liabilities-equity)]])]
      (is (= (value shallowed) {2018 0M, 2017 0M}))
      (is (= (label shallowed)
             "Assets & Liabilities & Stockholders' Equity")))))
