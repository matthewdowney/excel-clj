(ns excel-clj.cell
  "A lightweight wrapper over cell values that allows combining both simple
  and wrapped cells with new stile and dimensions."
  {:author "Matthew Downey"}
  (:require [taoensso.encore :as enc]))


(defn wrapped? [x] (:excel/wrapped? x))


(defn wrapped
  "If `x` contains cell data wrapped in a map (with style & dimension data),
  return it as-is. Otherwise return a wrapped version."
  [x]
  (if (wrapped? x)
    x
    {:excel/wrapped? true :excel/data x}))


(defn style
  "Get the style specification for `x`, or deep-merge its current style spec
  with the given `style-map`."
  ([x]
   (or (:excel/style x) {}))
  ([x style-map]
   (let [style-map (enc/nested-merge (style x) style-map)]
     (assoc (wrapped x) :excel/style style-map))))


(defn dims
  "Get the {:width N, :height N} dimension map for `x`, or merge in the given
  `dims-map` of the same format."
  ([x]
   (or (:excel/dims x) {:width 1 :height 1}))
  ([x dims-map]
   (let [dims-map (merge (dims x) dims-map)]
     (assoc (wrapped x) :excel/dims dims-map))))


(defn data
  "If `x` contains cell data wrapped in a map (with style & dimension data),
  return the wrapped cell value. Otherwise return as-is."
  [x]
  (if (wrapped? x)
    (:excel/data x)
    x))


(comment
  "You don't have to worry about if something is already wrapped or already
  styled:"

  (def cell
    (let [header-style {:border-bottom :thin :font {:bold true}}]
      (-> "Header"
          (style header-style)
          (dims {:height 2})
          (style {:vertical-alignment :center}))))

  (clojure.pprint/pprint cell)
  ; #:excel{:wrapped? true,
  ;         :data "Header",
  ;         :style
  ;         {:border-bottom :thin,
  ;          :font {:bold true},
  ;          :vertical-alignment :center},
  ;         :dims {:width 1, :height 2}}
  )
