(ns
  ^{:doc "The basic unit of spreadsheet data is the cell, which can either be a
  plain value or a value with style data.

    {:value 0.2345
     :style {:data-format :percent,
             :font {:bold true :font-height-in-points 10}}}

  The goal of the style map is to reuse all of the functionality built in to
  the underlying Apache POI objects, but with immutable data structures.

  The primary advantage -- beyond the things we're accustomed to loving about
  clojure data maps as opposed to mutable objects with getters/setters -- other
  than declarative syntax is the ease with which we can merge styles as maps
  rather than trying to create some new POI object out of two other objects,
  reading and combining all of their attributes and nested attributes.

  ## Mechanics

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
    :author "Matthew Downey"} excel-clj.style
  (:require [clojure.string :as string]
            [clojure.reflect :as reflect]
            [rhizome.viz :as viz])
  (:import (org.apache.poi.ss.usermodel
             DataFormat BorderStyle HorizontalAlignment FontUnderline
             FillPatternType)
           (org.apache.poi.xssf.usermodel
             XSSFWorkbook XSSFColor DefaultIndexedColorMap XSSFCell)))

;;; Code to allow specification of Excel CellStyle objects as nested maps. You
;;; might touch this code to add an implementation of `coerce-to-obj` for some
;;; cell style attribute.

(defn- do-set!
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

(defn- do-set-all!
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

(defmacro ^:private coerce-from-map
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
   :medium-dashed       BorderStyle/MEDIUM_DASHED
   :dash-dot            BorderStyle/DASH_DOT
   :medium-dash-dot     BorderStyle/MEDIUM_DASH_DOT
   :dash-dot-dot        BorderStyle/DASH_DOT_DOT
   :medium-dash-dot-dot BorderStyle/MEDIUM_DASH_DOT_DOT
   :slanted-dash-dot    BorderStyle/SLANTED_DASH_DOT})

(def fill-patterns
  {:no-fill             FillPatternType/NO_FILL
   :solid-foreground    FillPatternType/SOLID_FOREGROUND
   :fine-dots           FillPatternType/FINE_DOTS
   :alt-bars            FillPatternType/ALT_BARS
   :sparse-dots         FillPatternType/SPARSE_DOTS
   :thick-horz-bands    FillPatternType/THICK_HORZ_BANDS
   :thick-vert-bands    FillPatternType/THICK_VERT_BANDS
   :thick-backward-diag FillPatternType/THICK_BACKWARD_DIAG
   :thick-forward-diag  FillPatternType/THICK_FORWARD_DIAG
   :big-spots           FillPatternType/BIG_SPOTS
   :bricks              FillPatternType/BRICKS
   :thin-horz-bands     FillPatternType/THIN_HORZ_BANDS
   :thin-vert-bands     FillPatternType/THIN_VERT_BANDS
   :thin-backward-diag  FillPatternType/THIN_BACKWARD_DIAG
   :thin-forward-diag   FillPatternType/THIN_FORWARD_DIAG
   :squares             FillPatternType/SQUARES
   :diamonds            FillPatternType/DIAMONDS
   :less_dots           FillPatternType/LESS_DOTS
   :least_dots          FillPatternType/LEAST_DOTS})

(def data-formats
  {:accounting "_($* #,##0.00_);_($* (#,##0.00);_($* \"-\"??_);_(@_)"
   :ymd "yyyy-MM-dd"
   :percent "0.00%"})

(defn ^XSSFColor rgb-color
  "Create an XSSFColor object from the given r g b values."
  [r g b]
  (XSSFColor. (byte-array [r g b]) (DefaultIndexedColorMap.)))

(def colors
  {:white  (rgb-color 255 255 255)
   :red    (rgb-color 255 0 0)
   :orange (rgb-color 255 127 0)
   :yellow (rgb-color 250 255 204)
   :green  (rgb-color 221 255 204)
   :blue   (rgb-color 204 255 255)
   :purple (rgb-color 200 0 255)
   :gray   (rgb-color 232 232 232)
   :black  (rgb-color 0 0 0)})

(coerce-from-map :alignment alignments)
(coerce-from-map :underline underlines)
(coerce-from-map :border-top borders)
(coerce-from-map :border-left borders)
(coerce-from-map :border-right borders)
(coerce-from-map :border-bottom borders)
(coerce-from-map :fill-pattern fill-patterns)

(letfn [(if-color-not-found [_ _ color]
          (if (and (coll? color) (= (count color) 3))
            (apply rgb-color color)
            (-> "Can only create colors from rgb three-tuples or keywords."
                (ex-info {:given color})
                (throw))))]
  (coerce-from-map :color colors if-color-not-found)
  (coerce-from-map :fill-background-color colors if-color-not-found)
  (coerce-from-map :fill-foreground-color colors if-color-not-found)
  (coerce-from-map :left-border-color colors if-color-not-found)
  (coerce-from-map :right-border-color colors if-color-not-found)
  (coerce-from-map :top-border-color colors if-color-not-found)
  (coerce-from-map :bottom-border-color colors if-color-not-found))

(defmethod coerce-to-obj :font
  [^XSSFWorkbook wb _ font-attrs]
  (do-set-all! (.createFont wb) font-attrs))

(defmethod coerce-to-obj :data-format
  [^XSSFWorkbook wb _ format]
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

(defn- coerce-nested-to-obj
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
  [^XSSFWorkbook workbook attrs]
  (let [attrs' (coerce-nested-to-obj workbook attrs)]
    (try
      (do-set-all! (.createCellStyle workbook) attrs')
      (catch Exception e
        (-> "Failed to create cell style."
            (ex-info {:raw-attributes attrs :built-attributes attrs'} e)
            (throw))))))

(def default-style
  "The default cell style."
  {:font {:font-height-in-points 10 :font-name "Arial"}})

(defn merge-all
  "Recursively merge maps which may be nested.
      (merge-nested {:foo {:a :b}} {:foo {:c :d}})
        ; => {:foo {:a :b, :c :d}}"
  [& maps]
  (letfn [(merge? [left right]
            "Either merge maps or select the latter non-map value."
            (if (and (map? left) (map? right))
              (merge-2 left right)
              right))
          (merge-2 [m1 m2]
            "Recursively merge the entries in two maps."
            (reduce #(trampoline merge-entry %1 %2) (or m1 {}) (seq m2)))
          (merge-entry [m e]
            "Merge the entry, e, into map m."
            (let [k (key e) v (val e)]
              (if (contains? m k)
                #(assoc m k (merge? (get m k) v))
                (assoc m k v))))]
    (reduce merge-2 maps)))

(defn merge-style
  "Recursively merge the cell's current style with the provided style map,
  preserving any style that does not conflict."
  [cell style]
  (update
    (if (map? cell) cell {:value cell})
    :style (fn [s] (if-not s style (merge-all s style)))))

;;; Default table formatting functions to produce styles

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
    {:border-bottom :thin :font {:bold true}}))

;;; Default tree formatting functions to produce styles

(def default-tree-formatters
  {0 {:font {:bold true} :border-bottom :medium}
   1 {:font {:bold true}}
   2 {:indention 2}
   3 {:font {:italic true} :alignment :right}})

(def default-tree-total-formatters
  {0 {:font {:bold true} :border-top :medium}
   1 {:border-top :thin :border-bottom :thin}})

(defn example
  "If one wanted to visualize all of the nested setters & POI objects...
  Keep in mind that this requires $ apt-get install graphviz"
  []
  (let [param-type (fn [setter] (resolve (first (:parameter-types setter))))
        is-setter? (fn [{:keys [name parameter-types]}]
                     (and (string/starts-with? (str name) "set")
                          (= 1 (count parameter-types))))
        setters (fn [class]
                  (filter is-setter? (#'reflect/declared-methods class)))
        cell-style (first
                     (filter #(= 'setCellStyle (:name %)) (setters XSSFCell)))]
    (viz/view-tree
      #(instance? Class (param-type %)) (comp setters param-type) cell-style
      :node->descriptor #(->{:label ((juxt :name :parameter-types) %)}))))
