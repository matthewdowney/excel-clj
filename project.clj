(defproject org.clojars.mjdowney/excel-clj "2.0.1"
  :description "Generate Excel documents & PDFs from Clojure data."
  :url "https://github.com/matthewdowney/excel-clj"
  :license {:name "Eclipse Public License"
            :url "http://www.eclipse.org/legal/epl-v10.html"}
  :dependencies [[org.clojure/clojure "1.10.3"]
                 [com.taoensso/encore "3.21.0"]
                 [com.taoensso/tufte "2.2.0"]
                 [org.apache.poi/poi-ooxml "5.2.0"]
                 [org.jodconverter/jodconverter-local "4.4.2"]]
  :profiles {:test {:dependencies [[org.apache.logging.log4j/log4j-core "2.17.1"]
                                   [org.slf4j/slf4j-nop "1.7.36"]]}})
