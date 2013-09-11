(defproject zx "0.1.0-SNAPSHOT"
  :description "Apache POI wrapper for streaming Excel spreadsheet content"
  :url "http://github.com/atihonruk/clj-xlstream"
  :license {:name "Eclipse Public License"
            :url "http://www.eclipse.org/legal/epl-v10.html"}
  :repositories {"sonatype-oss-public" "https://oss.sonatype.org/content/groups/public/"}
  :dependencies [[org.clojure/clojure "1.5.1"]
                 [org.clojure/core.async "0.1.0-SNAPSHOT"]
                 [org.apache.poi/poi "3.9"]
                 [org.apache.poi/poi-ooxml "3.9"]]
  :main zx.core)
