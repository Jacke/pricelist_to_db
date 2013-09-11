(ns zx.core
  (:use
   [clojure.java.io :only [input-stream]])
  (:import
   (org.xml.sax InputSource)
   (org.xml.sax.helpers XMLReaderFactory)
   (org.apache.poi POIXMLDocument)
   (org.apache.poi.openxml4j.opc OPCPackage)
   (org.apache.poi.poifs.filesystem POIFSFileSystem)
   (org.apache.poi.hssf.eventusermodel HSSFEventFactory
                                       HSSFListener
                                       HSSFRequest)
   (org.apache.poi.hssf.record Record)
   (org.apache.poi.xssf.eventusermodel ReadOnlySharedStringsTable
                                       XSSFReader
                                       XSSFSheetXMLHandler
                                       XSSFSheetXMLHandler$SheetContentsHandler))
  (:require [clojure.core.async :refer :all])
  (:gen-class))



(defn- file-type
  [istream opt]
  (when (.markSupported istream)
    (cond 
     (POIFSFileSystem/hasPOIFSHeader istream) :hssf
     (POIXMLDocument/hasOOXMLHeader istream)  :xssf)))


(defn- make-hssf-listener
  [handler]
  (reify 
    HSSFListener
    (processRecord
      [_ record]
      (println record))))


(defmulti read-xls
  "Reads Excel file"
  file-type)


;; binary Excel files (.xls)
(defmethod read-xls :hssf
  [istream handler]
  (let [events (new HSSFEventFactory)
        request (new HSSFRequest)]
    (.addListenerForAllRecords request (make-hssf-listener handler))
    (.processEvents events request istream)))


;; OOXML Excel files (.xlsx)
(defmethod read-xls :xssf
  [istream handler]
  (let [package (OPCPackage/open istream)
        reader (new XSSFReader package)
        styles (.getStylesTable reader)
        strings (new ReadOnlySharedStringsTable package)
        sheethandler (new XSSFSheetXMLHandler styles strings handler true)
        sheets (.getSheetsData reader)]
    (doseq [sheet (iterator-seq sheets)]
      (.startSheet handler (.getSheetName sheets))
      (doto (XMLReaderFactory/createXMLReader)
        (.setContentHandler sheethandler)
        (.parse (new InputSource sheet)))
      (.endSheet handler))))


(defprotocol SheetListener
  (startSheet [this name])
  (endSheet [this]))

(def template (ref {}))
(def foo (ref {}))
(def result (ref {}))
(def itid (atom 0))
; Simulate db call for retrive header 
(def template_header {"Заказ шт" "count", "ONPP" "onpp", "Автор" "author", "Название" "title", "Стд" "std", "Цена Опт.грн" "price", "Издательство" "publisher", "Пер." "translate", "Год" "year", "Жанр" "genre", "Серия" "series", "Формат" "format", "ISBN" "isbn", "Стр." "pages", "TOV_KOD" "article", "с НДС,без НДС" "pricetax", "EAN" "ean", "Категория" "category", "NAIMEN" "header"})

(defn fetch-header [cellref]
(= 1 (first (map #(java.lang.Integer/parseInt (% 0)) 
     (re-seq #"\d+(\.\d+)?" cellref)))))
; ["Person Id" "Name" "Height" "Email Address" "DOB"] : Compare row letter
; ["A*"        "B*"   "C*"     "D*"            "E*" ] : to template

(defn fetch-rows [cellref]
  (template_header                    ; type in db
     (@template (.charAt cellref 0))) ; type in sheet
                  

)

; Return a header

(def myhandler
  (reify
    XSSFSheetXMLHandler$SheetContentsHandler
    (cell
      [_ cellReference formattedValue]
      (dosync
        (if (fetch-header cellReference) 
          ;(ref-set template (conj @template formattedValue))
          (ref-set template (merge {(.charAt cellReference 0) formattedValue} @template))
          (ref-set foo (merge {(fetch-rows cellReference) formattedValue} @foo))))
      ;(println "#CellValue: " formattedValue)
      )
    (endRow
      [_]
      ;(println "#EndRow")
      (dosync (ref-set result (merge {@itid, @foo} @result))
        (swap! itid inc)))
    (headerFooter
      [_ text isHeader tagName]
      ;(println "headerFooter" text tagName)
      )
    
    (startRow
      [_ rowNum]
      (def foo (ref {})) ;(println "#RowNum: " rowNum)
      )
    SheetListener
    (startSheet
      [_ name] ;(println "#Sheet: " name)
    )
    (endSheet
      [_] ;(println "#EndSheet")
    )))
(def speedhandler
  (reify
    XSSFSheetXMLHandler$SheetContentsHandler
    (cell
      [_ cellReference formattedValue])
    (endRow
      [_])
    (headerFooter
      [_ text isHeader tagName])
    (startRow
      [_ rowNum])
    SheetListener
    (startSheet
      [_ name])
    (endSheet
      [_])))

(defn -main
  [fname]
  (with-open [istream (input-stream fname)]
    (read-xls istream myhandler)))
(defn speed [fname]
  (with-open [istream (input-stream fname)]
    (read-xls istream speedhandler)))