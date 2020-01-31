;; "resources/spreadsheet.xlsx" "Research themes" "Research programmes" "UKCRC Categories" "Staff categories" "People" "Projects" "Publications"
;;
;; psuedocode
;;
;; Prepare JDBC or similar connection
;; get db connection
;; Prepare SQL statements
;; prepare table name (trim and downcase maybe join with underscore/dash??)
;; prepare column name (trim and downcase maybe join with underscore/dash??)
;; first row is column header
;; insert rest of rows into table
;;
;; useful libraries from Red
;;
;; [clojure.java.jdbc :refer [with-db-transaction]]
;; [clojure.string :as str]
;; hugsql

(ns excel-to-sql.core
  (:require [dk.ative.docjure.spreadsheet :as spreadsheet]))

;; this function will read and return rows from a workbook sheet
;; load a workbook from the given workbook path
;; load a sheet by the given sheet name
;; and then return rows as a lazy sequence
(defn read-rows
  [workbook sheet]
  (->> (spreadsheet/load-workbook workbook)
    (spreadsheet/select-sheet sheet)
    spreadsheet/row-seq
    (map spreadsheet/cell-seq)
    (map #(map spreadsheet/read-cell %))))

;; this function will
;; insert a row into a table given a list of
;; column names and a list of values
(defn import-row
  [columns, values]
  ; import column and values
  (println "Importing row"))

; this function will
; read rows from a sheet
; prepare table name from sheet name
; prepare column names from header
; prepare insert statements for table, column and rows
; create table with columns
; insert rows into table
(defn import-sheet
  [workbook sheet]
  (def rows (read-rows workbook sheet))
  (def headers (first rows))
  (println (str "Importing sheet - " sheet " from workbook " workbook))
  (println "--Sheet headers--")
  (doseq [header headers]
   (println header))
  ; (doseq [row (rest rows)]
  ;  (import-row headers row))
  (println "\n"))

(defn -main
  "Import excel workbook into SQL database."
  [workbook & sheets]
  (if sheets
    (do (println "\n")
      (println "Importing excel workbook into database ...")
      (println (str (count sheets) " sheets provided."))
      (println "\n")
      (doseq [sheet sheets]
        (import-sheet workbook sheet)))
    (throw (Exception. "Need at least one command line argument!"))))
