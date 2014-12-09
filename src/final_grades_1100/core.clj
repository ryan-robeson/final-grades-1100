(ns final-grades-1100.core
	(:gen-class)
	(:require [dk.ative.docjure.spreadsheet :refer :all]
		[clojure-csv.core :refer [parse-csv write-csv]]
		[clojure.tools.trace :refer [deftrace]]))

(defn round
	[s]
	(Math/round s))

(defn parseDouble
	[s]
	(cond
		(clojure.string/blank? s) 0
		:else (try (round (Double/parseDouble s))
			(catch NumberFormatException _ s))))

(defn sheet-name-from-path
	[p]
	(second (re-find #"-(\d{3}) -" p)))

(defn fix-name-header
	[header]
	(let
		[removed-name-columns
			(remove #{"First Name"}
				(remove #{"Last Name"}
					header))]
		(concat (list (first removed-name-columns)) (list "Name") (rest removed-name-columns))))

(defn fix-name-fields
	[row]
	(let [id (first row)
		  last-name (second row)
		  first-name (nth row 2)
		  full-name (format "%s, %s" last-name first-name)
		  r (drop 3 row)]
		  (concat (list id) (list full-name) r)))

(defn prepare-header
	[header]
	(drop 1 
		(drop-last 2 
			(fix-name-header
				(map 
					#(cond 
						(re-find #"(Final Exam.+) Points" %) (second (re-find #"(Final Exam.+) Points" %))
						(re-find #"(.+)\s-.+" %) (second (re-find #"(.+)\s-.+" %))
						:else %) header)))))

(defn add-letter-grade
	[row]
	(let [final-grade (last row)
		  letter-grade (cond
		  					(>= final-grade 93) "A"
		  					(>= final-grade 90) "A-"
		  					(>= final-grade 87) "B+"
		  					(>= final-grade 85) "B"
		  					(>= final-grade 83) "B-"
		  					(>= final-grade 80) "C+"
		  					(>= final-grade 77) "C"
		  					(>= final-grade 75) "C-"
		  					(>= final-grade 73) "D+"
		  					(>= final-grade 70) "D"
		  					:else "F")]
		  (concat row (list letter-grade))))

(defn attendance-bump
	[row attendance-start attendance-length]
	(let [letter-grade (last row)
		  perfect-attendance (not (some #(= 0 %) (take attendance-length (drop attendance-start row))))
		  new-grade (when perfect-attendance
		  				(cond
		  					(= "A-" letter-grade) "A"
		  					(= "B+" letter-grade) "A-"
		  					(= "B" letter-grade) "B+"
		  					(= "B-" letter-grade) "B"
		  					(= "C+" letter-grade) "B-"
		  					(= "C" letter-grade) "C+"
		  					(= "C-" letter-grade) "C"
		  					(= "D+" letter-grade) "D"
		  					(= "D" letter-grade) "D+"
		  					:else letter-grade))]
		  (if (and new-grade (not (= letter-grade new-grade)))
		  	(concat row (list new-grade))
		  	row)))

(defn prepare-fields
 [fields header]
 (let [attendance-start (first (keep-indexed #(when (re-matches #"WK01" %2) %1) header))
  attendance-length (reduce + (filter #(not (nil? %1)) (map #(when (re-find #"WK\d\d" %1) 1 ) header)))]
  (sort-by first
   (for [row fields]
    (attendance-bump (add-letter-grade (drop 1 (drop-last 2 (fix-name-fields (map #(parseDouble %) row)))))
     attendance-start
     attendance-length)))))

(defn parse-file
	[p]
	(let [data (parse-csv (slurp p))
		  header (prepare-header (first data))
		  fields (prepare-fields (rest data) header)]
		  (into [] (map vec (concat (list header) fields)))))

; Writes the given rows to a CSV file
; named after the section "016_grades.csv"
(defn to-csv-file!
  [section rows]
  (spit (str section "_grades.csv") (write-csv (map #(map #'str %1) rows))))

; Inserts grades in the given workbook.
; Also creates separate csv files for each
; section for further analysis outside of Excel.
(defn insert-grades
	[workbook p]
	(let [data (parse-file p)
      section (sheet-name-from-path p)
		  sheet (add-sheet! workbook section)]
		  (add-rows! sheet data)
      (to-csv-file! section data)))

(defn generate-grades-file
	[paths]
	(let [workbook (create-workbook "Sheet1" nil)]
		(doseq [p paths] (insert-grades workbook p))
		(save-workbook! "semester-grades.xlsx" workbook)))

; https://github.com/mjul/docjure
(defn -main
  [& args]
  (generate-grades-file args))
