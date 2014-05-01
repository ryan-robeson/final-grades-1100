# final-grades-1100

A quick command line program to take grades exported and convert them to an Excel file that Sam will accept.

This script parses a CSV, concatenates the first and last name into one field, removes unnecessary columns, determines the letter grade from the final numeric grade,
and performs the attendance bump if the student had perfect attendance.

## Usage

### Export Grades

Grades -> Enter Grades -> Export

1. Select:
  * Org Defined ID
  * Points Grade (only)
  * Last name & First name
2. Select all
3. Uncheck: 
  * All "Subtotal"s
  * Perfect Attendance Check
  * Final Letter Grade
  * Final Adjusted Grade
4. Export to CSV
5. __Do not__ rename the CSV as it is used by the script to name the sheets in the workbook.

### Perform any necessary adjustments to script
* The relevant lines must be changed if the number of assignments or number of weeks differs from the script's expectations (23 HWs and 14 Weeks)

### Run
`lein deps`

`lein run CSCI*csv`

### Manual Fixes

A `semester-grades.xlsx` file has been created in the current directory. It has a sheet for each CSV file named after the section found in the 
CSV file's filename.

* Double Check grades
* Paste into Sam's provided template
* Format to Sam's example
* Include any other required columns (Semester, Section, etc.)

## License

Copyright © 2014 Ryan Robeson

Distributed under the Eclipse Public License either version 1.0 or (at
your option) any later version.
