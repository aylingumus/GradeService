# GradeService
It's a windows service application which converts .cvs files to sqlite database files.

First this service creates a folder named as Grades. After creation a formatted .csv file in Grades folder, application converts that .csv file to a datatable. After convertion it saves datatable to SQLite database on Grades folder. When SQLite database creation is finished succesfully, deletes .csv file from Grades folder.
