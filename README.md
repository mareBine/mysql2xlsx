# MySQL to XLSX
#### a simple utility to export SQL query results directly to XLSX files (Python based)

### Features
* support for multiple worksheets, each sql query into single worksheet
* support for loops: define a MySQL variable (e.g. @var) and loop through values stored in an array (e.g. ['SI','AT'])
* support for xlsx validation: create xlsx files with specific fields 'locked' for predefined values

### Binary
* I have precompiled a windows binary (command line utility), which can be used straight away, usage:

`mysql2xlsx.exe example.sql`

### To run as an example
1. execute the `example_create_tables.sql` which will create 2 tables `tb_countries` and `bwd_status`
2. set your mysql server parameters in `xlsx.ini` file under `[default]` section
3. in windows cmd run: `mysql2xlsx.exe example.sql`
4. it should create 2 xlsx files `AT_example.sql.xlsx` and `SI_example.sql.xlsx` with 2 sheets each

### Installation (Windows) - if you wish to modify the code
* install python for windows: https://www.python.org/download
* install MySQL connector for python: https://pypi.python.org/pypi/MySQL-python/
* install pip - package installer for python: https://pip.pypa.io/en/latest/installing.html
    * Visual C++ compiler is a requirement: http://aka.ms/vcpython27
* install xlsxwriter package: C:\Python27\Scripts\pip install XlsxWriter

### TODO
* catch exception, if xls file already opened (permission denied)
* writes None into NULL fields, it would be neater if they were empty -> a quick fix is to have empty instead of NULL in tables
* split into more files if there are more than 1M rows (xlsx limitation)