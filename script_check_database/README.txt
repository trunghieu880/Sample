[Script use to check the path of each source C from database to compare with SUMMARY file]

Definition:
  + XLSX2CSV="/c/Python27/python.exe /c/Python27/Lib/site-packages/xlsx2csv.py "
  + SUMMARY_COEM="//hc-ut40346c/NHI5HC/hieunguyen/0000_Project/001_Prj/01_COEM/Summary_COEM.xlsx"

Option:
  + you_want_create_list_database_again : YES/NO
  + you_want_update_list_check_again : YES/NO

Library:
  + xlsx2csv.py
script_check_database.sh
Syntax: ./script_check_database.sh <your_input_path> | tee <yourfileoutput>

    Or option you_want_create_list_database_again = "NO":
        ./script_check_database.sh | tee <yourfileoutput>
