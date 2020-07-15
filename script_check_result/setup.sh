#!/bin/bash

my_dir_script="$( cd "$( dirname "${BASH_SOURCE[0]}" )" >/dev/null 2>&1 && pwd )"
cd ${my_dir_script}

#----------------------------------------#
# PUT YOUR NAME IN PIC COLUMN
#----------------------------------------#
export PIC="[0-9]*[0-9],Phuoc,"
export PR_PIC="Hieu"

export SHEET_NAME="SHEET_NAME" # Modify HERE
export BUG_REPORT_NAME="${SHEET_NAME}_No"

#----------------------------------------#
# PUT YOUR TESTOUTPUT LOCATION
#----------------------------------------#
export SRC_RESULT="/d/My_Document/2_Project/Deliverables/Branches/Task00499_38"

#----------------------------------------#
# PUT YOUR INPUT LOCATION FROM NISSAN ADAS
#----------------------------------------#
export SRC_INPUT_ADAS="/d/0000_Project_AA_work/38_P33A_PT1_MRD_20191101"

#----------------------------------------#
# PUT YOUR Summary File
#----------------------------------------#
export SUMMARY_FILE="/c/Users/hieutnguyen.gl/Desktop/Summary_BB.xlsx" # Modify HERE
export GROUP_INDEX=1
export VERSION_RELEASE=",1$|,2$|,3$|,4$"

#----------------------------------------#
# INSTALL PYTHON 2.7
# COPY xlsx2csv.py from https://github.com/dilshod/xlsx2csv into FOLDER PYTHON
# THEN MODIFY IF YOUR LOCATION IS CHANGED
#----------------------------------------#
export XLSX2CSV="/c/Python27/python.exe /c/Python27/Lib/xlsx2csv.py "

export my_list_testcase="LIST"

export PRINTING=0 # MODIFY HERE IF YOU ONLY NEED NG STATUS (1: ONLY NG, 0: PRINT ALL)

FLAG_FIND_SRC_INPUT_ADAS=0 # 1 : FIND 0 : SKIP

#****************************************#
# FIND SUMMARY
#****************************************#
#----------------------------------------#
# PIC & PR_PIC        = 0
# VERSION RELEASE     = 1
# NOT RUN             = 999
#----------------------------------------#
#FLAG_FIND_SUMMARY=0
#FLAG_FIND_SUMMARY=1
FLAG_FIND_SUMMARY=999

#****************************************#
# FIND SRC RESULT
#****************************************#
#----------------------------------------#
# ALL                 = 0
# PIC                 = 1
# PIC & GROUP         = 2
#----------------------------------------#
#FLAG_FIND_SRC_RESULT=0
FLAG_FIND_SRC_RESULT=1

export FLAG_DEBUG=0 #1:ENABLE 0:SKIP

#----------------------------------------#
# START CHECKING
#----------------------------------------#
date
bash ${my_dir_script}/script_collect_result.sh ${FLAG_FIND_SRC_INPUT_ADAS} ${FLAG_FIND_SUMMARY} ${FLAG_FIND_SRC_RESULT}
date
