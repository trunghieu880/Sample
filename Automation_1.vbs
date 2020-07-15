'Attribute VB_Name = "cmd"
Rem ------------------------------------
Const STUB_NAME = "AMSTB_SrcFile.c"
Const EXTENSION_XLSX = ".xlsx"
Const EXTENSION_CSV = ".csv"
Const EXTENSION_HTML = ".html"
Const EXTENSION_IE = "_IE" & EXTENSION_HTML
Const EXTENSION_OE = "_OE" & EXTENSION_HTML
Const EXTENSION_IO = "_IO" & EXTENSION_HTML
Const EXTENSION_TABLE = "_Table" & EXTENSION_HTML
Const EXTENSION_COV = ".txt"
Const RANGE_TB_1_2 = "A1"
Const VBA_OK = "vba_ok"
Rem update
Const SHEET_NAME_HISTORY = "A3"
Const SHEET_NAME_RESULT = "A4"
Const SHEET_NAME_IE = "A5"
Const SHEET_NAME_OE = "A6"
Const SHEET_NAME_IO = "A7"
Const SHEET_NAME_TABLE = "A8"
Const SHEET_NAME_COV = "A2"


Const CONST_TABLE_1_3 = "A10"
Const CONST_TABLE_1_4 = "A11"
Const CONST_STUB_TABLE_1_4 = "A12"
Const CONST_NONSTUB_TABLE_1_4 = "A13"
Const CONST_FILE_STUB_1_5 = "A14"

Const CONST_STUB = "A16"
Const CONST_COV = "A17"
Rem ------------------------------------

Const CONST_LABEL_FAULT = "Fault"
Dim FSO As Object

Sub main()

a = Import_StubFile("vPMS_SNR_Scan2.xlsx", Get_string_of_range(SHEET_NAME_RESULT), "B40", "C:\Users\phihn.pn\Desktop\New folder\Task00589_Group2_Sonar_scan_F.c_PhiNguyen\’P‘ÌƒeƒXƒgŒ‹‰Ê\vPMS_SNR_Scan2\AMSTB_SrcFile.c")

'a = Clear_SHEET_NAME_RESULT_Table_1_3("vPMS_SSF_IlrScanR.xlsx")
'SHEET_NAME_HISTORY = Workbooks("vba.xlsm").Worksheets("tool").Cells(1, 1)
MsgBox (a)

End Sub


'--------------------------------------------Function-----------------------------------------
Rem ##################################################################################################################
Function Import_HTML_File(ByVal wb_name As String, ByVal sheet_name As String, ByVal cell As String, ByVal htmlFile As String) As String
Application.ScreenUpdating = False
        On Error GoTo ErrorHandler
        Dim msg_return As String
1:     For Each ws In Workbooks(wb_name).Worksheets
        ws_name = ws.name
        If (InStr(ws_name, sheet_name)) Then
        
          Set FSO = CreateObject("Scripting.FileSystemObject")
2:        Workbooks.Open Filename:=htmlFile, ReadOnly:=True
    
          WBname_HTML = FSO.getBaseName(htmlFile) & EXTENSION_HTML
          Workbooks(WBname_HTML).Worksheets(Mid(FSO.getBaseName(WBname_HTML), 1, 31)).Activate
    
          Workbooks(WBname_HTML).Worksheets(Mid(FSO.getBaseName(WBname_HTML), 1, 31)).Cells.Copy
          DoEvents
3:        Workbooks(wb_name).Worksheets(ws_name).Activate
          Workbooks(wb_name).Worksheets(ws_name).Cells.PasteSpecial xlPasteAll
    
          Call Wrapping_off
          Application.CutCopyMode = False
          Application.DisplayAlerts = False
          Workbooks(WBname_HTML).Close SaveChanges:=False
          Set FSO = Nothing
    
          Exit For
        End If
      Next ws
        Call ctrl_home(wb_name, ws_name)
        If (msg_return = "") Then msg_return = VBA_OK
100:   Import_HTML_File = msg_return
        Exit Function
  
ErrorHandler:
    If (Erl = 1) Then
      msg_return = "Can't Activate " + wb_name + " || err.Description: " + Err.Description
      GoTo 100
    ElseIf (Erl = 2) Then
      msg_return = "Can't open html file " + " || err.Description: " + Err.Description
      GoTo 100
    ElseIf (Erl = 3) Then
      msg_return = "Can't Activate " + ws_name + " || err.Description: " + Err.Description
      GoTo 100
    Else
      msg_return = msg_return + "Other Error" + " || err.Description: " + Err.Description + ". "
      Resume Next
    End If
End Function
Rem ##################################################################################################################
Function Fmt_correction(ByVal wb_name As String, ByVal ws_name As String, ByVal Sum_sheet_name As String) As String
Application.ScreenUpdating = False
    On Error GoTo ErrorHandler
    Dim msg_return As String

1:  Workbooks(wb_name).Worksheets(ws_name).Activate
2:  Call Line_correction(Sum_sheet_name)
3:  Call Merge_cells(Sum_sheet_name)
4:  Call Fill_fault_sheet_tb(wb_name, ws_name, Sum_sheet_name)
5:  Call Autofix_format(wb_name, ws_name)
    Call ctrl_home(wb_name, ws_name)
    If (msg_return = "") Then msg_return = VBA_OK
100:   Fmt_correction = msg_return
    Exit Function
  
ErrorHandler:
    If (Erl = 1) Then
      msg_return = "Can't Activate " + wb_name + " OR " + ws_name + " || err.Description: " + Err.Description + ". "
      GoTo 100
    ElseIf (Erl = 2) Then
      msg_return = "Error Line_correction " + " || err.Description: " + Err.Description + ". "
      Resume Next
    ElseIf (Erl = 3) Then
      msg_return = msg_return + "Error Merge_cells " + " || err.Description: " + Err.Description + ". "
      Resume Next
    ElseIf (Erl = 4) Then
      msg_return = msg_return + "Error Fill_fault_sheet_tb " + " || err.Description: " + Err.Description + ". "
      Resume Next
    ElseIf (Erl = 5) Then
      msg_return = msg_return + "Error Autofix_format " + " || err.Description: " + Err.Description + ". "
      Resume Next
    Else
      msg_return = msg_return + "Other Error" + " || err.Description: " + Err.Description + ". "
      Resume Next
    End If

End Function
Rem ##################################################################################################################
Function Import_Coverage_File(ByVal wb_name As String, ByVal sheet_name As String, ByVal cell As String, ByVal fileCOV As String) As String
Application.ScreenUpdating = False
    On Error GoTo ErrorHandler
    Dim msg_return As String
    
1:  For Each ws In Workbooks(wb_name).Worksheets
        ws_name = ws.name
        If (InStr(ws_name, sheet_name)) Then
          Workbooks(wb_name).Activate
2:        Workbooks(wb_name).Worksheets(ws_name).Select
    
          Dim row_address_cov As Integer: row_address_cov = Workbooks(wb_name).Worksheets(ws_name).Range(cell).Row
          Dim col_address_cov As Integer: col_address_cov = Workbooks(wb_name).Worksheets(ws_name).Range(cell).Column
    
            Dim sInputRecord As String
            Dim fNum As Long
            fNum = FreeFile
3:          Open fileCOV For Input As #fNum
    
            Do While Not EOF(fNum)
              Line Input #fNum, sInputRecord
              Workbooks(wb_name).Worksheets(ws_name).Cells(row_address_cov, col_address_cov).Value = sInputRecord
              row_address_cov = row_address_cov + 1
            Loop
            Close #fNum
    
          Exit For
        End If
      Next ws
    Call ctrl_home(wb_name, ws_name)
    If (msg_return = "") Then msg_return = VBA_OK
100:   Import_Coverage_File = msg_return
  Exit Function
  
ErrorHandler:
    If (Erl = 1) Then
      msg_return = "Can't Activate " + wb_name + " || err.Description: " + Err.Description + ". "
      GoTo 100
    ElseIf (Erl = 2) Then
      msg_return = "Can't Activate " + ws_name + " || err.Description: " + Err.Description + ". "
      GoTo 100
    ElseIf (Erl = 3) Then
      msg_return = "Can't open cov file " + " || err.Description: " + Err.Description
      GoTo 100
    Else
      msg_return = msg_return + "Other Error" + " || err.Description: " + Err.Description + ". "
      Resume Next
    End If

End Function
Rem ##################################################################################################################

Function Import_StubFile(ByVal wb_name As String, ByVal sheet_name As String, ByVal cell As String, ByVal stubFile As String) As String
Application.ScreenUpdating = False
    On Error GoTo ErrorHandler
    Dim msg_return As String
    
1:    For Each ws In Workbooks(wb_name).Worksheets
          ws_name = ws.name
          If (InStr(ws_name, sheet_name)) Then
            Workbooks(wb_name).Activate
2:          Workbooks(wb_name).Worksheets(ws_name).Select
            
            If (InStr(Cells(Range(cell).Row - 1, Range(cell).Column), Get_string_of_range(CONST_STUB)) < 1) Then
                cell = Find_adrr_string(wb_name, sheet_name, Get_string_of_range(CONST_STUB)).Address
            End If
            Dim row_address_stub As Integer: row_address_stub = Range(cell).Row + 1
            Dim col_address_stub As Integer: col_address_stub = Range(cell).Column
            
            For Each ole In Workbooks(wb_name).Worksheets(ws_name).OLEObjects
                ole.Delete
            Next
            
            Workbooks(wb_name).Worksheets(ws_name).Cells(row_address_stub, col_address_stub).Select
            
            Dim ol As OLEObject
3:          Set ol = Workbooks(wb_name).Worksheets(ws_name).OLEObjects.Add(Filename:=stubFile, Link:=False, DisplayAsIcon:=True)
      
            ol.Top = Workbooks(wb_name).Worksheets(ws_name).Range(cell).Offset(1, 0).Top
            ol.Left = Workbooks(wb_name).Worksheets(ws_name).Range(cell).Offset(1, 0).Left
      
            Workbooks(wb_name).Worksheets(ws_name).Rows(row_address_stub).RowHeight = 42
            Workbooks(wb_name).Worksheets(ws_name).Cells(row_address_stub, col_address_stub) = ""
         
            Exit For
          End If
        Next ws
        Call ctrl_home(wb_name, ws_name)
        If (msg_return = "") Then msg_return = VBA_OK
100:   Import_StubFile = msg_return
        Exit Function
  
ErrorHandler:
    If (Erl = 1) Then
      msg_return = "Can't Activate " + wb_name + " || err.Description: " + Err.Description + ". "
      GoTo 100
    ElseIf (Erl = 2) Then
      msg_return = "Can't Activate " + ws_name + " || err.Description: " + Err.Description + ". "
      GoTo 100
    ElseIf (Erl = 3) Then
      msg_return = "Can't open stub file " + " || err.Description: " + Err.Description
      GoTo 100
    Else
      msg_return = msg_return + "Other Error" + " || err.Description: " + Err.Description + ". "
      Resume Next
    End If
    
End Function
Rem ##################################################################################################################
Function Fill_Table_1_4(ByVal wb_name As String, ByVal sheet_name As String, ByVal cell As String, ParamArray stub_A() As Variant) As String
Application.ScreenUpdating = False

    On Error GoTo ErrorHandler
    Dim msg_return As String
    Dim outCol As Long
    Dim outRow As Long
          
      For Each ws In Workbooks(wb_name).Worksheets
        ws_name = ws.name
        If (InStr(ws_name, sheet_name)) Then
          Workbooks(wb_name).Activate
          Workbooks(wb_name).Worksheets(ws_name).Select
    
            If (InStr(Cells(Range(cell).Row - 1, Range(cell).Column), Get_string_of_range(CONST_STUB_TABLE_1_4)) < 1) Then
                cell = Find_adrr_string(wb_name, sheet_name, Get_string_of_range(CONST_STUB_TABLE_1_4)).Address
            End If
            Dim row_stub As Integer: row_stub = Range(cell).Row + 1
            Dim col_stub As Integer: col_stub = Range(cell).Column - 1
    
    
            Dim count_stub As Integer: count_stub = 0
            For count_stub = 0 To UBound(stub_A(0))
              If count_stub > 0 Then
                  Workbooks(wb_name).Worksheets(ws_name).Rows(row_stub + count_stub).Select
                  Selection.Insert Shift:=xlDown
              End If
              Workbooks(wb_name).Worksheets(ws_name).Cells(row_stub + count_stub, col_stub).Value = count_stub + 1
              Workbooks(wb_name).Worksheets(ws_name).Cells(row_stub + count_stub, col_stub + 1).Value = stub_A(0)(count_stub, 0)
              Workbooks(wb_name).Worksheets(ws_name).Cells(row_stub + count_stub, col_stub + 5).Value = stub_A(0)(count_stub, 1)
            Next count_stub
    
            Call Fill_Border_Table_1_4(wb_name, sheet_name, row_stub - 1, col_stub, row_stub + count_stub - 1, col_stub + 8)
    
            Workbooks(wb_name).Worksheets(ws_name).Range("A1").Select
            Debug.Print ("Complete fill table 1.4 for sheet " & ws_name)
    
          Exit For
        End If
      Next ws
        Call ctrl_home(wb_name, ws_name)
        If (msg_return = "") Then msg_return = VBA_OK
100:   Fill_Table_1_4 = msg_return
        Exit Function
ErrorHandler:
        msg_return = msg_return + "Other Error" + " || err.Description: " + Err.Description + ". "
        Resume Next

End Function
Rem ##################################################################################################################
Function Fill_Table_1_3(ByVal wb_name As String, ByVal sheet_name As String, ByVal cell As String, ParamArray init_A() As Variant) As String
Application.ScreenUpdating = False
    On Error GoTo ErrorHandler
    Dim msg_return As String

      For Each ws In Workbooks(wb_name).Worksheets
        ws_name = ws.name
        If (InStr(ws_name, sheet_name)) Then
          Workbooks(wb_name).Activate
          Workbooks(wb_name).Worksheets(ws_name).Select
    
            If (InStr(Cells(Range(cell).Row - 1, Range(cell).Column), Get_string_of_range(CONST_TABLE_1_3)) < 1) Then
                cell = Find_adrr_string(wb_name, sheet_name, Get_string_of_range(CONST_TABLE_1_3)).Address
            End If
            Dim row_table_1_3 As Integer: row_table_1_3 = Range(cell).Row + 2
            Dim col_table_1_3 As Integer: col_table_1_3 = Range(cell).Column
    
            For i = 0 To UBound(init_A(0))
              If (count_variable > 0) Then
                  Workbooks(wb_name).Worksheets(ws_name).Rows(row_table_1_3 + count_variable).Select
                  Selection.Insert Shift:=xlDown
              End If
              Workbooks(wb_name).Worksheets(ws_name).Cells(row_table_1_3 + count_variable, col_table_1_3).Value = init_A(0)(i, 0)
              Workbooks(wb_name).Worksheets(ws_name).Cells(row_table_1_3 + count_variable, col_table_1_3 + 5).Value = init_A(0)(i, 1)
              count_variable = count_variable + 1
            Next i
    
              row_start = row_table_1_3 - 1
              col_start = col_table_1_3
              row_end = row_table_1_3 + count_variable - 1
              col_end = col_table_1_3 + 5
              Call Fill_Border_Table_1_3(wb_name, sheet_name, row_start, col_start, row_end, col_end)
              Workbooks(wb_name).Worksheets(ws_name).Range("A1").Select
              Debug.Print ("Complete fill table 1.3 for sheet " & ws_name)
              Exit For
    
        End If
      Next ws
        Call ctrl_home(wb_name, ws_name)
        If (msg_return = "") Then msg_return = VBA_OK
100:   Fill_Table_1_3 = msg_return
        Exit Function
ErrorHandler:
        msg_return = msg_return + "Other Error" + " || err.Description: " + Err.Description + ". "
        Resume Next
End Function
Rem ##################################################################################################################
Function Fill_Table_1_2(ByVal wb_name As String, ByVal sheet_name As String, ByVal cell As String, ByVal point_num As String, ParamArray init_A() As Variant) As String
Application.ScreenUpdating = False
    On Error GoTo ErrorHandler
    Dim msg_return As String
  If cell = "" Then
    cell = "B17"
  End If
  For Each ws In Workbooks(wb_name).Worksheets
    ws_name = ws.name
    If (InStr(ws_name, sheet_name)) Then
      Workbooks(wb_name).Activate
      Workbooks(wb_name).Worksheets(ws_name).Select

        Dim row_table_1_2 As Integer: row_table_1_2 = Workbooks(wb_name).Worksheets(ws_name).Range(cell).Row + 2
        Dim col_table_1_2 As Integer: col_table_1_2 = Workbooks(wb_name).Worksheets(ws_name).Range(cell).Column

        Dim count As Integer: count = 0
        Do While count < 1000
          If InStr(Cells(row_table_1_2, col_table_1_2), point_num) <> 0 Then
            Dim count_variable As Integer: count_variable = 0
            Workbooks("vba.xlsm").Activate
            Workbooks("vba.xlsm").Worksheets("tool").Select
            text_temp = Get_string_of_range(RANGE_TB_1_2)
            Workbooks(wb_name).Activate
            Workbooks(wb_name).Worksheets(ws_name).Select
            Workbooks(wb_name).Worksheets(ws_name).Cells(row_table_1_2, col_table_1_2 + 5).Value = text_temp
            For i = 0 To UBound(init_A(0))
              If (count_variable > 0) Then
                  Workbooks(wb_name).Worksheets(ws_name).Rows(row_table_1_2 + count_variable).Select
                  Selection.Insert Shift:=xlDown
              End If
              Workbooks(wb_name).Worksheets(ws_name).Cells(row_table_1_2 + count_variable, col_table_1_2 + 6).Value = init_A(0)(i, 0)
              Workbooks(wb_name).Worksheets(ws_name).Cells(row_table_1_2 + count_variable, col_table_1_2 + 6).HorizontalAlignment = xlLeft
              Workbooks(wb_name).Worksheets(ws_name).Cells(row_table_1_2 + count_variable, col_table_1_2 + 7).Value = init_A(0)(i, 1)
              Workbooks(wb_name).Worksheets(ws_name).Cells(row_table_1_2 + count_variable, col_table_1_2 + 8).Value = "-"
              count_variable = count_variable + 1
            Next i
            Call merge_cell_tb_1_2(row_table_1_2, col_table_1_2, row_table_1_2 + count_variable - 1, col_table_1_2 + 5)
            Debug.Print ("Complete fill " & point_num & " table 1.2 for sheet " & ws_name)
            Exit Do
          End If
          row_table_1_2 = row_table_1_2 + 1
        count = count + 1
        Loop
        Workbooks(wb_name).Worksheets(ws_name).Range("A1").Select
      Exit For

    End If
  Next ws
        Call ctrl_home(wb_name, ws_name)
        If (msg_return = "") Then msg_return = VBA_OK
100:   Fill_Table_1_2 = msg_return
        Exit Function
ErrorHandler:
        msg_return = msg_return + "Other Error" + " || err.Description: " + Err.Description + ". "
        Resume Next
End Function
Rem ##################################################################################################################
Function Fill_cell(ByVal wb_name As String, ByVal sheet_name As String, ByVal cell As String, ByVal text_s As String) As String
Application.ScreenUpdating = False
        On Error GoTo ErrorHandler
        Dim msg_return As String
        Workbooks(wb_name).Activate
        Workbooks(wb_name).Worksheets(sheet_name).Select
        Cells(Range(cell).Row, Range(cell).Column).Value = text_s
        
        Call ctrl_home(wb_name, sheet_name)
        If (msg_return = "") Then msg_return = VBA_OK
100:   Fill_cell = msg_return
        Exit Function
ErrorHandler:
        msg_return = msg_return + "Other Error" + " || err.Description: " + Err.Description + ". "
        Resume Next
End Function
Rem ##################################################################################################################
Function Fill_fault_conditional_format_n_reset_pointer(ByVal wb_name) As String
Application.ScreenUpdating = False
        On Error GoTo ErrorHandler
        Dim msg_return As String
        For Each ws In Workbooks(wb_name).Worksheets
            ws_name = ws.name
            Workbooks(wb_name).Activate
            Workbooks(wb_name).Sheets(ws_name).Activate
        
            Workbooks(wb_name).Sheets(ws.name).Select
            ActiveWindow.Zoom = 100
            Call ctrl_home(wb_name, ws_name)
        Next ws
        Workbooks(wb_name).Sheets(1).Select
        If (msg_return = "") Then msg_return = VBA_OK
100:   Fill_fault_conditional_format_n_reset_pointer = msg_return
        Exit Function
ErrorHandler:
        msg_return = msg_return + "Other Error" + " || err.Description: " + Err.Description + ". "
        Resume Next
End Function

Rem ###################################################################################################################### 14/11/2019
Function Remove_Row_Description(ByVal wb_name As String, ByVal ws_name As String) As String
  Application.ScreenUpdating = False
  On Error GoTo ErrorHandler
  Dim msg_return As String
  Dim new_row As Long
  Dim new_col As Long
  Dim cell_first_row As Long
  Dim cell_first_col As Long
  Dim cell_last_row As Long
  Dim cell_last_col As Long

1:  Workbooks(wb_name).Sheets(ws_name).Activate
    new_row = -1
    new_col = -1
    Dim pat_descripton As String: pat_descripton = "Description"
    Dim index As Long
    For index = 1 To 30
    temp_str = Workbooks(wb_name).Worksheets(ws_name).Cells(index, 1)
    If InStr(temp_str, pat_descripton) Then
      new_row = index
      new_col = 1
      Exit For
    End If
    Next index
    
    If (new_row > 0 And new_col > 0) Then
    str_address = Workbooks(wb_name).Sheets(ws_name).Cells(new_row, new_col).Address(RowAbsolute:=True, ColumnAbsolute:=True)
  
    Set ma = Workbooks(wb_name).Worksheets(ws_name).Range(str_address).MergeArea
    If ma.Address <> str_address Then
      Dim new_addr As String: new_addr = Split(Replace(ma.Address, "$", ""), ":")(1)
      new_row = Workbooks(wb_name).Worksheets(ws_name).Range(new_addr).Row
      new_col = Workbooks(wb_name).Worksheets(ws_name).Range(new_addr).Column
    End If
    
    cell_first_row = new_row
    cell_first_col = new_col + 1
  
    Workbooks(wb_name).Worksheets(ws_name).Cells.SpecialCells(xlLastCell).Select
    cell_last_row = new_row
    cell_last_col = ActiveCell.Column
    Workbooks(wb_name).Worksheets(ws_name).Range(Workbooks(wb_name).Worksheets(ws_name).Cells(cell_first_row, cell_first_col), Workbooks(wb_name).Worksheets(ws_name).Cells(cell_last_row, cell_last_col)).Value = ""
    End If
        Call ctrl_home(wb_name, ws_name)
        If (msg_return = "") Then msg_return = VBA_OK
100:   Remove_Row_Description = msg_return
        Exit Function
  
ErrorHandler:
    If (Erl = 1) Then
      msg_return = "Can't Activate " + wb_name + " || err.Description: " + Err.Description
      GoTo 100
    Else
      msg_return = msg_return + "Other Error" + " || err.Description: " + Err.Description + ". "
      Resume Next
    End If

End Function
Rem ##################################################################################################################

Function Clear_Sheet_html(ByVal wb_name As String) As String
  Application.ScreenUpdating = False
  On Error GoTo ErrorHandler
  Dim msg_return As String
  Dim offset_sheet_cov As Integer: offset_sheet_cov = 3
  
  Workbooks(wb_name).Activate
  
1:  For Each ws In Workbooks(wb_name).Worksheets
    ws_name = ws.name
    If (InStr(ws_name, Get_string_of_range(SHEET_NAME_IE)) _
        Or InStr(ws_name, Get_string_of_range(SHEET_NAME_OE)) _
        Or InStr(ws_name, Get_string_of_range(SHEET_NAME_IO)) _
        Or InStr(ws_name, Get_string_of_range(SHEET_NAME_TABLE)) _
        Or InStr(ws_name, Get_string_of_range(SHEET_NAME_COV)) _
      ) Then
      Workbooks(wb_name).Worksheets(ws_name).Select
      If (InStr(ws_name, Get_string_of_range(SHEET_NAME_COV)) < 1) Then
        Cells.Select
        Selection.Delete Shift:=xlUp
      Else
        Dim rngAddr As Range
        Set rngAddr = Find_adrr_string(wb_name, ws_name, Get_string_of_range(CONST_COV))
        If rngAddr Is Nothing Then
          Debug.Print ("BUG: not found CONST_COV")
        Else
          Dim row_address_cov As Integer: row_address_cov = Workbooks(wb_name).Worksheets(ws_name).Range(rngAddr.Address).Row + offset_sheet_cov
          MAX_ROW = Workbooks(wb_name).Worksheets(ws_name).Range("A" & Workbooks(wb_name).Worksheets(ws_name).Rows.count).End(xlUp).Row
          If MAX_ROW > 1 Then
              Workbooks(wb_name).Worksheets(ws_name).Range("A" & row_address_cov & ":A" & MAX_ROW).Select
              Selection.Delete Shift:=xlUp
          End If
        End If
      End If
      Workbooks(wb_name).Worksheets(ws_name).Range("A1").Select
    End If
  Next ws
  
        If (msg_return = "") Then msg_return = VBA_OK
100:    Clear_Sheet_html = msg_return
        Exit Function

ErrorHandler:
    If (Erl = 1) Then
      msg_return = "Can't Activate " + wb_name + " || err.Description: " + Err.Description
      GoTo 100
    Else
      msg_return = msg_return + "Other Error" + " || err.Description: " + Err.Description + ". "
      Resume Next
    End If
End Function
Rem ##################################################################################################################

Function Clear_SHEET_NAME_RESULT_Table_1_4(ByVal wb_name As String) As String
  Application.ScreenUpdating = False
  On Error GoTo ErrorHandler
  Dim msg_return As String
  
1:  For Each ws In Workbooks(wb_name).Worksheets
    ws_name = ws.name
    If (InStr(ws_name, Get_string_of_range(SHEET_NAME_RESULT))) Then
      Workbooks(wb_name).Worksheets(ws_name).Select

      Dim rngAddr_stub As Range
      Set rngAddr_stub = Find_adrr_string(wb_name, ws_name, Get_string_of_range(CONST_STUB_TABLE_1_4))

      Dim rngAddr_nonstub As Range
      Set rngAddr_nonstub = Find_adrr_string(wb_name, ws_name, Get_string_of_range(CONST_NONSTUB_TABLE_1_4))

      If rngAddr_nonstub Is Nothing Then
        Debug.Print ("BUG: Not Found label nonstub at function Fill_Table_1_4")
      Else
    
          Dim row_stub As Integer: row_stub = Workbooks(wb_name).Worksheets(ws_name).Range(rngAddr_stub.Address).Row + 2
          Dim col_stub As Integer: col_stub = Workbooks(wb_name).Worksheets(ws_name).Range(rngAddr_stub.Address).Column - 1
          Dim row_nonstub As Integer: row_nonstub = Workbooks(wb_name).Worksheets(ws_name).Range(rngAddr_nonstub.Address).Row
          Dim col_nonstub As Integer: col_nonstub = Workbooks(wb_name).Worksheets(ws_name).Range(rngAddr_nonstub.Address).Column + 2
          
          If (row_stub <> row_nonstub) Then
            If (row_stub < row_nonstub) Then
                For i = row_stub To row_nonstub - 1
                    Workbooks(wb_name).Worksheets(ws_name).Rows(row_stub).Select
                    Selection.Delete Shift:=xlUp
                Next i
            Else
                Debug.Print ("BUG: ROW_STUB table 1.4")
            End If
          Else
            Debug.Print ("No need to clear table 1.4 stub")
          End If
          
          Set rngAddr_file_stub_1_5 = Find_adrr_string(wb_name, ws_name, Get_string_of_range(CONST_FILE_STUB_1_5))
    
          Set rngAddr_nonstub = Find_adrr_string(wb_name, ws_name, Get_string_of_range(CONST_NONSTUB_TABLE_1_4))
    
          Dim row_file_stub_1_5 As Integer: row_file_stub_1_5 = Range(rngAddr_file_stub_1_5.Address).Row - 2
          row_nonstub = Range(rngAddr_nonstub.Address).Row
          
          If (row_nonstub <> row_file_stub_1_5) Then
            If (row_nonstub < row_file_stub_1_5) Then
              For i = row_nonstub To row_file_stub_1_5 - 1
                Workbooks(wb_name).Worksheets(ws_name).Rows(row_nonstub + 1).Select
                Selection.Delete Shift:=xlUp
              Next i
            Else
              Debug.Print ("BUG ROW_STUB table 1.4")
            End If
          Else
            Debug.Print ("No need to clear table 1.4 nonstub")
          End If
    
          Workbooks(wb_name).Worksheets(ws_name).Cells(row_nonstub - 1, col_stub).Value = "1"
          Workbooks(wb_name).Worksheets(ws_name).Cells(row_nonstub - 1, col_stub + 1).Value = "-"
          Workbooks(wb_name).Worksheets(ws_name).Cells(row_nonstub - 1, col_stub + 1 + 4).Value = "-"
          Workbooks(wb_name).Worksheets(ws_name).Cells(row_nonstub - 1, col_stub + 1 + 4 + 2).Value = ""
    
          Rem Phi update
          Workbooks(wb_name).Worksheets(ws_name).Cells(row_nonstub, col_nonstub).Value = ""
          Workbooks(wb_name).Worksheets(ws_name).Cells(row_nonstub + 1, col_nonstub).Value = ""
    
          Workbooks(wb_name).Worksheets(ws_name).Range("A1").Select
          Debug.Print ("Complete clear table 1.4 sheet " & ws_name)
          Exit For
        End If
    End If
  Next ws
        If (msg_return = "") Then msg_return = VBA_OK
100:    Clear_SHEET_NAME_RESULT_Table_1_4 = msg_return
        Exit Function

ErrorHandler:
    If (Erl = 1) Then
      msg_return = "Can't Activate " + wb_name + " || err.Description: " + Err.Description
      GoTo 100
    Else
      msg_return = msg_return + "Other Error" + " || err.Description: " + Err.Description + ". "
      Resume Next
    End If

End Function
Rem ##################################################################################################################

Function Clear_SHEET_NAME_RESULT_Table_1_3(ByVal wb_name As String) As String
  Application.ScreenUpdating = False
  On Error GoTo ErrorHandler
  Dim msg_return As String
  
  Dim offset_template_tb_1_3 As Integer: offset_template_tb_1_3 = 3
  Dim offset_template_tb_1_4 As Integer: offset_template_tb_1_4 = 2
  Dim offset_template_inside_table_1_3_row As Integer: offset_template_inside_table_1_3_row = 1
  Dim offset_template_inside_table_1_3_col As Integer: offset_template_inside_table_1_3_col = 5
  
1:  For Each ws In Workbooks(wb_name).Worksheets
    ws_name = ws.name
    If (InStr(ws_name, Get_string_of_range(SHEET_NAME_RESULT))) Then
      Workbooks(wb_name).Worksheets(ws_name).Select

      Dim rngAddr_table_1_3 As Range
      Set rngAddr_table_1_3 = Find_adrr_string(wb_name, ws_name, Get_string_of_range(CONST_TABLE_1_3))
      Dim rngAddr_table_1_4 As Range
      Set rngAddr_table_1_4 = Find_adrr_string(wb_name, ws_name, Get_string_of_range(CONST_TABLE_1_4))
      
      Dim row_table_1_3 As Integer: row_table_1_3 = Workbooks(wb_name).Worksheets(ws_name).Range(rngAddr_table_1_3.Address).Row + offset_template_tb_1_3
      Dim col_table_1_3 As Integer: col_table_1_3 = Workbooks(wb_name).Worksheets(ws_name).Range(rngAddr_table_1_3.Address).Column
      Dim row_table_1_4 As Integer: row_table_1_4 = Workbooks(wb_name).Worksheets(ws_name).Range(rngAddr_table_1_4.Address).Row - offset_template_tb_1_4
      Dim col_table_1_4 As Integer: col_table_1_4 = Workbooks(wb_name).Worksheets(ws_name).Range(rngAddr_table_1_4.Address).Column
      
      If (row_table_1_3 <> row_table_1_4) Then
        If (row_table_1_3 < row_table_1_4) Then
            For i = row_table_1_3 To row_table_1_4 - 1
                Workbooks(wb_name).Worksheets(ws_name).Rows(row_table_1_3).Select
                Selection.Delete Shift:=xlUp
                Rem Debug.Print ("delete row " & i)
            Next i
        Else
            Debug.Print ("BUG: ROW_STUB table 1.3")
        End If
      Else
        Debug.Print ("No need to clear table 1.3 initial value")
      End If
      
      Workbooks(wb_name).Worksheets(ws_name).Cells(row_table_1_3 - offset_template_inside_table_1_3_row, col_table_1_3).Value = "-"
      Workbooks(wb_name).Worksheets(ws_name).Cells(row_table_1_3 - offset_template_inside_table_1_3_row, col_table_1_3 + offset_template_inside_table_1_3_col).Value = "-"

      Workbooks(wb_name).Worksheets(ws_name).Range("A1").Select
      Debug.Print ("Complete clear table 1.3 sheet " & ws_name)
      Exit For
    End If
  Next ws
        If (msg_return = "") Then msg_return = VBA_OK
100:    Clear_SHEET_NAME_RESULT_Table_1_3 = msg_return
        Exit Function

ErrorHandler:
    If (Erl = 1) Then
      msg_return = "Can't Activate " + wb_name + " || err.Description: " + Err.Description
      GoTo 100
    Else
      msg_return = msg_return + "Other Error" + " || err.Description: " + Err.Description + ". "
      Resume Next
    End If
    
End Function






Rem ##################################################################################################################
Private Function Get_string_of_range(ByVal name As String) As String

    Wb_Excel_active = ActiveWorkbook.name
    Ws_Excel_active = ActiveSheet.name
    
    Workbooks("vba.xlsm").Activate
    Workbooks("vba.xlsm").Worksheets("tool").Select
    Get_string_of_range = Workbooks("vba.xlsm").Worksheets("tool").Cells(Range(name).Row, Range(name).Column).Text
    
    Workbooks(Wb_Excel_active).Activate
    Workbooks(Wb_Excel_active).Worksheets(Ws_Excel_active).Select
    
End Function
Rem ##################################################################################################################

Private Function Find_adrr_string(ByVal wb_name As String, ByVal sheet_name As String, ByVal string_find As String) As Range

   Set Find_adrr_string = Workbooks(wb_name).Sheets(sheet_name).Cells.Find(What:=string_find, After:=Range("A65536"), LookIn:=xlValues, LookAt:=xlWhole, _
        SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False)

End Function
Rem ##################################################################################################################

Private Sub Fill_fault_sheet_tb(ByVal wb_name As String, ByVal sheet_name As String, ByVal Sum_sheet_name As String)

  If Sum_sheet_name = "TC" Then
    lRow = Workbooks(wb_name).Sheets(sheet_name).Range("A" & Sheets(sheet_name).Rows.count).End(xlUp).Row
    lCol = Workbooks(wb_name).Sheets(sheet_name).Cells(4, Columns.count).End(xlToLeft).Column
    
    Dim ADDRESS_LAST_CELL As String: ADDRESS_LAST_CELL = Workbooks(wb_name).Sheets(sheet_name).Cells(lRow, lCol).Address(RowAbsolute:=False, ColumnAbsolute:=False)
    ADDRESS_LAST_CELL = Workbooks(wb_name).Sheets(sheet_name).Cells(lRow, lCol + 3).Address(RowAbsolute:=False, ColumnAbsolute:=False)
    Dim ADDRESS_START_FORMAT_CELL As String: ADDRESS_START_FORMAT_CELL = Workbooks(wb_name).Sheets(sheet_name).Cells(5, lCol).Address(RowAbsolute:=False, ColumnAbsolute:=False)

    Workbooks(wb_name).Sheets(sheet_name).Range(ADDRESS_START_FORMAT_CELL & ":" & ADDRESS_LAST_CELL).Select
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, _
      Formula1:="=""Fault"""
    Selection.FormatConditions(Selection.FormatConditions.count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
      .PatternColorIndex = xlAutomatic
      .Color = 13408767
      .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
  End If
End Sub
Private Sub Autofix_format(ByVal wb_name As String, ByVal sheet_name As String)

    Windows(wb_name).Activate
    Workbooks(wb_name).Sheets(sheet_name).Select
    lRow = Workbooks(wb_name).Sheets(sheet_name).Range("A" & Sheets(sheet_name).Rows.count).End(xlUp).Row
    lCol = Workbooks(wb_name).Sheets(sheet_name).Cells(4, Columns.count).End(xlToLeft).Column
    Dim ADDRESS_LAST_CELL As String: ADDRESS_LAST_CELL = Workbooks(wb_name).Sheets(sheet_name).Cells(lRow, lCol).Address(RowAbsolute:=False, ColumnAbsolute:=False)

    Workbooks(wb_name).Sheets(sheet_name).Range("A1").Select
    Selection.WrapText = False
    Workbooks(wb_name).Sheets(sheet_name).Range("A3:" & ADDRESS_LAST_CELL).Select
    With Selection
      .WrapText = True
      .Columns.AutoFit
    End With
    Workbooks(wb_name).Sheets(sheet_name).Range("A1").Select
    Selection.WrapText = False
      
End Sub

Private Sub ctrl_home(ByVal wb_name As String, ByVal ws_name As String)
    Workbooks(wb_name).Sheets(ws_name).Range("A1").Select
    Workbooks(wb_name).Sheets(ws_name).UsedRange.SpecialCells (xlCellTypeLastCell)
    Application.Goto Reference:=Range("A1"), Scroll:=True
End Sub

          
Rem ##################################################################################################################

Rem ****
Private Sub merge_cell_tb_1_2(ByVal start_cell_convert_row As Integer, ByVal start_cell_convert_col As Integer, ByVal last_cell_addr_row As Integer, ByVal last_cell_addr_col As Integer)
  For j = start_cell_convert_col To last_cell_addr_col
    Dim temp_rng_addr_start As String: temp_rng_addr_start = Cells(start_cell_convert_row, j).Address(RowAbsolute:=False, ColumnAbsolute:=False)
    Dim temp_rng_addr_end As String: temp_rng_addr_end = Cells(last_cell_addr_row, j).Address(RowAbsolute:=False, ColumnAbsolute:=False)

    Dim temp_rng As String: temp_rng = temp_rng_addr_start & ":" & temp_rng_addr_end
    Dim temp_merge_rng As Range: Set temp_merge_rng = Application.Range(temp_rng)
    temp_merge_rng.Merge
  Next j
End Sub




Private Sub Fill_Border_Table_1_4(ByVal wb_name As String, ByVal sheet_name As String, ByVal row_start As Integer, ByVal col_start As Integer, ByVal row_end As Integer, ByVal col_end As Integer)
  For Each ws In Workbooks(wb_name).Worksheets
    ws_name = ws.name
    If (InStr(ws_name, sheet_name)) Then
      Workbooks(wb_name).Worksheets(ws_name).Select
      Workbooks(wb_name).Worksheets(ws_name).Range(Cells(row_start, col_start), Cells(row_end, col_end)).Select

      With Selection
        .WrapText = True
        .Rows.AutoFit
      End With

      Selection.Borders(xlDiagonalDown).LineStyle = xlNone
      Selection.Borders(xlDiagonalUp).LineStyle = xlNone
      With Selection.Borders(xlEdgeLeft)
          .LineStyle = xlContinuous
          .ColorIndex = 0
          .TintAndShade = 0
          .Weight = xlMedium
      End With
      With Selection.Borders(xlEdgeTop)
          .LineStyle = xlContinuous
          .ColorIndex = 0
          .TintAndShade = 0
          .Weight = xlMedium
      End With
      With Selection.Borders(xlEdgeBottom)
          .LineStyle = xlContinuous
          .ColorIndex = 0
          .TintAndShade = 0
          .Weight = xlMedium
      End With
      With Selection.Borders(xlEdgeRight)
          .LineStyle = xlContinuous
          .ColorIndex = 0
          .TintAndShade = 0
          .Weight = xlMedium
      End With
      With Selection.Borders(xlInsideVertical)
          .LineStyle = xlContinuous
          .ColorIndex = 0
          .TintAndShade = 0
          .Weight = xlThin
      End With
      With Selection.Borders(xlInsideHorizontal)
          .LineStyle = xlContinuous
          .ColorIndex = 0
          .TintAndShade = 0
          .Weight = xlThin
      End With

      start_cell_convert_row = row_start + 1
      start_cell_convert_col = col_start + 1
      last_cell_addr_row = row_end
      last_cell_addr_col = col_start + 1 + 3

      Workbooks(wb_name).Worksheets(ws_name).Range(Cells(start_cell_convert_row, start_cell_convert_col), Cells(last_cell_addr_row, last_cell_addr_col)).Select

      With Selection
          .HorizontalAlignment = xlLeft
          .VerticalAlignment = xlCenter
          .WrapText = True
      End With

      Selection.Borders(xlDiagonalDown).LineStyle = xlNone
      Selection.Borders(xlDiagonalUp).LineStyle = xlNone
      With Selection.Borders(xlEdgeLeft)
          .LineStyle = xlContinuous
          .ColorIndex = 0
          .TintAndShade = 0
          .Weight = xlThin
      End With
      With Selection.Borders(xlEdgeTop)
          .LineStyle = xlContinuous
          .ColorIndex = 0
          .TintAndShade = 0
          .Weight = xlThin
      End With
      With Selection.Borders(xlEdgeBottom)
          .LineStyle = xlContinuous
          .ColorIndex = 0
          .TintAndShade = 0
          .Weight = xlMedium
      End With
      With Selection.Borders(xlEdgeRight)
          .LineStyle = xlContinuous
          .ColorIndex = 0
          .TintAndShade = 0
          .Weight = xlThin
      End With
      Selection.Borders(xlInsideVertical).LineStyle = xlNone
      With Selection.Borders(xlInsideHorizontal)
          .LineStyle = xlContinuous
          .ColorIndex = 0
          .TintAndShade = 0
          .Weight = xlThin
      End With

      Call my_merge_cell(start_cell_convert_row, start_cell_convert_col, last_cell_addr_row, last_cell_addr_col)

      start_cell_convert_row = row_start + 1
      start_cell_convert_col = col_start + 4 + 1
      last_cell_addr_row = row_end
      last_cell_addr_col = col_start + 4 + 2

      Workbooks(wb_name).Worksheets(ws_name).Range(Cells(start_cell_convert_row, start_cell_convert_col), Cells(last_cell_addr_row, last_cell_addr_col)).Select

      With Selection
          .HorizontalAlignment = xlLeft
          .VerticalAlignment = xlCenter
          .WrapText = True
      End With

      Selection.Borders(xlDiagonalDown).LineStyle = xlNone
      Selection.Borders(xlDiagonalUp).LineStyle = xlNone
      With Selection.Borders(xlEdgeLeft)
          .LineStyle = xlContinuous
          .ColorIndex = 0
          .TintAndShade = 0
          .Weight = xlThin
      End With
      With Selection.Borders(xlEdgeTop)
          .LineStyle = xlContinuous
          .ColorIndex = 0
          .TintAndShade = 0
          .Weight = xlThin
      End With
      With Selection.Borders(xlEdgeBottom)
          .LineStyle = xlContinuous
          .ColorIndex = 0
          .TintAndShade = 0
          .Weight = xlMedium
      End With
      With Selection.Borders(xlEdgeRight)
          .LineStyle = xlContinuous
          .ColorIndex = 0
          .TintAndShade = 0
          .Weight = xlThin
      End With
      Selection.Borders(xlInsideVertical).LineStyle = xlNone
      With Selection.Borders(xlInsideHorizontal)
          .LineStyle = xlContinuous
          .ColorIndex = 0
          .TintAndShade = 0
          .Weight = xlThin
      End With

      Call my_merge_cell(start_cell_convert_row, start_cell_convert_col, last_cell_addr_row, last_cell_addr_col)

      start_cell_convert_row = row_start + 1
      start_cell_convert_col = col_end - 1
      last_cell_addr_row = row_end
      last_cell_addr_col = col_end

      Workbooks(wb_name).Worksheets(ws_name).Range(Cells(start_cell_convert_row, start_cell_convert_col), Cells(last_cell_addr_row, last_cell_addr_col)).Select

      Selection.Borders(xlDiagonalDown).LineStyle = xlNone
      Selection.Borders(xlDiagonalUp).LineStyle = xlNone
      With Selection.Borders(xlEdgeLeft)
          .LineStyle = xlContinuous
          .ColorIndex = 0
          .TintAndShade = 0
          .Weight = xlThin
      End With
      With Selection.Borders(xlEdgeTop)
          .LineStyle = xlContinuous
          .ColorIndex = 0
          .TintAndShade = 0
          .Weight = xlThin
      End With
      With Selection.Borders(xlEdgeBottom)
          .LineStyle = xlContinuous
          .ColorIndex = 0
          .TintAndShade = 0
          .Weight = xlMedium
      End With
      With Selection.Borders(xlEdgeRight)
          .LineStyle = xlContinuous
          .ColorIndex = 0
          .TintAndShade = 0
          .Weight = xlMedium
      End With
      Selection.Borders(xlInsideVertical).LineStyle = xlNone
      With Selection.Borders(xlInsideHorizontal)
          .LineStyle = xlContinuous
          .ColorIndex = 0
          .TintAndShade = 0
          .Weight = xlThin
      End With

      Call my_merge_cell(start_cell_convert_row, start_cell_convert_col, last_cell_addr_row, last_cell_addr_col)

      Dim temp_string As String: temp_string = row_start & ":" & row_end
      Workbooks(wb_name).Worksheets(ws_name).Rows(temp_string).Select
      With Selection
        .WrapText = True
        .EntireRow.AutoFit
      End With
      Workbooks(wb_name).Worksheets(ws_name).Range("A1").Select
      Exit For
    End If
  Next ws
End Sub


Private Sub Fill_Border_Table_1_3(ByVal wb_name As String, ByVal sheet_name As String, ByVal row_start As Integer, ByVal col_start As Integer, ByVal row_end As Integer, ByVal col_end As Integer)
  For Each ws In Workbooks(wb_name).Worksheets
    ws_name = ws.name
    If (InStr(ws_name, sheet_name)) Then
      Workbooks(wb_name).Worksheets(ws_name).Select

      Workbooks(wb_name).Worksheets(ws_name).Range(Cells(row_start, col_start), Cells(row_end, col_end)).Select

      Selection.Borders(xlDiagonalDown).LineStyle = xlNone
      Selection.Borders(xlDiagonalUp).LineStyle = xlNone
      With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
      End With
      With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
      End With
      With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
      End With
      With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
      End With
      With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
      End With
      With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
      End With

      start_cell_convert_row = row_start + 1
      start_cell_convert_col = col_start
      last_cell_addr_row = row_end
      last_cell_addr_col = col_end - 1

      Workbooks(wb_name).Worksheets(ws_name).Range(Cells(start_cell_convert_row, start_cell_convert_col), Cells(last_cell_addr_row, last_cell_addr_col)).Select

      With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .WrapText = True
      End With

      Selection.Borders(xlDiagonalDown).LineStyle = xlNone
      Selection.Borders(xlDiagonalUp).LineStyle = xlNone
      With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
      End With
      With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
      End With
      With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
      End With
      With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
      End With
      Selection.Borders(xlInsideVertical).LineStyle = xlNone
      With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
      End With

      start_cell_convert_row = row_start + 1
      start_cell_convert_col = col_start
      last_cell_addr_row = row_end
      last_cell_addr_col = col_end - 1

      Call my_merge_cell(start_cell_convert_row, start_cell_convert_col, last_cell_addr_row, last_cell_addr_col)

      Dim temp_string As String: temp_string = row_start & ":" & row_end
      Workbooks(wb_name).Worksheets(ws_name).Rows(temp_string).Select
      With Selection
        .WrapText = True
        .EntireRow.AutoFit
      End With
      Workbooks(wb_name).Worksheets(ws_name).Range("A1").Select
      Exit For
    End If
  Next ws
End Sub

Private Sub my_merge_cell(ByVal start_cell_convert_row As Integer, ByVal start_cell_convert_col As Integer, ByVal last_cell_addr_row As Integer, ByVal last_cell_addr_col As Integer)
  For index_row = start_cell_convert_row To last_cell_addr_row
    Dim temp_rng_addr_start As String: temp_rng_addr_start = Cells(index_row, start_cell_convert_col).Address(RowAbsolute:=False, ColumnAbsolute:=False)
    Dim temp_rng_addr_end As String: temp_rng_addr_end = Cells(index_row, last_cell_addr_col).Address(RowAbsolute:=False, ColumnAbsolute:=False)

    Dim temp_rng As String: temp_rng = temp_rng_addr_start & ":" & temp_rng_addr_end
    Dim temp_merge_rng As Range: Set temp_merge_rng = Application.Range(temp_rng)
    temp_merge_rng.Merge
  Next index_row
End Sub



Private Sub Wrapping_off()
  ActiveSheet.Range("A1").Select
  ActiveSheet.Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
  With Selection
    .VerticalAlignment = xlCenter
    .WrapText = False
    .Orientation = 0
    .AddIndent = False
    .IndentLevel = -1
    .ShrinkToFit = False
    .ReadingOrder = xlContext
  End With
  ActiveSheet.Range("A1").Select
End Sub

Private Sub Line_correction(ByVal Sum_sheet_name As String)
  Dim Cell_Base_Row As Long
  Dim Cell_Base_Column As Long
  Dim cell_last_row As Long
  Dim Cell_Last_Column As Long
  Dim Interior_Color As Long
  Dim White As Long
  Dim LightBlue As Long
  Dim Blue As Long
  Dim Green As Long
  Dim Orange As Long

  '16777164:light blue,13434828:green,16777215:White,52479:orange
  White = 16777215
  LightBlue = 16777164
  Blue = 16764057
  Green = 13434828 '6750105
  Orange = 52479
  Black = 14277081
    Dim new_row As Integer
    Dim new_col As Integer

    For i = 1 To 100
      If (Cells(i, 2).Interior.Color = 6750105 And Sum_sheet_name = "IE") Then
        new_row = i
        new_col = 1
        ActiveSheet.Cells(new_row, new_col).Select
        Exit For
      ElseIf (Cells(i, 2).Interior.Color = LightBlue And Sum_sheet_name = "TC") Then
        new_row = i
        new_col = 2
        ActiveSheet.Cells(new_row, new_col).Select
        Exit For
      End If
    Next i

  Cell_Base_Row = ActiveCell.Row
  Cell_Base_Column = ActiveCell.Column

  'Debug.Print (Cell_Base_Row)
  'Debug.Print (Cell_Base_Column)

  ActiveCell.SpecialCells(xlLastCell).Select
  cell_last_row = ActiveCell.Row
  Cell_Last_Column = ActiveCell.Column
  'Debug.Print (Cell_Last_Row)
  'Debug.Print (Cell_Last_Column)

  For i = Cell_Base_Row To cell_last_row
    ActiveSheet.Cells(i, Cell_Base_Column).Select
    Interior_Color = ActiveSheet.Cells(i, Cell_Base_Column).Interior.Color
    'Debug.Print (Interior_Color)
    If Interior_Color = LightBlue Or Interior_Color = Green Or Interior_Color = Blue Then
      'Debug.Print (Interior_Color)
    ElseIf Interior_Color = White Or Interior_Color = Orange Then
      'Debug.Print (Interior_Color)
      ActiveSheet.Range(ActiveSheet.Cells(i, Cell_Base_Column), ActiveSheet.Cells(i, Cell_Last_Column)).Cut
      ActiveSheet.Range(ActiveSheet.Cells(i, Cell_Base_Column), ActiveSheet.Cells(i, Cell_Last_Column)).Select
      Selection.Offset(0, 1).Select
      DoEvents
      ActiveSheet.Paste
      'Debug.Print (Interior_Color)
    Else
      'Debug.Print (Interior_Color)
      ActiveSheet.Range(ActiveSheet.Cells(i, Cell_Base_Column), ActiveSheet.Cells(i, Cell_Last_Column)).Cut
      ActiveSheet.Range(ActiveSheet.Cells(i, Cell_Base_Column), ActiveSheet.Cells(i, Cell_Last_Column)).Select
      If Sum_sheet_name = "IE" Then
        Selection.Offset(0, 2).Select
      Else
        Selection.Offset(0, 1).Select
      End If
      DoEvents
      ActiveSheet.Paste
    End If

  Next i

End Sub

Private Sub Merge_cells(ByVal Sum_sheet_name As String)
    Dim i As Long
    Dim Cell_Base_Row As Long
    Dim Cell_Base_Column As Long
    Dim cell_last_row As Long
    Dim Cell_Last_Column As Long
    Dim Interior_Color As Long
    Dim Next_InteriorColor As Long
    Dim Cell_Merge_Start As Long
    Dim Cell_Merge_End As Long
    Dim LightBlue As Long
    Dim Green As Long
    Dim Blue As Long
    Dim IE_Offset As Long
    Dim Table_Offset As Long

    '16777164:light blue,13434828:green,16777215:White,52479:orange
    LightBlue = 16777164
    Green = 13434828
    Blue = 16764057

    IE_Offset = 12
    Table_Offset = 4

    Dim new_row As Integer
    Dim new_col As Integer

    For i = 1 To 100
      If (Cells(i, 2).Interior.Color = 6750105 And Sum_sheet_name = "IE") Then
        new_row = i
        new_col = 1
        ActiveSheet.Cells(new_row, new_col).Select
        Exit For
      ElseIf (Cells(i, 2).Interior.Color = LightBlue And Sum_sheet_name = "TC") Then
        new_row = i
        new_col = 2
        ActiveSheet.Cells(new_row, new_col).Select
        Exit For
      End If
    Next i

    Cell_Base_Row = ActiveCell.Row
    Cell_Base_Column = ActiveCell.Column

    'Debug.Print (Cell_Base_Row)
    'Debug.Print (Cell_Base_Column)

    ActiveCell.SpecialCells(xlLastCell).Select
    cell_last_row = ActiveCell.Row
    Cell_Last_Column = ActiveCell.Column
    'Debug.Print (Cell_Last_Row)
    'Debug.Print (Cell_Last_Column)

    For i = Cell_Base_Row To cell_last_row - 1

        ActiveSheet.Cells(i, Cell_Base_Column).Select
        Interior_Color = ActiveSheet.Cells(i, Cell_Base_Column).Interior.Color
        Next_InteriorColor = ActiveSheet.Cells(i + 1, Cell_Base_Column).Interior.Color
        'Debug.Print (Interior_Color)

        '16777164:light blue,13434828:green,16777215:White
        If Interior_Color = LightBlue Or Interior_Color = Green Or Interior_Color = Blue Then
            'Debug.Print (Interior_Color)
            Cell_Merge_Start = i
        Else
            Cell_Merge_End = i
            'Debug.Print (Cell_Merge_End)

            If Next_InteriorColor = LightBlue Or Next_InteriorColor = Blue Or i = cell_last_row - 1 Then
              ActiveSheet.Range(ActiveSheet.Cells(Cell_Merge_Start, Cell_Base_Column), ActiveSheet.Cells(Cell_Merge_End, Cell_Base_Column)).Merge
              If Sum_sheet_name = "IE" Then
                ActiveSheet.Range(ActiveSheet.Cells(Cell_Merge_Start + 1, Cell_Base_Column + 1), ActiveSheet.Cells(Cell_Merge_End - 1, Cell_Base_Column + 1)).Merge


                ActiveSheet.Range(ActiveSheet.Cells(Cell_Merge_Start + 1, Cell_Base_Column + 1), ActiveSheet.Cells(Cell_Merge_End - 1, Cell_Base_Column + 1)).Select

                With Selection.Borders(xlEdgeLeft)
                  .LineStyle = xlContinuous
                  .Weight = xlThick
                End With
                With Selection.Borders(xlEdgeTop)
                  .LineStyle = xlDash
                  .Weight = xlMedium
                End With
                With Selection.Borders(xlEdgeBottom)
                  .LineStyle = xlDash
                  .Weight = xlMedium
                End With
                With Selection.Borders(xlEdgeRight)
                  .LineStyle = xlContinuous
                  .Weight = xlMedium
                End With
              End If
            End If
        End If

    Next i

    If Sum_sheet_name = "IE" Then
      ActiveSheet.Range(ActiveSheet.Cells(Cell_Base_Row, Cell_Base_Column), ActiveSheet.Cells(cell_last_row - 1, Cell_Base_Column)).Select
      With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThick
      End With
      With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThick
      End With
      With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlThick
      End With
      With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThick
      End With
    Else
      ActiveSheet.Range(ActiveSheet.Cells(Cell_Base_Row, Cell_Base_Column), ActiveSheet.Cells(cell_last_row - 1, Cell_Base_Column)).Select
      With Selection.Borders(xlEdgeTop)
        .LineStyle = xlDouble
        .Weight = xlThick
      End With
      With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThick
      End With
      With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThick
      End With
      With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlMedium
      End With
      With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThick
      End With
    End If

End Sub