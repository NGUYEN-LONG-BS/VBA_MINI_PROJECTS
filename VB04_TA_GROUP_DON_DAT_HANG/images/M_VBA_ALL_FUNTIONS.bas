Attribute VB_Name = "M_VBA_ALL_FUNTIONS"
Option Explicit
'====================================================================================================================================================
'====================================...Variable Declare
'====================================================================================================================================================
Public VAR_NAM_TAI_CHINH As String
'====================================================================================================================================================
'====================================...Function F_FIND_ANY_RANGE_IN_ANY_SHEET_01
'====================================================================================================================================================
Sub S_TEST_FUNCTION_F_FIND_ANY_RANGE_IN_ANY_SHEET_01()
    Dim range1 As Range
    Set range1 = F_FIND_ANY_RANGE_IN_ANY_SHEET_01("RANGE_COMBOBOX_TEN_KHACH_HANG", "SH_RANGE_MA_KH_01", True)
End Sub
Function F_FIND_ANY_RANGE_IN_ANY_SHEET_01(RangeNameString As String, SheetNameString As String, GetTitle As Boolean) As Range
    Dim OutputRange_with_header As Range
    Dim OutputRange_no_header As Range
    Dim FirstCellsRow As Long
    Dim FirstCellsCol As Long
    Dim LastCellsRow As Long
    Dim LastCellsCol As Long
    Dim RangeNameRow As Long
    Dim RangeNameCol As Long
    
    Dim ws As Worksheet
    Dim sheet_index As Integer
    For Each ws In ThisWorkbook.Worksheets
         If ws.CodeName = SheetNameString Then
            sheet_index = ws.Index
            Exit For
         End If
    Next ws
    
'    Debug.Print sheet_index
'    Debug.Print Worksheets(1).Name
'    Debug.Print Worksheets(sheet_index).Name
    
    FirstCellsCol = Application.WorksheetFunction.Match(RangeNameString, Worksheets(sheet_index).Rows(1), 0)
    FirstCellsRow = 3
    
    Set OutputRange_with_header = Worksheets(sheet_index).Cells(FirstCellsRow, FirstCellsCol).CurrentRegion
    
    LastCellsRow = OutputRange_with_header.Rows.count
    LastCellsCol = OutputRange_with_header.Columns.count
    
'    Debug.Print Worksheets(sheet_index).Name
'    Debug.Print FirstCellsRow
'    Debug.Print FirstCellsCol
'    Debug.Print LastCellsRow
'    Debug.Print LastCellsCol
    
    'way 1: VBA ==> bi loi: Application-denied or object-denied error 1004
    'Giai thich: chi hieu qua voi sheet active, neu thuc hien tren sheet khong active se bi loi nhu tren
    'Sheet du lieu phai activate moi khong bi loi
    ThisWorkbook.Worksheets(sheet_index).Activate
    Set OutputRange_no_header = Worksheets(sheet_index).Range(Cells(FirstCellsRow, FirstCellsCol).Offset(1, 0), Cells(FirstCellsRow, FirstCellsCol).Offset(LastCellsRow, LastCellsCol))
    'way 2: VBA ==> bi loi: Application-denied or object-denied error 1004
    'Giai thich: chi hieu qua voi sheet active, neu thuc hien tren sheet khong active se bi loi nhu tren
'    Set OutputRange_no_header = OutputRange_with_header.Range(Cells(FirstCellsRow, FirstCellsCol).Offset(1, 0), Cells(FirstCellsRow, FirstCellsCol).Offset(LastCellsRow - 1, LastCellsCol))
    'way 3: has one empty row at the end
'    Set OutputRange_no_header = OutputRange_with_header.Offset(1, 0)
        
    If GetTitle = True Then
        Set F_FIND_ANY_RANGE_IN_ANY_SHEET_01 = OutputRange_with_header
    Else
        Set F_FIND_ANY_RANGE_IN_ANY_SHEET_01 = OutputRange_no_header
    End If
    
'    Sheet8.Range("A1").Activate
'    Range(ActiveCell.Offset(1, 1), ActiveCell.Offset(5, 2)).Select
    
    RangeNameRow = Worksheets(sheet_index).Cells(FirstCellsRow, FirstCellsCol).CurrentRegion.Rows.count
    RangeNameCol = Worksheets(sheet_index).Cells(FirstCellsRow, FirstCellsCol).CurrentRegion.Columns.count
    
'    Debug.Print OutputRange.Rows.Count
'    Debug.Print OutputRange.Columns.Count
'    Debug.Print OutputRange.Address
'    OutputRange.Select
End Function
'====================================================================================================================================================
'====================================...Function F_GET_RANGE
'====================================================================================================================================================
Private Function F_GET_RANGE() As Range
    Set F_GET_RANGE = F_FIND_ANY_RANGE_IN_ANY_SHEET_01("RANGE_LISTBOX_DON_DAT_HANG", "SH_VT01_LISTBOX_DON_DAT_HANG", True)
    Set F_GET_RANGE = F_GET_RANGE.Offset(1).Resize(F_GET_RANGE.Rows.count - 1)
End Function
'====================================================================================================================================================
'====================================...Function F_GET_FILE_NAME_FROM_PATH
'====================================================================================================================================================
Function F_GET_FILE_NAME_FROM_PATH(path)
    Dim fileName As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    F_GET_FILE_NAME_FROM_PATH = fso.GetFilename(path)
End Function
'====================================================================================================================================================
'====================================...S_CLOSE_THISWORKBOOK_WITHOUT_SAVE
'====================================================================================================================================================
Sub S_CLOSE_THISWORKBOOK_WITHOUT_SAVE()
    If MsgBox("Do you want to logout?", vbYesNo, "Logout?") = vbYes Then
        Dim a As Integer
        a = Workbooks.count
        Debug.Print a
        If a = 1 Then
            ThisWorkbook.Close SaveChanges:=False
            Application.Quit
        ElseIf a > 1 Then
            ThisWorkbook.Close SaveChanges:=False
        End If
    End If
End Sub
