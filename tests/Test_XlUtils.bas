Attribute VB_Name = "Test_XlUtils"
Option Explicit
Option Base 0 ' Default

Private Sub test_all()
    gStart "XlUtils"
    
    test_rowColToExcel
    
    gStop
End Sub

Private Sub test_getOpenWorkbook()
    gStart "getOpenWorkbook"
    
    On Error Resume Next
    XlUtils.getOpenWorkbook "X:\non_existing_workbook.xlsx"
    checkError E_WORKBOOKNOTOPEN, "should throw on not opened workbook"
    On Error GoTo 0
    
    gStop
End Sub

Private Sub test_rowColToExcel()
    gStart "rowColToExcel"
    
    equals XlUtils.rowColToExcel(1, 1), "A1"
    equals XlUtils.rowColToExcel(1, 26), "Z1"
    equals XlUtils.rowColToExcel(1, 27), "AA1"
    equals XlUtils.rowColToExcel(20000, 1), "A20000"
    equals XlUtils.rowColToExcel(20000, 256), "IV20000"
    
    On Error Resume Next
    XlUtils.rowColToExcel 0, 1
    checkError E_INDEXOUTOFRANGE
    On Error GoTo 0

    On Error Resume Next
    XlUtils.rowColToExcel 1, 0
    checkError E_INDEXOUTOFRANGE
    On Error GoTo 0
    
    On Error Resume Next
    XlUtils.rowColToExcel -1, 1
    checkError E_INDEXOUTOFRANGE
    On Error GoTo 0

    On Error Resume Next
    XlUtils.rowColToExcel 1, -1
    checkError E_INDEXOUTOFRANGE
    On Error GoTo 0
    
    gStop
End Sub
