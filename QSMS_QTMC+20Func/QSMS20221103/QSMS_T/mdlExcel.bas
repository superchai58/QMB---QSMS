Attribute VB_Name = "mdlExcel"
Option Explicit


Function TransposeDim(V As Variant) As Variant
' Custom Function to Transpose a 0-based array (v)
    
    Dim X As Long, Y As Long, Xupper As Long, Yupper As Long
    Dim tempArray As Variant
    
    Xupper = UBound(V, 2)
    Yupper = UBound(V, 1)
    
    ReDim tempArray(Xupper, Yupper)
    For X = 0 To Xupper
        For Y = 0 To Yupper
            tempArray(X, Y) = V(Y, X)
        Next Y
    Next X
    
    TransposeDim = tempArray

End Function
Public Sub OutPutExcel(ByVal Rs As ADODB.Recordset, xlApplication As Excel.Application, SheetName As String)
Dim iCurRow As Long
Dim I As Long
Dim xlWorkSheet As Excel.Worksheet
'AddSheet
Set xlWorkSheet = xlApplication.Worksheets.Add
 xlWorkSheet.Name = SheetName
'Print Column Name
  iCurRow = 1
  For I = 0 To Rs.Fields.Count - 1
       xlWorkSheet.Cells(iCurRow, I + 1).Value = Rs.Fields(I).Name
  Next I
 While Rs.EOF = False
    '====Print Detail=======
    iCurRow = iCurRow + 1
    For I = 0 To Rs.Fields.Count - 1
        xlWorkSheet.Cells(iCurRow, I + 1).Value = Rs(I)
    Next I
    Rs.MoveNext
Wend
xlApplication.Selection.CurrentRegion.Columns.AutoFit
xlApplication.Selection.CurrentRegion.Rows.AutoFit
Set xlWorkSheet = Nothing
End Sub



Public Sub SaveToExcel(ByVal Rst As ADODB.Recordset, Path As String, FileName As String, SheetNO)
 Dim xlApp As Excel.Application
 Dim xlsBook As Excel.Workbook
 Dim xlWs As Object
 Dim recArray As Variant
 Dim strDB As String
 Dim fldCount As Long
 Dim recCount As Long
 Dim iCol As Long
 Dim iRow As Long
 Dim FilePath As String
    Set xlApp = CreateObject("Excel.Application")
    Set xlsBook = xlApp.Workbooks.Add
    'important for disabled alerts
    xlApp.DisplayAlerts = False
    Set xlWs = xlApp.Worksheets(SheetNO)
  
    xlApp.UserControl = True
    
    ' Copy field names to the fiRst row of the worksheet
    fldCount = Rst.Fields.Count
    For iCol = 1 To fldCount
        xlWs.Cells(1, iCol).Value = Rst.Fields(iCol - 1).Name
    Next
        
    ' Check veRsion of Excel
    If Val(Left$(xlApp.Version, 1)) > 8 Or Val(Left$(xlApp.Version, 2)) > 8 Then
'        'EXCEL 2000: Use CopyFromRecordset
'        ' Copy the recordset to the worksheet, starting in cell A2
        xlWs.Cells(2, 1).CopyFromRecordset Rst
'        'Note: CopyFromRecordset will fail if the recordset
'        'contains an OLE object field or array data such
'        'as hierarchical recordsets
'
    Else
        'EXCEL 97 or earlier: Use GetRows then copy array to Excel
        ' Copy recordset to an array
        recArray = Rst.GetRows
        'Note: GetRows returns a 0-based array where the fiRst
        'dimension contains fields and the second dimension
        'contains records. We will transpose this array so that
        'the fiRst dimension contains records, allowing the
        'data to appeaRs properly when copied to Excel
        
        ' Determine number of records
        recCount = UBound(recArray, 2) + 1 '+ 1 since 0-based array
        
        ' Check the array for contents that are not valid when
        ' copying the array to an Excel worksheet
        For iCol = 0 To fldCount - 1
            For iRow = 0 To recCount - 1
                ' Take care of Date fields
                If IsDate(recArray(iCol, iRow)) Then
                    recArray(iCol, iRow) = Format(recArray(iCol, iRow))
                ' Take care of OLE object fields or array fields
                ElseIf IsArray(recArray(iCol, iRow)) Then
                    recArray(iCol, iRow) = "Array Field"
                End If
            Next iRow 'next record
        Next iCol 'next field
            
        ' Transpose and Copy the array to the worksheet,
        ' starting in cell A2
        xlWs.Cells(2, 1).Resize(recCount, fldCount).Value = _
            TransposeDim(recArray)
    End If

    ' Auto-fit the column widths and row heights
    xlApp.Selection.CurrentRegion.Columns.AutoFit
    xlApp.Selection.CurrentRegion.Rows.AutoFit
    ' Save File  --moya
    FilePath = Path & "\" & FileName
    xlWs.SaveAs FilePath
    
   ' xlApp.Visible = True
    ' Close ADO objects
    Rst.Close
    
    xlApp.Quit
    Set Rst = Nothing
    
    ' Release Excel references
'    Set xlWs = Nothing
    Set xlApp = Nothing
    Set xlsBook = Nothing
    
  
End Sub


Public Function ExportGridToExcel(ObjectRst As ADODB.Recordset, ExcelFile As String, ExcelFileHead As String) As Long
       Dim RowCount As Long
       Dim ColCount As Long
       
       Dim objExcel As Excel.Application
       Dim objField As Field
       Dim objWorkbook As Excel.Workbook
       Dim objWorkSheet As Excel.Worksheet
       
       Screen.MousePointer = vbHourglass
       
       On Error GoTo ErrorProcessSection
       
       Set objExcel = New Excel.Application
       
       '------------------------------------------------
       ' A0 Excel 相關設定作業。
       '------------------------------------------------
       Set objExcel = New Excel.Application
    
       ' 不讓使用者操作。
       objExcel.Interactive = False

       ' 背後作業。
       If objExcel.Visible = Not False Then
          objExcel.Visible = Not True
       End If
    
       ' 視窗最大化。
       objExcel.WindowState = xlMaximized
    
       ' 設定 Wokkbook 物件。
       Set objWorkbook = objExcel.Workbooks.Add
    
       ' 設定 Worksheet 物件，指向 Sheet 1。
       Set objWorkSheet = objWorkbook.Worksheets.Add
       
       objWorkSheet.PageSetup.CenterHeader = ExcelFileHead
       objWorkSheet.PageSetup.LeftHeader = "Date/Time : " & Format(Now, "YYYY/MM/DD HH:NN:SS")
       
       '------------------------------------------------
       ' A1 Excel 表頭部份相關設定作業。
       '------------------------------------------------
       With objWorkSheet.Columns
            .Font.Size = 10
            .HorizontalAlignment = xlCenter
       End With
       
       ColCount = 1
       For Each objField In ObjectRst.Fields
           'Select Case ObjectRst.Type
              ' 下述資料型態則予以略過。
           '   Case adGUID, adLongVarBinary, adLongVarWChar
           '   Case Else
                objWorkSheet.Cells(1, ColCount).Value = objField.Name
                objWorkSheet.Cells(1, ColCount).Interior.ColorIndex = 33
                objWorkSheet.Cells(1, ColCount).Font.Bold = True
                objWorkSheet.Cells(1, ColCount).BorderAround xlContinuous
                ColCount = ColCount + 1
           'End Select
       Next objField
       
       '------------------------------------------------
       ' A2 Excel 表身部份相關設定作業。
       '------------------------------------------------
       ObjectRst.MoveFirst
       RowCount = 2
       Do
          Select Case ObjectRst.EOF()
             Case True
               Exit Do
             Case False
               ColCount = 1
               For Each objField In ObjectRst.Fields
                   'Select Case ObjectRst.Type
                      ' 下述資料型態則予以略過。
                   '   Case adGUID, adLongVarBinary, adLongVarWChar
                   '   Case Else
                        objWorkSheet.Cells(RowCount, ColCount).Value = Trim(ObjectRst.Fields(objField.Name).Value & vbNullString)
                        ColCount = ColCount + 1
                   'End Select
               Next objField
               ObjectRst.MoveNext
               RowCount = RowCount + 1
          End Select
       Loop
       
       '------------------------------------------------
       ' A3 Excel 自動調整欄寬。
       '------------------------------------------------
        ColCount = 1
        For Each objField In ObjectRst.Fields
            'Select Case ObjectRst.Type
               ' 下述資料型態則予以略過。
            '   Case adGUID, adLongVarBinary, adLongVarWChar
            '   Case Else
                 objWorkSheet.Columns(ColCount).AutoFit
                 ColCount = ColCount + 1
            'End Select
        Next objField
       
        '------------------------------------------------
        ' B2 另存檔案。
        '------------------------------------------------
        objWorkSheet.SaveAs ExcelFile
    
        '------------------------------------------------
        ' Z0 結束作業。
        '------------------------------------------------
        ' 關閉 Workbook。
        objWorkbook.Close
               
        ' 結束 Excel 作業。
        objExcel.Quit
              
        ' 釋放物件所佔空間。
        Set objField = Nothing
        Set objWorkSheet = Nothing
        Set objWorkbook = Nothing
        Set objExcel = Nothing

        Screen.MousePointer = vbDefault
        
        Exit Function
        
ErrorProcessSection:
        
        Select Case Err.Number
           Case 0
           Case Else
             ' 出現錯誤訊息。
             MsgBox "匯出失敗，原因如下：" & vbCrLf & vbCrLf & Err.Number & ": " & Err.Description, _
                    vbOKOnly + vbCritical, "匯出失敗"
             ' 關閉 Workbook。
             objWorkbook.Close
                        
             ' 結束 Excel 作業。
             objExcel.Quit
                   
             ' 載出物件變數。
             Set objField = Nothing
             Set objWorkSheet = Nothing
             Set objWorkbook = Nothing
             Set objExcel = Nothing
        End Select
End Function


Public Sub CopyToExcelByModel(Rs As ADODB.Recordset)
 Dim xlApp As Excel.Application
 Dim xlsBook As Excel.Workbook

 Dim iCol As Long
 Dim iRow As Long
 Dim Qty As Long
 Dim Model As String
    Set xlApp = CreateObject("Excel.Application")
    Set xlsBook = xlApp.Workbooks.Add
    xlApp.DisplayAlerts = False
    xlApp.UserControl = True
    xlApp.Visible = True
   
      Do While Not Rs.EOF
          Model = GetModel(Trim(Rs!part_num))
          DoEvents
          If Trim(xlApp.Worksheets(1).Name) = Trim(Model) Then
             iRow = iRow + 1
             Qty = Qty + 1
             For iCol = 1 To Rs.Fields.Count
                 xlApp.Worksheets(1).Cells(iRow, iCol).Value = Rs(iCol - 1)
             Next
          ElseIf Trim(xlApp.Worksheets(1).Name) = "Sheet1" Then
             'xlApp.Worksheets.Add
             xlApp.Worksheets(1).Name = Trim(Model)
             For iCol = 1 To Rs.Fields.Count
                xlApp.Worksheets(1).Cells(1, iCol).Value = Rs.Fields(iCol - 1).Name
                xlApp.Worksheets(1).Cells(2, iCol).Value = Rs(iCol - 1)
             Next
             iRow = 2
             Qty = 1
          Else
             xlApp.Worksheets(1).Cells(iRow + 1, 2).Value = "QUANTITY"
             xlApp.Worksheets(1).Cells(iRow + 1, 3).Value = Qty
             xlApp.Worksheets.Add
             xlApp.Worksheets(1).Name = Trim(Model)
             For iCol = 1 To Rs.Fields.Count
                xlApp.Worksheets(1).Cells(1, iCol).Value = Rs.Fields(iCol - 1).Name
                xlApp.Worksheets(1).Cells(2, iCol).Value = Rs(iCol - 1)
             Next
             iRow = 2
             Qty = 1
          End If
          Rs.MoveNext
          If Rs.EOF Then
             xlApp.Worksheets(1).Cells(iRow + 1, 2).Value = "QUANTITY"
             xlApp.Worksheets(1).Cells(iRow + 1, 3).Value = Qty
          End If
      Loop
      xlApp.Visible = True
    ' Close ADO objects
   
    Rs.Close
    Set Rs = Nothing
    
    ' Release Excel references
'    Set xlWs = Nothing
    Set xlApp = Nothing
End Sub
Public Function GetModel(PN As String) As String
  If IsNumeric(Mid(Trim(PN), 5, 1)) = True Then
     GetModel = Mid(Trim(PN), 2, 3)
  Else
     GetModel = Mid(Trim(PN), 2, 4)
  End If
End Function
