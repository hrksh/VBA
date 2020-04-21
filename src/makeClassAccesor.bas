Attribute VB_Name = "Module1"
Option Explicit
Const wsnameMem As String = "メンバ"
Const wsnameCls As String = "クラス"
Const errMsg_nosheetname As String = "シートが存在しません。"
Const errMsg_noVal As String = "変数が記入されていません。"

Dim errFlag As Boolean

Sub showErrMsg(msg As String)
  MsgBox msg, vbCritical, "error"
End Sub

Function getRowEnd(wsname As String, col As Integer) As Integer
  errFlag = changeSheet(wsname)
  If errFlag Then
    getRowEnd = -1
    Exit Function
  Else
    getRowEnd = ActiveSheet.Cells(Rows.Count, col).End(xlUp).row
  End If
End Function

Function getColEnd(wsname As String, row As Integer) As Integer
  errFlag = changeSheet(wsname)
  If errFlag Then
    getRowEnd = -1
    Exit Function
  Else
    getColEnd = ActiveSheet.Cells(row, Columns.Count).End(xlToLeft).Column
  End If
End Function

Function changeSheet(wsname As String) As Boolean
  changeSheet = True
  If isExistSheet(wsname) Then
    ThisWorkbook.Worksheets(wsname).Activate
    changeSheet = False
  End If
End Function

Function isExistSheet(wsname As String) As Boolean
  Dim ws As Worksheet
  isExistSheet = False
  For Each ws In ThisWorkbook.Worksheets()
    If ws.Name = wsname Then
      isExistSheet = True
      Exit Function
    End If
  Next
  showErrMsg (errMsg_nosheetname)
End Function

Sub writeToCell(wsname As String, row As Integer, col As Integer, value As String)
  ThisWorkbook.Worksheets(wsname).Cells(row, col).value = value
End Sub

Public Sub makeClsAccessor(celladd As String)
  
  Dim item As Object
  Dim row As Integer
  Dim valiable As range
  
  Application.ScreenUpdating = False
  
  changeSheet (wsnameMem)
  If range(celladd).value = "" Then
    Application.ScreenUpdating = True
    showErrMsg (errMsg_noVal)
    Exit Sub
  End If
  Set valiable = getRange(celladd)
  
  row = 1
  For Each item In valiable
    writeToCell wsnameCls, row, 1, naming(item.value)
    row = row + 1
  Next
  
  row = getRowEnd(wsnameCls, 1) + 2
  For Each item In valiable
    row = writeGetProperty(item.value, row)
  Next
  
  row = getRowEnd(wsnameCls, 1) + 2
  For Each item In valiable
    row = writeSetProperty(item.value, row)
  Next
  
  Application.ScreenUpdating = True
  
End Sub

Function getRange(celladd As String) As range

  Dim rowMax As Integer
  Dim row As Integer
  Dim col As Integer
  
  ActiveSheet.range(celladd).Select
  row = Selection.row
  col = Selection.Column
  rowMax = Selection.End(xlDown).row
  Set getRange = ActiveSheet.range(Cells(row, col), Cells(rowMax, col))
  
End Function

Function naming(target As String) As String
  naming = "Private " & target & "_ As String"
End Function

Function writeGetProperty(target As String, row As Integer) As Integer
  writeToCell wsnameCls, row, 1, "Property Get " & target & "() As String"
  writeToCell wsnameCls, row + 1, 1, "  " & target & " = " & target & "_"
  writeToCell wsnameCls, row + 2, 1, "End Property"
  writeGetProperty = row + 3
End Function

Function writeSetProperty(target As String, row As Integer) As Integer
  writeToCell wsnameCls, row, 1, "Property Let " & target & "(newinput As String)"
  writeToCell wsnameCls, row + 1, 1, "  " & target & "_ = newinput"
  writeToCell wsnameCls, row + 2, 1, "End Property"
  writeSetProperty = row + 3
End Function
