Sub main()
'###Open File###
Dim mb As Integer
Dim str_FileName As String
str_FileName = Range("b1")
mb = MsgBox("This is message." & vbCrLf & "open file :" & str_FileName, vbYesNo, "Title:open files")
Workbooks.Open FileName:=str_FileName

MsgBox ActiveWorkbook.Name


'###データ開始位置###
Dim x, y As Integer
x = 2
y = 4


Dim int_maxrow As Integer
Dim int_maxcol As Integer
int_maxrow = Cells(Rows.Count, x).End(xlUp).Row
int_maxcol = Cells(Columns.Count, y).End(xlToLeft).Column
mb = MsgBox("int_maxrow =" & int_maxrow, vbYesNo, "Title:open files")
mb = MsgBox("int_maxcol =" & int_maxcol, vbYesNo, "Title:open files")


End Sub

Function Fn_read(a As String)


End Function
