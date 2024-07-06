Attribute VB_Name = "Módulo1"
Sub galileu() 'ctrl shift R
Attribute galileu.VB_ProcData.VB_Invoke_Func = " \n14"
'
' galileu Macro
'
'
'    Workbooks.Open Filename:="C:\Users\Win10\Desktop\Gerenciamento de Viagem.xls"
    Rows("1:2").Select
    Range("BP1").Activate
    Selection.Delete Shift:=xlUp
    Columns("A:G").Select
    Selection.Delete Shift:=xlToLeft
    Columns("B:B").Select
    Selection.Insert Shift:=xlToRight
    Selection.Insert Shift:=xlToRight
    Selection.Insert Shift:=xlToRight
    Columns("I:I").Select
    Selection.Cut
    Columns("E:E").Select
    Selection.Insert Shift:=xlToRight
    Columns("F:F").Select
    Selection.Cut
    Columns("I:I").Select
    Selection.Insert Shift:=xlToRight
    Columns("L:L").Select
    Selection.Cut
    Columns("H:H").Select
    Selection.Insert Shift:=xlToRight
    Columns("L:L").Select
    Selection.Copy
    Columns("C:C").Select
    ActiveSheet.Paste
End Sub
