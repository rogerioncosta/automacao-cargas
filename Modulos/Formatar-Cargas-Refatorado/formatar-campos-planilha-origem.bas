Attribute VB_Name = "Módulo1"
Function procuraCabecalho(cabecalho As String, coluna As String, acao As String)
    Rows(1).Find(what:=cabecalho, LookAt:=xlWhole).Select
    ActiveCell.EntireColumn.Select
    
    If acao = "Copy" Then
        Selection.Copy
    End If
    
    If acao = "Cut" Then
        Selection.Cut
    End If
    
    Columns(coluna).Select
    ActiveSheet.Paste

End Function

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
    Columns("A:A").Select
    'Selection.Delete Shift:=xlToLeft
    'Columns("A:A").Select
    
    For i = 1 To 17
        Selection.Insert Shift:=xlToRight
    
    Next i
    
    Call procuraCabecalho("Previsão de Coleta", "A:A", "Cut")
'    Rows(1).Find(what:="Previsão de Coleta", LookAt:=xlWhole).Select
'    ActiveCell.EntireColumn.Select
'    Selection.Cut
'    Columns("A:A").Select
'    ActiveSheet.Paste

    Call procuraCabecalho("Tipo de Operação", "C:C", "Cut")
'    Rows(1).Find(what:="Tipo de Operação", LookAt:=xlWhole).Select
'    ActiveCell.EntireColumn.Select
'    Selection.Cut
'    Columns("C:C").Select
'    ActiveSheet.Paste

    Call procuraCabecalho("Mun. Destino", "E:E", "Cut")
'    Rows(1).Find(what:="Mun. Destino", LookAt:=xlWhole).Select
'    ActiveCell.EntireColumn.Select
'    Selection.Cut
'    Columns("E:E").Select
'    ActiveSheet.Paste

    Call procuraCabecalho("Placa do Cavalo", "F:F", "Cut")
'    Rows(1).Find(what:="Placa do Cavalo", LookAt:=xlWhole).Select
'    ActiveCell.EntireColumn.Select
'    Selection.Cut
'    Columns("F:F").Select
'    ActiveSheet.Paste

    Call procuraCabecalho("Placa da Carreta", "G:G", "Cut")
'    Rows(1).Find(what:="Placa da Carreta", LookAt:=xlWhole).Select
'    ActiveCell.EntireColumn.Select
'    Selection.Cut
'    Columns("G:G").Select
'    ActiveSheet.Paste
'
    Call procuraCabecalho("CPF do Motorista", "H:H", "Cut")
'    Rows(1).Find(what:="CPF do Motorista", LookAt:=xlWhole).Select
'    ActiveCell.EntireColumn.Select
'    Selection.Cut
'    Columns("H:H").Select
'    ActiveSheet.Paste
'
    Call procuraCabecalho("Motorista Principal", "I:I", "Cut")
'    Rows(1).Find(what:="Motorista Principal", LookAt:=xlWhole).Select
'    ActiveCell.EntireColumn.Select
'    Selection.Cut
'    Columns("I:I").Select
'    ActiveSheet.Paste

    Call procuraCabecalho("Embarcador", "N:N", "Copy")
'    Rows(1).Find(what:="Embarcador", LookAt:=xlWhole).Select
'    ActiveCell.EntireColumn.Select
'    Selection.Copy
'    Range("N:N").Select
'    ActiveSheet.Paste
    
    activeColumn = ActiveCell.Column
    
    ActiveCell.Offset(1, 1).Select
    
    
    Dim embarcador
    embarcador = Cells(Rows.Count, activeColumn).End(xlUp).Row
    
    Dim celulaEmbarcador
    
    For rw = 1 To embarcador
    celulaEmbarcador = ActiveCell.Offset(0, -1)
        If celulaEmbarcador <> "" Then
            cargaFormat = Mid(celulaEmbarcador, 1, 15)
            ActiveCell.Value = cargaFormat
            ActiveCell.Offset(1, 0).Select
        Else: ActiveCell.Offset(1, 0).Select
        End If
    Next rw
    
    Selection.EntireColumn.Select
    Selection.Cut
    Range("N:N").Select
    ActiveSheet.Paste
    
    
    Call procuraCabecalho("ID-THub Destino", "L:L", "Copy")
'    Rows(1).Find(what:="ID-THub Destino", LookAt:=xlWhole).Select
'    ActiveCell.EntireColumn.Select
'    Selection.Copy
'    Range("L:L").Select
'    ActiveSheet.Paste

    Range("L1").Select
    
    Call procuraCabecalho("Tipo de Operação", "M:M", "Copy")
'    Rows(1).Find(what:="Tipo de Operação", LookAt:=xlWhole).Select
'    ActiveCell.EntireColumn.Select
'    Selection.Copy
'    Columns("M:M").Select
'    ActiveSheet.Paste
    
    
End Sub
