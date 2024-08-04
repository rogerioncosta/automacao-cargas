Attribute VB_Name = "Módulo2"
Sub carga() 'Ctrl Shift C
Attribute carga.VB_ProcData.VB_Invoke_Func = "C\n14"

    'Set wb = Workbooks.Open("C:\Users\Usuario\Desktop\Gerenciamento de Viagem (1).xls")
    
    Dim filePath As String
    
    ' Construir o caminho dinâmico para o arquivo no desktop do usuário
    filePath = Environ("USERPROFILE") & "\Desktop\Gerenciamento de Viagem (1).xls"
    
    ' Abre a planilha com o caminho dinâmico
    On Error Resume Next
    ' Abre a planilha 2
    Set wb = Workbooks.Open(filePath)
    If Err.Number <> 0 Then
        MsgBox "O arquivo não pôde ser encontrado. Verifique o caminho ou o nome do arquivo.", vbExclamation
        Err.Clear
        Exit Sub
    End If
    On Error GoTo 0
    
    Rows("1:2").Select
    Range("BP1").Activate
    Selection.Delete Shift:=xlUp
    
    Rows(1).Find(What:="Cargas", LookAt:=xlWhole).Select
    
    'ActiveCell.Offset(1, 0).Select
    Dim activeColumn
    activeColumn = ActiveCell.Column
    
    ActiveCell.Offset(1, 1).Select
    
    Dim rws
    rws = Cells(Rows.Count, activeColumn).End(xlUp).Row
    'rws = Columns.Item(1).Rows.Count
    
    Dim celulaCarga
    
    For rw = 1 To rws
    celulaCarga = ActiveCell.Offset(0, -1)
        If celulaCarga <> "" Then
            cargaFormat = Mid(celulaCarga, 10, 10)
            ActiveCell.Value = cargaFormat
            ActiveCell.Offset(1, 0).Select
        Else: GoTo continua
        End If
    Next rw
    
continua:
    
    'Application.Wait Now + TimeValue("00:00:02")
    
    ActiveCell.EntireColumn.Copy
    
    Dim wbDestination As Workbook
    Set wbDestination = Workbooks("Gerenciamento de Viagem.xls")
    
    Set wsDestination = wbDestination.Sheets("Gerenciamento de Viagem")
    wsDestination.Range("D:D").PasteSpecial Paste:=xlPasteValues
    
    Workbooks("Gerenciamento de Viagem (1).xls").Close False
    
    'Range("D:D").Paste
    'wsDestination.Range("P:P").Select
    Selection.Copy
    Range("P:P").Select
    ActiveSheet.Paste
    Range("L:L").Select
    Selection.Cut
    Range("M:M").Select
    ActiveSheet.Paste
    Rows(1).Find(What:="Embarcador", LookAt:=xlWhole).Select
    Selection.EntireColumn.Select
    Selection.Copy
    Range("N:N").Select
    ActiveSheet.Paste
    
    ' Formata o valor do fornecedor para encurtar o nome
    Rows(1).Find(What:="Embarcador", LookAt:=xlWhole).Select
    
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
        Else: GoTo continua2
        End If
    Next rw
    
continua2:

    Selection.EntireColumn.Select
    Selection.Cut
    Range("N:N").Select
    ActiveSheet.Paste
    
End Sub

