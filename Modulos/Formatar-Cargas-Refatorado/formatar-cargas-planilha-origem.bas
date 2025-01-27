Attribute VB_Name = "M�dulo2"
Sub carga() 'Ctrl Shift C
Attribute carga.VB_ProcData.VB_Invoke_Func = "C\n14"

    'Set wb = Workbooks.Open("C:\Users\Usuario\Desktop\Gerenciamento de Viagem (1).xls")
    
    Dim filePath As String
    
    ' Construir o caminho din�mico para o arquivo no desktop do usu�rio
    filePath = Environ("USERPROFILE") & "\Desktop\Gerenciamento de Viagem (1).xls"
    
    ' Abre a planilha com o caminho din�mico
    On Error Resume Next
    ' Abre a planilha 2
    Set wb = Workbooks.Open(filePath)
    If Err.Number <> 0 Then
        MsgBox "O arquivo n�o p�de ser encontrado. Verifique o caminho ou o nome do arquivo.", vbExclamation
        Err.Clear
        Exit Sub
    End If
    On Error GoTo 0
    
    Rows("1:2").Select
    Range("BP1").Activate
    Selection.Delete Shift:=xlUp
    
    Rows(1).Find(what:="Cargas", LookAt:=xlWhole).Select
    
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
        Else: ActiveCell.Offset(1, 0).Select
        End If
    Next rw

    
    'Application.Wait Now + TimeValue("00:00:02")
    
    ActiveCell.EntireColumn.Copy
    
    Dim wbDestination As Workbook
    Set wbDestination = Workbooks("Gerenciamento de Viagem.xls")
    
    Set wsDestination = wbDestination.Sheets("Gerenciamento de Viagem")
    wsDestination.Range("D:D").PasteSpecial Paste:=xlPasteValues
    
    wsDestination.Range("D:D").EntireColumn.Copy
    
    wsDestination.Range("P:P").PasteSpecial Paste:=xlPasteValues
    
    wsDestination.Range("L:L").EntireColumn.Copy
    wsDestination.Range("O:O").PasteSpecial Paste:=xlPasteValues
    wsDestination.Range("L:L").EntireColumn.ClearContents
        
    MsgBox "Os dados foram formatados!", vbInformation
    
    Workbooks("Gerenciamento de Viagem (1).xls").Close False
    
    
End Sub

