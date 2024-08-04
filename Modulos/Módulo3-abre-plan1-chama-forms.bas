Attribute VB_Name = "Módulo3"
Sub gal() 'ctrl shift S
Attribute gal.VB_ProcData.VB_Invoke_Func = "G\n14"
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim filePath As String
    
    ' Construir o caminho dinâmico para o arquivo no desktop do usuário
    filePath = Environ("USERPROFILE") & "\Desktop\Gerenciamento de Viagem.xls"
    
    'Unload UserForm2
    
    If UserForm2.Visible = False Then
        UserForm2.Show
    
    Else
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
        
        'Set wb = Workbooks.Open("C:\Users\Usuario\Desktop\Gerenciamento de Viagem.xls")
        'Set ws = wb.Sheets("Gerenciamento de Viagem") ' Substitua "NomeDaPlanilha2" pelo nome real da sua planilha 2
        
        Application.Wait Now + TimeValue("00:00:01")
        
        UserForm1.Show
        
        UserForm3.Show
    End If
    'Set wb = Workbooks.Open("C:\Users\Win10\Desktop\Gerenciamento de Viagem (1).xls")
    
    'Application.SendKeys ("%{Tab}")
    ' Executa a macro da planilha 2
    'Application.SendKeys "^+R"
    
    ' Fecha a planilha 2 sem salvar as alterações
    'wb.Close SaveChanges:=False
End Sub
