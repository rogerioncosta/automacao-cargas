Attribute VB_Name = "Módulo3"
Sub gal() 'ctrl shift S
Attribute gal.VB_ProcData.VB_Invoke_Func = "G\n14"
    Dim wb As Workbook
    Dim ws As Worksheet
    
    'Unload UserForm2
    
    If UserForm2.Visible = False Then
        UserForm2.Show
    
    Else
        ' Abre a planilha 1
        Set wb = Workbooks.Open("C:\Users\Win10\Desktop\Gerenciamento de Viagem.xls")
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
