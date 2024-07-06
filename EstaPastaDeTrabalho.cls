VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EstaPastaDeTrabalho"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_Open()
    UserForm2.Show
End Sub

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)

    Dim Senha As String
    Senha = "******"
    
    If UserForm1.Visible = True Then
    Exit Sub
    
    Else

    If InputBox("Digite a senha para Salvar, ou apenas feche.", "Proteção") = Senha Then
        Exit Sub
    Else
        If SaveAsUI = True Then
            MsgBox "Não é permitido Salvar Como"
            Cancel = True
            Exit Sub
        End If
        
        If SaveAsUI = False Then
            MsgBox "Não é permitido 'Salvar'"
            Cancel = True
            Exit Sub
        End If
    End If
End If

End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    'senha plan bloqueadas "142"
    Application.DisplayAlerts = False
    Application.Quit
End Sub


