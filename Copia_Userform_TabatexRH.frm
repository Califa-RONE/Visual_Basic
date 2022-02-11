VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm7 
   Caption         =   "TABATEXRH"
   ClientHeight    =   5688
   ClientLeft      =   84
   ClientTop       =   372
   ClientWidth     =   8844.001
   OleObjectBlob   =   "Copia_Userform_TabatexRH.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton2_Click()

    resposta = MsgBox("Deseja realmente sair do sistema?", vbYesNo + vbQuestion, "Pergunta")
        If resposta = vbNo Then
            Exit Sub
        End If
        
    UserForm7.TextBox2.Value = ""
    UserForm7.Hide
    Load UserForm1
    UserForm1.TextBox1.Value = ""
    UserForm1.TextBox2.Value = ""
    UserForm1.StartUpPosition = 2
    UserForm1.Show
    Cancel = True

End Sub

Private Sub CommandButton3_Click()
    
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
    Dim fso
    Dim fDlg As FileDialog
    Dim lArquivo As String, Destino As String
    
    On Error GoTo Erro
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set fDlg = Application.FileDialog(FileDialogType:=msoFileDialogOpen)
    
    Matricula = TextBox2.Value
    
    If Matricula = "" Or Not IsNumeric(Matricula) Then
        resposta = MsgBox("Você deve digitar um número de matrícula válido!", vbOKOnly + vbExclamation, "Alerta")
        Application.DisplayAlerts = True
        Application.ScreenUpdating = True
        Exit Sub
    End If
    
    Destino = "\\192.168.0.9\RH-Sistema\Arquivos_Gerais\" & Matricula & "\"
    
    If Not fso.FolderExists(Destino) Then
        resposta = MsgBox("Não existem funcionários cadastrados com esta matrícula, contate a Assistencia Tecnica!", vbOKOnly + vbExclamation, "Alerta")
        Application.DisplayAlerts = True
        Application.ScreenUpdating = True
        Exit Sub
    End If
    With fDlg
        .AllowMultiSelect = True
        .InitialView = msoFileDialogViewDetails
        .Filters.Add "PDF", "*.pdf", 1
        .InitialFileName = "C:\"
    End With
    
    i = 1
    x = 1
    
    If fDlg.Show = -1 Then
        While i <= x
            lArquivo = fDlg.SelectedItems(i)
            If lArquivo <> "" Then
                fso.CopyFile lArquivo, Destino
            Else
                Application.DisplayAlerts = True
                Application.ScreenUpdating = True
                Exit Sub
            End If
            i = i + 1
            x = x + 1
        Wend
    Else
        MsgBox "Não foi selecionado nenhum arquivo"
    End If

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
Exit Sub
Erro:
    resposta = MsgBox("Ação concluída!", vbOKOnly + vbInformation, "Informação")
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Exit Sub
End Sub

Private Sub UserForm_Initialize()

    CommandButton3.Default = True
    CommandButton2.Cancel = True

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    
    If CloseMode = vbFormControlMenu Then
        resposta = MsgBox("Deseja realmente sair do sistema?", vbYesNo + vbQuestion, "Pergunta")
        If resposta = vbNo Then
            Exit Sub
        End If
        
        UserForm7.Hide
        Load UserForm1
        UserForm1.TextBox2.Value = ""
        UserForm1.StartUpPosition = 2
        UserForm1.Show
        Cancel = True
    End If
  
End Sub
