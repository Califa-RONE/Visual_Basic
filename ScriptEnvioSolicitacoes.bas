Attribute VB_Name = "Enviar_Solicitacao"
Sub Enviar_Solicitacao()

    Dim resposta As Integer
    Dim wsOrigem As Worksheet
    Dim wsDestino As Worksheet
    Dim wsFornecedor As Worksheet
    Dim OutlookApp As Object
    Dim OutlookMail As Object
    Dim EnviarPara As String
    Dim Nome As String
    Dim Cliente As String
    Dim wsEmail As Worksheet
    
    ThisWorkbook.Application.DisplayAlerts = False
    ThisWorkbook.Application.DisplayFullScreen = True
    
    On Error GoTo Erro
     
    resposta = MsgBox("Deseja mesmo continuar?", vbYesNo + vbQuestion, "Enviar solicita��es")
    
    If resposta = vbNo Then Exit Sub
    
    ThisWorkbook.Application.ScreenUpdating = False
     
    Workbooks.Open Filename:="B:\PLANILHAS DE COTA��ES\BACKUP_COTACOES.xlsm"
    
    Set wsDestinoB = Workbooks("BACKUP_COTACOES.xlsm").Worksheets("CADASTROS")
    Set wsDestino = Workbooks("BACKUP_COTACOES.xlsm").Worksheets("COTACOES")
    Set wsOrigem = ThisWorkbook.Worksheets("COTACOES_CONTROLE")
    Set wsFornecedor = ThisWorkbook.Worksheets("DataBase")
    Set wsEmail = ThisWorkbook.Worksheets("emails")
          
    wsOrigem.Unprotect "@Mcgrath951902!"
    X = wsOrigem.Cells(Rows.Count, 3).End(xlUp).Row
    
    codiA = "�����������������������������������������������.-"
    codiB = "aaaaaeeeeiiiiooooouuuuAAAAAEEEEIIIOOOOOUUUUcCnN  "
    
    If X <= 2 Then
        Workbooks("BACKUP_COTACOES.xlsm").Close savechanges:=True
        resposta = MsgBox("N�o existem itens para serem enviados!", vbOKOnly + vbExclamation, "Alerta")
        wsOrigem.Protect "@Mcgrath951902!"
        ThisWorkbook.Application.ScreenUpdating = True
        Exit Sub
    End If
    
    i = 3
    check = 0
    check1 = 0
    
    check2 = 0
    check3 = 0
    check4 = 0
    check5 = 0
    check6 = 0
    check7 = 0
    
    checkall = 0
    
    While X >= i
       marca = ""
       subgrupo = ""
       forn = ""
       NCM = ""
       If wsOrigem.Cells(X, 2).Value = "COTAR" Or wsOrigem.Cells(X, 2).Value = "CADASTRAR" Or wsOrigem.Cells(X, 2).Value = "INCLUIR" Or wsOrigem.Cells(X, 2).Value = "ALTERAR" Then
          If wsOrigem.Cells(X, 7).Value <> "" Or _
             wsOrigem.Cells(X, 11).Value <> "" Or _
             wsOrigem.Cells(X, 13).Value <> "" Or _
             wsOrigem.Cells(X, 14).Value <> "" Or _
             wsOrigem.Cells(X, 15).Value <> "" Or _
             wsOrigem.Cells(X, 17).Value <> "" Or _
             wsOrigem.Cells(X, 18).Value <> "" Or _
             wsOrigem.Cells(X, 16).Value <> "" Or _
             wsOrigem.Cells(X, 19).Value <> "" Or _
             wsOrigem.Cells(X, 20).Value <> "" Or _
             wsOrigem.Cells(X, 21).Value <> "" Or _
             wsOrigem.Cells(X, 25).Value <> "" Or _
             wsOrigem.Cells(X, 30).Value <> "" Or _
             wsOrigem.Cells(X, 36).Value <> "" And _
             wsOrigem.Cells(X, 15).Value <> "A DEFINIR" _
             Then
             
            If wsOrigem.Cells(X, 7).Value = "M�DIA" Or wsOrigem.Cells(X, 7).Value = "ALTA" Then
                If wsOrigem.Cells(X, 8).Value = "N/A" Or wsOrigem.Cells(X, 8).Value = "" Then
                    Cod = wsOrigem.Cells(X, 16).Value
                    Motivo = InputBox("Digite o motivo pelo qual definiu o item de c�digo " & Cod & " como urgente: ", "Insira um Motivo")
                    wsOrigem.Cells(X, 8).Value = Motivo
                    If Motivo = False Or Motivo = "" Then
                    If check2 = 0 Then
                        resposta = MsgBox("Para definir itens com urgencia alta ou m�dia voc� deve inserir um motivo v�lido!", vbOKOnly + vbCritical, "ERRO")
                    End If
                    check2 = 1
                    wsOrigem.Cells(X, 2).Value = "REVISAR"
                    GoTo continuar
                    End If
                End If
            End If
            If wsOrigem.Cells(X, 11).Value = "A DEFINIR" Then
                Cod = wsOrigem.Cells(X, 16).Value
                marca = InputBox("Digite o nome da Marca/Fabricante do item de c�digo: " & Cod, "Nova Marca")
                wsOrigem.Cells(X, 11).Value = marca
                wsOrigem.Cells(X, 10).Value = "0"
                If marca = False Or marca = "" Then
                   wsOrigem.Cells(X, 2).Value = "REVISAR"
                   GoTo continuar
                End If
            End If
            If wsOrigem.Cells(X, 13).Value = "A DEFINIR" Then
                Cod = wsOrigem.Cells(X, 16).Value
                subgrupo = InputBox("Digite o nome do Modelo da M�quina/Sub Grupo do item de c�digo: " & Cod, "Novo SubGrupo")
                wsOrigem.Cells(X, 13).Value = subgrupo
                wsOrigem.Cells(X, 12).Value = "0"
                If subgrupo = False Or subgrupo = "" Then
                    wsOrigem.Cells(X, 2).Value = "REVISAR"
                    GoTo continuar
                End If
            End If
            If wsOrigem.Cells(X, 24).Value = "A DEFINIR" Then
                Cod = wsOrigem.Cells(X, 16).Value
                NCM = InputBox("Digite o NCM do item de c�digo: " & Cod, "Novo NCM")
                wsOrigem.Cells(X, 24).Value = NCM
                wsOrigem.Cells(X, 23).Value = "0"
                If NCM = False Or NCM = "" Then
                    wsOrigem.Cells(X, 2).Value = "REVISAR"
                    GoTo continuar
                End If
            End If
            If wsOrigem.Cells(X, 15).Value = "A DEFINIR" Or wsOrigem.Cells(X, 14).Value = "0" Then
                If check7 = 0 Then
                    resposta = MsgBox("Voc� deve inserir um forncedor v�lido! Entre em contato com a Assist�ncia T�cnica caso necess�rio.", vbOKOnly + vbCritical, "ERRO")
                End If
                    check7 = 1
                    wsOrigem.Cells(X, 2).Value = "REVISAR"
                    GoTo continuar
            End If
            
            If wsOrigem.Cells(X, 32).Value <> "N/A" And wsOrigem.Cells(X, 32).Value <> "" Then
                
                Y = wsDestinoB.Cells(Rows.Count, 3).End(xlUp).Row
            
                If Y = 2 And wsDestinoB.Cells(Y, 6).Value = "" Then
                    GoTo continuar4
                End If
                wsDestinoB.Range("Tabela2").ListObject.ListRows.Add AlwaysInsert:=True
                Y = wsDestinoB.Cells(Rows.Count, 3).End(xlUp).Row
continuar4:
                wsOrigem.Activate
                wsOrigem.Cells(X, 16).Value = UCase(wsOrigem.Cells(X, 16).Value)
        
                temp = wsOrigem.Cells(X, 16).Value
    
                For h = 1 To Len(temp)
                    p = InStr(codiA, Mid(temp, h, 1))
                    If p > 0 Then Mid(temp, h, 1) = Mid(codiB, p, 1)
                Next h
       
                wsOrigem.Cells(X, 16).Value = temp
                
                For h = 1 To Len(wsOrigem.Cells(X, 16).Value)
                    If wsOrigem.Range(Cells(X, 16), Cells(X, 16)).Characters(h, 1).Text = " " Then
                        wsOrigem.Range(Cells(X, 16), Cells(X, 16)).Characters(h, 1).Delete
                    End If
                Next h
                
                wsOrigem.Activate
                wsOrigem.Range(Cells(X, 2), Cells(X, 37)).Copy
            
                wsDestinoB.Activate
                wsDestinoB.Range(Cells(Y, 2), Cells(Y, 37)).PasteSpecial xlPasteAll
                checkall = 1
                
                wsOrigem.Activate
                wsOrigem.Range(Cells(X, 15), Cells(X, 15)).Copy
            
                wsDestinoB.Activate
                wsDestinoB.Range(Cells(Y, 15), Cells(Y, 15)).PasteSpecial xlPasteValues
            
                If marca <> "" Then
                    wsDestinoB.Cells(Y, 11).Value = marca
                End If
                If subgrupo <> "" Then
                    wsDestinoB.Cells(Y, 13).Value = subgrupo
                End If
                If NCM <> "" Then
                    wsDestinoB.Cells(Y, 24).Value = NCM
                End If
                wsDestinoB.Cells(Y, 4).Value = Date
                wsDestinoB.Cells(Y, 5).Value = Time
                wsDestinoB.Cells(Y, 2).Value = "CADASTRAR"
                
continuo:
            
            Z = wsEmail.Cells(Rows.Count, 1).End(xlUp).Row
            
            produto = wsDestinoB.Cells(Y, 29).Value
            Cod = wsDestinoB.Cells(Y, 16).Value
            Nome = wsDestinoB.Cells(Y, 6).Value
            Cliente = wsDestinoB.Cells(Y, 36).Value
            Data = wsDestinoB.Cells(Y, 4).Value
            Fornecedor = wsDestinoB.Cells(Y, 15).Value
            
            For b = 1 To Z
            If Nome = wsEmail.Cells(b, 1).Value Then
                EnviarPara = wsEmail.Cells(b, 2).Value
                Set OutlookApp = CreateObject("Outlook.Application")
                Set OutlookMail = OutlookApp.CreateItem(0)
                    With OutlookMail
                        .To = EnviarPara & ";" & "pdm@tabatex.com.br"
                        .CC = ""
                        .BCC = ""
                        .Subject = "ITEM " & Cod & " ENVIADO PARA CADASTRO."
                        If Cliente <> "" Then
                        .htmlBody = "Mensagem de confirma��o de envio de solicita��o. <br><br>" & _
                        "O item " & produto & " de c�digo " & Cod & " solicitado pelo cliente " & Cliente & " no dia " & Data & " cotado no fornecedor " & Fornecedor & " e enviado para cadastro.<br><br>" & _
                        "Esta � uma mensagem autom�tica enviada do sistema de envio de solicita��o."
                        Else
                        .htmlBody = "Mensagem de confirma��o de envio de solicita��o. <br><br>" & _
                        "O item " & produto & " de c�digo " & Cod & " solicitado no dia " & Data & " cotado no fornecedor " & Fornecedor & " e enviado para cadastro.<br><br>" & _
                        "Esta � uma mensagem autom�tica enviada do sistema de envio de solicita��o."
                        End If
                        .Send
                    End With
                Set OutlookMail = Nothing
                Set OutlookApp = Nothing
            End If
            Next b
            
                wsOrigem.Cells(X, 2).Value = "AGUARDANDO CADASTRO"
                 For j = 2 To 37
                    wsOrigem.Cells(X, j).Locked = True
                    wsOrigem.Cells(X, j).FormulaHidden = True
                Next j
                GoTo continuarB
                
            End If
            
            Y = wsDestino.Cells(Rows.Count, 3).End(xlUp).Row
            
            If Y = 2 And wsDestino.Cells(Y, 6).Value = "" Then
                GoTo continuar5
            End If
            wsDestino.Range("Tabela1").ListObject.ListRows.Add AlwaysInsert:=True
            Y = wsDestino.Cells(Rows.Count, 3).End(xlUp).Row
continuar5:
            wsOrigem.Activate
            wsOrigem.Range(Cells(X, 2), Cells(X, 37)).Copy
            
            wsDestino.Activate
            wsDestino.Range(Cells(Y, 2), Cells(Y, 37)).PasteSpecial xlPasteAll
            checkall = 1
            
            wsOrigem.Activate
            wsOrigem.Range(Cells(X, 15), Cells(X, 15)).Copy
            
            wsDestino.Activate
            wsDestino.Range(Cells(Y, 15), Cells(Y, 15)).PasteSpecial xlPasteValues
            
            If marca <> "" Then
                wsDestino.Cells(Y, 11).Value = marca
            End If
            If subgrupo <> "" Then
                wsDestino.Cells(Y, 13).Value = subgrupo
            End If
            If NCM <> "" Then
                wsDestino.Cells(Y, 24).Value = NCM
            End If
            wsDestino.Cells(Y, 4).Value = Date
            wsDestino.Cells(Y, 5).Value = Time
            wsDestino.Cells(Y, 2).Value = "COTAR"
            
continuoB:
            produto = wsDestino.Cells(Y, 29).Value
            Cod = wsDestino.Cells(Y, 16).Value
            Nome = wsDestino.Cells(Y, 6).Value
            Cliente = wsDestino.Cells(Y, 36).Value
            Data = wsDestino.Cells(Y, 4).Value
            Fornecedor = wsDestino.Cells(Y, 15).Value
            
            For b = 1 To Z
            If Nome = wsEmail.Cells(b, 1).Value Then
                EnviarPara = wsEmail.Cells(b, 2).Value
                Set OutlookApp = CreateObject("Outlook.Application")
                Set OutlookMail = OutlookApp.CreateItem(0)
                    With OutlookMail
                        .To = EnviarPara
                        .CC = ""
                        .BCC = ""
                        .Subject = "ITEM " & Cod & " ENVIADO PARA COTA��O."
                        If Cliente <> "" Then
                        .htmlBody = "Mensagem de confirma��o de envio de solicita��o. <br><br>" & _
                        "O item " & produto & " de c�digo " & Cod & " solicitado pelo cliente " & Cliente & " no dia " & Data & " foi enviado para cota��o no fornecedor " & Fornecedor & ".<br><br>" & _
                        "Esta � uma mensagem autom�tica enviada do sistema de envio de solicita��o."
                        Else
                        .htmlBody = "Mensagem de confirma��o de envio de solicita��o. <br><br>" & _
                        "O item " & produto & " de c�digo " & Cod & " solicitado no dia " & Data & " foi enviado para cota��o no fornecedor " & Fornecedor & ".<br><br>" & _
                        "Esta � uma mensagem autom�tica enviada do sistema de envio de solicita��o."
                        End If
                        .Send
                    End With
                Set OutlookMail = Nothing
                Set OutlookApp = Nothing
            End If
            Next b
            
            wsOrigem.Cells(X, 2).Value = "AGUARDANDO COTA��O"
            
            For j = 2 To 37
                wsOrigem.Cells(X, j).Locked = True
                wsOrigem.Cells(X, j).FormulaHidden = True
            Next j
            
continuarB:
           
          Else
            If check = 0 Then
                resposta = MsgBox("Existem dados importantes que n�o foram preenchidos!", vbOKOnly + vbCritical, "ERRO")
                check = 1
            Else
                GoTo continuar
            End If
          End If
       Else

continuar:
       
        If X = i And checkall = 0 Then
            Workbooks("BACKUP_COTACOES.xlsm").Close savechanges:=True
            resposta = MsgBox("N�o existem itens para serem enviados!", vbOKOnly + vbExclamation, "Alerta")
            wsOrigem.Activate
            wsOrigem.Protect "@Mcgrath951902!"
            ThisWorkbook.Application.ScreenUpdating = True
            Exit Sub
       End If
       X = X - 1
        
       End If
      
    Wend
    
    wsOrigem.Activate
    wsOrigem.Protect "@Mcgrath951902!"
    Workbooks("BACKUP_COTACOES.xlsm").Close savechanges:=True
    ThisWorkbook.Application.ScreenUpdating = True
    resposta = MsgBox("A a��o foi conclu�da!", vbOKOnly + vbInformation, "Solicita��es Enviadas")
Exit Sub
Erro:
    Workbooks("BACKUP_COTACOES.xlsm").Close savechanges:=True
    resposta = MsgBox("Por algum motivo a a��o n�o pode ser conclu�da. Tente novamente mais tarde.", vbOKOnly + vbExclamation, "Alerta")
    wsOrigem.Protect "@Mcgrath951902!"
    ThisWorkbook.Application.ScreenUpdating = True
End Sub
