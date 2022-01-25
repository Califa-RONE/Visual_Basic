VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EstaPastaDeTrabalho"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_BeforeClose(Cancel As Boolean)

        ThisWorkbook.Close savechanges:=True

End Sub

Private Sub Workbook_Open()

        Dim pesquisa As Range
        Dim wsOrigem As Worksheet
        Dim wsDestino As Worksheet
        Dim wsOrigem2 As Worksheet
        Dim fs, f, s, f2, s2
        
        Application.DisplayAlerts = False
        Application.ScreenUpdating = False
        Application.DisplayStatusBar = False
        Application.EnableEvents = False
        ActiveSheet.DisplayPageBreaks = False
        Application.Calculation = xlCalculationManual
        Application.UserName = usuario
        
        On Error GoTo Erro
        
        If Dir("\\192.168.0.9\p&d\PDM\Solicitação de Cadastro\Common\Consulta_Produtos\DadosProduto.xls") = "" Then
            Application.DisplayAlerts = True
            Application.ScreenUpdating = True
            Application.DisplayStatusBar = True
            Application.EnableEvents = True
            ActiveSheet.DisplayPageBreaks = True
            Application.Calculation = xlCalculationAutomatic
            resposta = MsgBox("Está faltando a tabela DadosProduto!", vbOKOnly + vbExclamation, "Alerta")
            Exit Sub
        End If
        If Dir("\\192.168.0.9\p&d\PDM\Solicitação de Cadastro\Common\Consulta_Produtos\CodEquivalentes.xls") = "" Then
            Application.DisplayAlerts = True
            Application.ScreenUpdating = True
            Application.DisplayStatusBar = True
            Application.EnableEvents = True
            ActiveSheet.DisplayPageBreaks = True
            Application.Calculation = xlCalculationAutomatic
            resposta = MsgBox("Está faltando a tabela CodEquivalentes!", vbOKOnly + vbExclamation, "Alerta")
            Exit Sub
        End If
        
        Set fs = CreateObject("Scripting.FileSystemObject")
        Set f = fs.Getfile("\\192.168.0.9\p&d\PDM\Solicitação de Cadastro\Common\Consulta_Produtos\DadosProduto.xls")
        Set f2 = fs.Getfile("\\192.168.0.9\p&d\PDM\Solicitação de Cadastro\Common\Consulta_Produtos\CodEquivalentes.xls")
        Set f3 = fs.Getfile("\\192.168.0.9\p&d\PDM\Solicitação de Cadastro\Common\Consulta_Produtos\vicunha.xlsx")
        
        s = f.DateCreated
        s2 = f2.DateCreated
                
        Set wsDestino2 = Workbooks("vicunha.xlsx").Worksheets("Consulta_Produtos")
        Set wsDestino = Workbooks("Consulta_Produtos.xlsm").Worksheets("Consulta_Produtos")
                
        wsDestino.Activate
        
        If s <> wsDestino.Cells(1, 22).Value Or s2 <> wsDestino.Cells(1, 23).Value And usuario = "PDM" Or _
        usuario = "Livia Trevisoni" Or usuario = "Bruno Henrique" Then
            resposta = MsgBox("Existem dados para serem importados, deseja continuar? " + _
            vbCrLf + vbCrLf + "(O tempo de importação pode variar entre 30 - 60 minutos ou mais, dependendo da quantidade de dados a serem importados)", _
            vbYesNo + vbExclamation, "Alerta")
            If resposta = vbNo Then
                Application.DisplayAlerts = True
                Application.ScreenUpdating = True
                Application.DisplayStatusBar = True
                Application.EnableEvents = True
                ActiveSheet.DisplayPageBreaks = True
                Application.Calculation = xlCalculationAutomatic
                Exit Sub
            End If
            wsDestino.Cells(1, 22).Value = s
            wsDestino.Cells(1, 23).Value = s2
        Else
            Application.DisplayAlerts = True
            Application.ScreenUpdating = True
            Application.DisplayStatusBar = True
            Application.EnableEvents = True
            ActiveSheet.DisplayPageBreaks = True
            Application.Calculation = xlCalculationAutomatic
            resposta = MsgBox("Data e hora de atualização dos arquivos de dados dos produtos: " + vbCrLf + vbCrLf + "Lista de produtos " & s & ";" + vbCrLf + "Códigos equivalentes " & s2 & ".", vbOKOnly + vbInformation, "Informação")
            Call Vicunha
            wsDestino.Activate
            Cells(1, 1).Select
            Exit Sub
        End If
        
        Workbooks.Open Filename:="\\192.168.0.9\p&d\PDM\Solicitação de Cadastro\Common\Consulta_Produtos\CodEquivalentes.xls"
        Workbooks.Open Filename:="\\192.168.0.9\p&d\PDM\Solicitação de Cadastro\Common\Consulta_Produtos\DadosProduto.xls"
        
        Set wsOrigem = Workbooks("DadosProduto.xls").Worksheets("DadosProduto")
        Set wsOrigem2 = Workbooks("CodEquivalentes.xls").Worksheets("CodEquivalentes")
        
        wsDestino.Activate
        Application.GoTo Reference:="Tabela1"
        Selection.ClearContents
        Application.GoTo Reference:="Tabela1"
        Selection.Clear
            
        wsOrigem.Activate
        x = wsOrigem.Cells(Rows.Count, 1).End(xlUp).Row
            
        wsOrigem.ListObjects.Add(xlSrcRange, Range("$A$1:$S$" & x), , xlYes).Name = "Tabela1"
        Application.GoTo Reference:="Tabela1"
        Selection.Copy
            
        wsDestino.Activate
        Application.GoTo Reference:="Tabela1"
        ActiveSheet.Paste
        wsDestino.Range("Tabela1[[Cod. Equivalente]:[Saldo Itmax]]").Select
        Selection.Copy
    
        wsDestino.Range("E2").Select
        ActiveSheet.Paste
            
        wsDestino.Range("Tabela1[[Cod. Equivalente]:[Cod. Equivalente]]").Select
        Selection.ClearContents
        
        Workbooks("DadosProduto.xls").Close savechanges:=False
        
        wsOrigem2.Activate
        Z = wsOrigem2.Cells(Rows.Count, 1).End(xlDown).Row
        i = 2
        
        wsDestino.Activate
        wsDestino.Range("A:A").Select
        
        While i <= Z
            procurar = wsOrigem2.Cells(i, 1).Value
            valor = wsOrigem2.Cells(i, 3).Value
            Set pesquisa = Selection.Find(What:=procurar, After:=ActiveCell, LookIn:=xlValues, LookAt _
            :=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlNext, MatchCase _
            :=False, SearchFormat:=False)
            If Not pesquisa Is Nothing Then
                linha = pesquisa.Row
                If valor <> wsDestino.Cells(linha, 4).Value Then
                    If wsDestino.Cells(linha, 4).Value = "" Then
                        wsDestino.Cells(linha, 4).Value = wsDestino.Cells(linha, 4).Value & valor & "; "
                    Else
                        wsDestino.Cells(linha, 4).Value = wsDestino.Cells(linha, 4).Value & valor & "; "
                    End If
                End If
            End If
            i = i + 1
        Wend
        
        Workbooks("CodEquivalentes.xls").Close savechanges:=False
        resposta = MsgBox("Data e hora de atualização dos arquivos de dados dos produtos: " + vbCrLf + vbCrLf + "Lista de produtos " & s & "" + vbCrLf + "Códigos equivalentes " & s2 & ".", vbOKOnly + vbInformation, "Informação")
        
        Call Vicunha_Saurer
        wsDestino.Activate
        Cells(1, 1).Select
               
        Application.DisplayAlerts = True
        Application.ScreenUpdating = True
        Application.DisplayStatusBar = True
        Application.EnableEvents = True
        ActiveSheet.DisplayPageBreaks = True
        Application.Calculation = xlCalculationAutomatic
        
Exit Sub
Erro:
        Application.DisplayAlerts = True
        Application.ScreenUpdating = True
        Application.DisplayStatusBar = True
        Application.EnableEvents = True
        ActiveSheet.DisplayPageBreaks = True
        Application.Calculation = xlCalculationAutomatic
        Workbooks("DadosProduto.xls").Close savechanges:=False
        Workbooks("CodEquivalentes.xls").Close savechanges:=False
        resposta = MsgBox("Não foi possível prosseguir com a atualização dos dados dos produtos, tente novamente mais tarde!", vbOKOnly + vbCritical, "ERRO")
        ThisWorkbook.Close savechanges:=True
End Sub

Public Sub Vicunha_Saurer()
    
    Dim pesquisa As Range
    Dim wsOrigem As Worksheet
    Dim wsDestino As Worksheet
    Dim fs, f, s
    
    On Error GoTo Erro2
    
    If Dir("\\192.168.0.9\p&d\PDM\Solicitação de Cadastro\Common\Consulta_Produtos\vicunha.xlsx") = "" Then Exit Sub
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.Getfile("\\192.168.0.9\p&d\PDM\Solicitação de Cadastro\Common\Consulta_Produtos\DadosProduto.xls")
    
Exit Sub
Erro2:

End Sub
