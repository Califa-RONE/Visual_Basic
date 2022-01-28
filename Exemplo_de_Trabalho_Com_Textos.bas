Attribute VB_Name = "Módulo1"
Public Sub Vicunha_Saurer()

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    Dim pesquisa As Range
    Dim wsOrigem As Worksheet
    Dim wsDestino As Worksheet
    Dim wsPesquisa As Worksheet
    Dim testPos As Integer
    Dim fs, f, s
    
    codiA = "./-+=',;:()[]{}^~><\|!@#$%&&*§ªº° "
    codiB = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ"
    
    'On Error GoTo Erro2
    
    If Dir("\\192.168.0.9\p&d\PDM\Solicitação de Cadastro\Common\Consulta_Produtos\vicunha.xlsx") = "" Then Exit Sub
    
    Workbooks.Open Filename:="\\192.168.0.9\p&d\PDM\Solicitação de Cadastro\Common\Consulta_Produtos\vicunha.xlsx"
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.Getfile("\\192.168.0.9\p&d\PDM\Solicitação de Cadastro\Common\Consulta_Produtos\vicunha.xlsx")
    
    s = f.DateCreated
    
    Set wsOrigem = Workbooks("vicunha.xlsx").Worksheets("Sheet1")
    Set wsDestino = Workbooks("Consulta_Produtos.xlsm").Worksheets("Vicunha")
    Set wsPesquisa = Workbooks("Consulta_Produtos.xlsm").Worksheets("Consulta_Produtos")
    
    wsDestino.Activate
    
    If s <> wsDestino.Cells(1, 16).Value Then
        wsDestino.Cells(1, 16).Value = s
    Else
        Workbooks("vicunha.xlsx").Close savechanges:=False
        Application.ScreenUpdating = True
        Application.DisplayAlerts = True
        Exit Sub
    End If
    
    y = wsDestino.Cells(Rows.Count, 1).End(xlUp).Row
    If y <> 1 Then
        wsDestino.Range(Cells(2, 1), Cells(y, 15)).Clear
    End If
    
    wsOrigem.Activate
    wsOrigem.AutoFilter.Sort.SortFields.Clear
    wsOrigem.AutoFilter.Sort.SortFields.Add2 Key:=Range("L1"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With wsOrigem.AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    X = wsOrigem.Cells(Rows.Count, 1).End(xlUp).Row
    
    wsOrigem.Range(Cells(2, 1), Cells(X, 12)).Copy
    wsDestino.Activate
    wsDestino.Range(Cells(2, 1), Cells(X, 12)).PasteSpecial xlPasteAll
    
    For i = 2 To X
        testPos = InStr(wsOrigem.Cells(i, 6).Value, "PC")
        If testPos = 0 Then
            testPos = InStr(wsOrigem.Cells(i, 6).Value, "PP")
            If testPos = 0 Then GoTo continuai
        End If
        codigo_vicunha = Mid(wsOrigem.Cells(i, 6).Value, testPos, 10)
        wsDestino.Cells(i, 13).Value = codigo_vicunha
        
'        wsPesquisa.Activate
'        wsPesquisa.Range("A:A").Select
'        Set pesquisa = Selection.Find(What:=codigo_vicunha, After:=ActiveCell, LookIn:=xlValues, _
'        LookAt:=xlPart, SearchOrder:=xlByColumns, SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False)
'        If Not pesquisa Is Nothing Then
'            linha = pesquisa.Row
'            wsDestino.Cells(i, 15).Value = wsPesquisa.Cells(linha, 2).Value
'            wsDestino.Cells(i, 14).Value = wsPesquisa.Cells(linha, 1).Value
'        End If
continuai:
        codigo_vicunha = ""
        omega = Len(wsOrigem.Cells(i, 6).Value)
        testPos = InStr(wsOrigem.Cells(i, 6).Value, "REF")
        If testPos = 0 Then GoTo FINAL
        alpha = testPos + 3
volta:
        epsilon = Mid(wsOrigem.Cells(i, 6).Value, alpha, 1)
        If IsNumeric(epsilon) = False Then
            If alpha = omega Then GoTo continuai2
            For h = 1 To Len(codiA)
                If epsilon = Mid(codiA, h, 1) Then
                    alpha = alpha + 1
                    GoTo volta
                End If
            Next h
            For h = 1 To Len(codiB)
                If epsilon = Mid(codiB, h, 1) Then
                    alpha = alpha + 1
                    GoTo volta
                End If
            Next h
        Else
            If alpha = omega Then GoTo continuai2
            codigo_vicunha = codigo_vicunha & epsilon
            alpha = alpha + 1
            GoTo volta
        End If
continuai2:
        wsDestino.Cells(i, 14).Value = codigo_vicunha
FINAL:
    Next i
    
    Workbooks("vicunha.xlsx").Close savechanges:=False
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
Exit Sub
Erro2:

End Sub
