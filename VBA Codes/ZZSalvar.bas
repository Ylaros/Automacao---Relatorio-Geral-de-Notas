Attribute VB_Name = "ZZSalvar"
Sub ZZSalvarAbas()
    Dim wbOriginal As Workbook
    Dim wbNovo As Workbook
    Dim wsContSaidas As Worksheet, wsContEntradas As Worksheet, wsContCFe As Worksheet
    Dim wsCompSaidas As Worksheet, wsCompEntradas As Worksheet, wsCompCFe As Worksheet
    Dim wsNNLsSaidas As Worksheet, wsNNLsCFe As Worksheet
    Dim dataFinal As Date, dataInicial As Date
    Dim dataFinalFormatada As String, dataInicialFormatada As String
    Dim novoNomeArquivo As String, caminhoSalvamento As String

    ' Define as referências
    Set wbOriginal = ThisWorkbook
    Set wsContEntradas = wbOriginal.Sheets("Cont-Entradas")

    ' Lê as datas da planilha Cont-Entradas, células D3 (inicial) e D4 (final)
    dataInicial = wsContEntradas.Range("D3").Value
    dataFinal = wsContEntradas.Range("E3").Value

    ' Formata as datas para o nome do arquivo
    dataInicialFormatada = Format(dataInicial, "dd-mm-yyyy")
    dataFinalFormatada = Format(dataFinal, "dd-mm-yyyy")

    ' Caminho de salvamento
    caminhoSalvamento = "Z:\18 - T.I\Relatório Geral de Notas\"

    ' Cria novo workbook e copia as planilhas desejadas
    Set wbNovo = Workbooks.Add

    ' Copia as planilhas do original para o novo workbook
    wbOriginal.Sheets("Cont-Saidas").Copy Before:=wbNovo.Sheets(1)
    wbOriginal.Sheets("Cont-Entradas").Copy After:=wbNovo.Sheets(wbNovo.Sheets.Count)
    wbOriginal.Sheets("Cont-CFe").Copy After:=wbNovo.Sheets(wbNovo.Sheets.Count)
    wbOriginal.Sheets("Comp-Saidas").Copy After:=wbNovo.Sheets(wbNovo.Sheets.Count)
    wbOriginal.Sheets("Comp-Entradas").Copy After:=wbNovo.Sheets(wbNovo.Sheets.Count)
    wbOriginal.Sheets("Comp-CFe").Copy After:=wbNovo.Sheets(wbNovo.Sheets.Count)
    wbOriginal.Sheets("NNLs-Saidas").Copy After:=wbNovo.Sheets(wbNovo.Sheets.Count)
    wbOriginal.Sheets("NNLs-CFe").Copy After:=wbNovo.Sheets(wbNovo.Sheets.Count)

    ' Exclui a planilha padrão (independente do nome)
    Application.DisplayAlerts = False
    Dim wsTemp As Worksheet
    For Each wsTemp In wbNovo.Sheets
        If wbOriginal.Sheets.Count = 1 Or wsTemp.Name Like "Planilha*" Or wsTemp.Name Like "Sheet*" Then
            wsTemp.Delete
            Exit For
        End If
    Next wsTemp
    Application.DisplayAlerts = True

    ' Monta o nome do novo arquivo
    novoNomeArquivo = "Relatório Gerais de Notas " & dataInicialFormatada & " até " & dataFinalFormatada

    ' Salva o novo workbook
    wbNovo.SaveAs fileName:=caminhoSalvamento & novoNomeArquivo & ".xlsx", FileFormat:=xlOpenXMLWorkbook

    ' Fecha o novo workbook
    wbNovo.Close SaveChanges:=False
End Sub

