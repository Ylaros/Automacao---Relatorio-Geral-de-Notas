Attribute VB_Name = "DCupom"
Sub DCupomFiscal()

    ' Desativa atualizações e alertas para melhor desempenho
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False


    Dim wsEmpresasDom As Worksheet
    Dim wsCF As Worksheet
    Dim lastRow As Long
    Dim i As Long

    ' Adiciona uma nova aba chamada "Cont-CFe"
    On Error Resume Next ' Evita erro se a aba já existir
    Set wsCF = ThisWorkbook.Sheets("Cont-CFe")
    On Error GoTo 0
    
    If wsCF Is Nothing Then
        Set wsCF = ThisWorkbook.Sheets.Add
        wsCF.Name = "Cont-CFe"
        ' Move a aba para o final
        wsCF.Move After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
        
    End If

    ' Define a aba "Empresas_Dom"
    Set wsEmpresasDom = ThisWorkbook.Sheets("Empresas_Dom")

    ' Apaga as 5 primeiras linhas de "Empresas_Dom"
    'wsEmpresasDom.Rows("1:5").Delete

    ' Encontra a última linha preenchida na coluna A
    lastRow = wsEmpresasDom.Cells(wsEmpresasDom.Rows.Count, "A").End(xlUp).Row

    ' Apaga todas as linhas não numéricas na coluna A a partir de A2
    For i = lastRow To 2 Step -1
        If Not IsNumeric(wsEmpresasDom.Cells(i, "A").Value) Then
            wsEmpresasDom.Rows(i).Delete
        End If
    Next i


    ' Copia as colunas A e G de "Empresas_Dom" para as colunas A e B de "Cont Ent-Sai"
    wsEmpresasDom.Range("A2:A" & lastRow).Copy Destination:=wsCF.Range("A3")
    wsEmpresasDom.Range("G2:G" & lastRow).Copy Destination:=wsCF.Range("B3")
    wsEmpresasDom.Range("I2:I" & lastRow).Copy Destination:=wsCF.Range("C3")

    
    'Grupo Dados
    wsCF.Range("A1").Value = "Dados Empresa"
    
    wsCF.Range("A2").Value = "Cód"
    wsCF.Range("B2").Value = "Descrição"
    wsCF.Range("C2").Value = "CNPJ"
    
    
    'Grupo Data
    wsCF.Range("D1").Value = "Data Relatório"
    
    wsCF.Range("D2").Value = "D. Inicial"
    wsCF.Range("E2").Value = "D. Final"
    
    
    'Grupo Contagem
    wsCF.Range("F1").Value = "Número de Notas"
    
    wsCF.Range("F2").Value = "Sieg Válidas"
    wsCF.Range("G2").Value = "Sieg Canceladas"
    wsCF.Range("H2").Value = "Dom Válidas"
    wsCF.Range("I2").Value = "Dom Canceladas"
    
    
    'Grupo Contabilização
    wsCF.Range("J1").Value = "Contabilização"
    
    wsCF.Range("J2").Value = "Sieg Válidas"
    wsCF.Range("K2").Value = "Dom Válidas"
    wsCF.Range("L2").Value = "Diferença"

    
    CFApagarLinhasEspecificas
    
End Sub


Private Sub CFApagarLinhasEspecificas()

    ' Desativa atualizações e alertas para melhor desempenho
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False


    Dim wsCont As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim valoresParaApagar As Variant

    ' Definindo a planilha
    Set wsCont = ThisWorkbook.Sheets("Cont-CFe")

    ' Valores que devem ser apagados
    valoresParaApagar = Array(11, 13, 15, 16, 275, 977, 9990, 9991, 9992, 9993, 9994, 9995)

    ' Encontrar a última linha com dados na coluna A
    lastRow = wsCont.Cells(wsCont.Rows.Count, "A").End(xlUp).Row

    ' Percorrer a coluna A a partir da última linha até a linha 2
    For i = lastRow To 2 Step -1
        If Not IsError(Application.Match(wsCont.Cells(i, 1).Value, valoresParaApagar, 0)) Then
            wsCont.Rows(i).Delete
        End If
    Next i

    CFContarValores
    
End Sub



Private Sub CFContarValores()

    ' Desativa atualizações e alertas para melhor desempenho
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    Dim wsContCF As Worksheet, wsCFe As Worksheet
    Dim dictD As Object
    Dim lastRowC As Long, lastRowCFe As Long
    Dim i As Long, key As String
    Dim countCFe As Long, countDev As Long
    
    ' Definir as planilhas
    Set wsContCF = ThisWorkbook.Sheets("Cont-CFe")
    Set wsCFe = ThisWorkbook.Sheets("CFe_Sieg")

    
    ' Definir o intervalo de valores nas colunas C, D e G
    lastRowC = wsContCF.Cells(wsContCF.Rows.Count, "C").End(xlUp).Row
    lastRowCFe = wsCFe.Cells(wsCFe.Rows.Count, "B").End(xlUp).Row

    
    ' Inicializar os dicionários
    Set dictD = CreateObject("Scripting.Dictionary")

    ' Preencher o dicionário com os valores da coluna D da planilha CFe_Sieg
    For i = 5 To lastRowCFe
    
        ' Verifica tanto "Autorizado o uso do CFe" quanto "Cancelamento"
        If wsCFe.Cells(i, "N").Value = "Autorizado o uso do CFe" Or wsCFe.Cells(i, "N").Value = "Cancelamento" Then
            key = wsCFe.Cells(i, "D").Value & "_" & wsCFe.Cells(i, "N").Value
            If Not dictD.Exists(key) Then
                dictD(key) = 0
            End If
            dictD(key) = dictD(key) + 1
        End If
        
    Next i
    
    ' Preencher a coluna F e G da planilha CFe-Entradas
    For i = 3 To lastRowC
        countCFe = 0
        countDev = 0

        ' Verifica a primeira chave (Entradas Autorizadas)
        key = wsContCF.Cells(i, "C").Value & "_Autorizado o uso do CFe"
        If dictD.Exists(key) Then
            countCFe = countCFe + dictD(key)
        End If
        
        ' Verifica a chave para Cancelamento
        key = wsContCF.Cells(i, "C").Value & "_Cancelamento"
        If dictD.Exists(key) Then
            countDev = countDev + dictD(key)
        End If
        
        wsContCF.Cells(i, "F").Value = countCFe
        wsContCF.Cells(i, "G").Value = countDev
    Next i
    
    CFSomarColunaJ
    
End Sub



Private Sub CFSomarColunaJ()
    
    ' Desativa atualizações e alertas para melhor desempenho
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    Dim wsContCFe As Worksheet
    Dim wsCFe As Worksheet
    Dim dictSai As Object, dictDev As Object
    Dim lastRowCCFe As Long, lastRowCFe As Long
    Dim i As Long, key As String
    Dim soma As Double

    ' Definir as planilhas
    Set wsContCFe = ThisWorkbook.Sheets("Cont-CFe")
    Set wsCFe = ThisWorkbook.Sheets("CFe_Sieg")

    ' Encontrar as últimas linhas das planilhas
    lastRowCCFe = wsContCFe.Cells(wsContCFe.Rows.Count, "C").End(xlUp).Row
    lastRowCFe = wsCFe.Cells(wsCFe.Rows.Count, "D").End(xlUp).Row

    ' Criar dicionários para armazenar somas
    Set dictSai = CreateObject("Scripting.Dictionary")
    Set dictDev = CreateObject("Scripting.Dictionary")

    ' Preencher os dicionários com as somas das colunas D e G em CFe_Sieg
    For i = 6 To lastRowCFe
        If wsCFe.Cells(i, "N").Value = "Autorizado o uso do CFe" Then
            
            key = CStr(wsCFe.Cells(i, "D").Value)
            If Not dictSai.Exists(key) Then dictSai(key) = 0
            dictSai(key) = Round(dictSai(key) + wsCFe.Cells(i, "I").Value, 2)
        End If
        
    Next i

    ' Preencher a coluna J em Cont-Saídas com as somas
    For i = 3 To lastRowCCFe
        key = CStr(wsContCFe.Cells(i, "C").Value)
        soma = 0
        If dictSai.Exists(key) Then soma = Round(soma + dictSai(key), 2)
        wsContCFe.Cells(i, "J").Value = IIf(soma <> 0, soma, 0)
    Next i

    CFSomarValoresContCF

End Sub


Private Sub CFSomarValoresContCF()

    Dim wsContCFs As Worksheet
    Dim wsCFeDom As Worksheet
    Dim lastRowContCFe As Long
    Dim lastRowCFeDom As Long
    Dim i As Long, valorC As String
    Dim dict As Object
    Dim chave As String


    ' Desativa atualizações e alertas para melhor desempenho
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ' Define as abas
    Set wsContCFs = ThisWorkbook.Sheets("Cont-CFe")
    Set wsCFeDom = ThisWorkbook.Sheets("CFs_Dom")

    ' Encontra a última linha preenchida em "Cont-CFe" coluna C
    lastRowContCFe = wsContCFs.Cells(wsContCFs.Rows.Count, "C").End(xlUp).Row

    ' Encontra a última linha preenchida em "CFs_Dom" coluna B
    lastRowCFeDom = wsCFeDom.Cells(wsCFeDom.Rows.Count, "B").End(xlUp).Row

    ' Criar o dicionário
    Set dict = CreateObject("Scripting.Dictionary")

    ' Preencher o dicionário com as somas de "CFs_Dom"
    For i = 7 To lastRowCFeDom
        If wsCFeDom.Cells(i, "F").Value <> 2 And wsCFeDom.Cells(i, "F").Value <> 7 And wsCFeDom.Cells(i, "F").Value <> -1 Then
            chave = wsCFeDom.Cells(i, "B").Value
            If Not dict.Exists(chave) Then
                dict(chave) = 0
            End If
            dict(chave) = dict(chave) + 1
        End If
    Next i

    ' Preencher a coluna H em "Cont-CFe" com as somas do dicionário
    For i = 3 To lastRowContCFe
        valorC = wsContCFs.Cells(i, "C").Value
        If dict.Exists(valorC) Then
            wsContCFs.Cells(i, "H").Value = dict(valorC)
        Else
            wsContCFs.Cells(i, "H").Value = 0
        End If
    Next i

    CFSomarValoresCF

End Sub


Private Sub CFSomarValoresCF()

    Dim wsContCFe As Worksheet
    Dim wsCFsDom As Worksheet
    Dim dictCFe As Object
    Dim lastRowContCFe As Long
    Dim lastRowCFe As Long
    Dim i As Long
    Dim key As String
    Dim countValue As Long
    

    ' Desativa atualizações e alertas para melhor desempenho
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ' Define as planilhas
    Set wsContCFe = ThisWorkbook.Sheets("Cont-CFe")
    Set wsCFsDom = ThisWorkbook.Sheets("CFs_Dom")

    ' Cria um dicionário para armazenar as contagens
    Set dictCFe = CreateObject("Scripting.Dictionary")
    
    ' Preenche o dicionário com os valores de Saidas_Dom
    lastRowCFe = wsCFsDom.Cells(wsCFsDom.Rows.Count, "B").End(xlUp).Row
    For i = 7 To lastRowCFe
        key = wsCFsDom.Cells(i, "B").Value
        
        If (wsCFsDom.Cells(i, "F").Value = "2" Or wsCFsDom.Cells(i, "F").Value = "7") Then

            If dictCFe.Exists(key) Then
                dictCFe(key) = dictCFe(key) + 1
            Else
                dictCFe.Add key, 1
            End If
        End If
    Next i

    ' Preenche a coluna H de "Cont-Saídas" com as contagens do dicionário
    lastRowContCFe = wsContCFe.Cells(wsContCFe.Rows.Count, "C").End(xlUp).Row
    For i = 3 To lastRowContCFe
        key = wsContCFe.Cells(i, "C").Value
        If dictCFe.Exists(key) Then
            wsContCFe.Cells(i, "I").Value = dictCFe(key)
        Else
            wsContCFe.Cells(i, "I").Value = 0
        End If
    Next i

    CFSomarColunaNCFeDom

End Sub






Private Sub CFSomarColunaNCFeDom()

    Dim wsContCFe As Worksheet
    Dim wsCFeDom As Worksheet
    Dim dictSomas As Object
    Dim lastRowContCFe As Long
    Dim lastRowCFeDom As Long
    Dim i As Long
    Dim key As String
    Dim soma As Double
    

    ' Desativa atualizações e alertas para melhor desempenho
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ' Define as planilhas
    Set wsContCFe = ThisWorkbook.Sheets("Cont-CFe")
    Set wsCFeDom = ThisWorkbook.Sheets("CFs_Dom")

    ' Cria um dicionário para armazenar as somas
    Set dictSomas = CreateObject("Scripting.Dictionary")
    
    ' Preenche o dicionário com as somas da coluna N de CFs_Dom
    lastRowCFeDom = wsCFeDom.Cells(wsCFeDom.Rows.Count, "B").End(xlUp).Row
    For i = 7 To lastRowCFeDom
        key = wsCFeDom.Cells(i, "B").Value
        
        If dictSomas.Exists(key) Then
            dictSomas(key) = Round(dictSomas(key) + wsCFeDom.Cells(i, "I").Value, 2)
        Else
            dictSomas.Add key, Round(wsCFeDom.Cells(i, "I").Value, 2)
        End If
    Next i

    ' Preenche a coluna K de "Cont-Saídas" com as somas do dicionário
    lastRowContCFe = wsContCFe.Cells(wsContCFe.Rows.Count, "C").End(xlUp).Row
    For i = 3 To lastRowContCFe
        key = wsContCFe.Cells(i, "C").Value
        If dictSomas.Exists(key) Then
            wsContCFe.Cells(i, "K").Value = Round(dictSomas(key), 2)
        Else
            wsContCFe.Cells(i, "K").Value = 0
        End If
    Next i

    SubtrairJMenosK

End Sub


Private Sub SubtrairJMenosK()

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    
    ' Desativa atualizações e alertas para melhor desempenho
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ' Define a planilha "Cont-Saídas"
    Set ws = ThisWorkbook.Sheets("Cont-CFe")

    ' Encontra a última linha preenchida na coluna J
    lastRow = ws.Cells(ws.Rows.Count, "J").End(xlUp).Row

    ' Loop através das linhas a partir da linha 3
    For i = 3 To lastRow
        ' Realiza a subtração de J - K e armazena o resultado na coluna L
        ws.Cells(i, "L").Value = ws.Cells(i, "J").Value - ws.Cells(i, "K").Value
    Next i




    PreencherDatasContCFe

End Sub


Private Sub PreencherDatasContCFe()
    Dim wsSIEG As Worksheet
    Dim wsContCFe As Worksheet
    Dim lastRowSIEG As Long
    Dim i As Long
    Dim dictDates As Object
    Dim key As Variant
    Dim minDate As Date
    Dim maxDate As Date
    
    ' Desativa atualizações e alertas para melhor desempenho
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ' Definir as planilhas
    Set wsSIEG = ThisWorkbook.Sheets("SIEG")
    Set wsContCFe = ThisWorkbook.Sheets("Cont-CFe")
    
    ' Encontrar a última linha preenchida na coluna C de "SIEG"
    lastRowSIEG = wsSIEG.Cells(wsSIEG.Rows.Count, "C").End(xlUp).Row
    
    ' Criar e inicializar o dicionário
    Set dictDates = CreateObject("Scripting.Dictionary")
    
    ' Preencher o dicionário com as datas da coluna C a partir da linha 5
    For i = 5 To lastRowSIEG
        If IsDate(wsSIEG.Cells(i, "C").Value) Then
            dictDates(CStr(wsSIEG.Cells(i, "C").Value)) = wsSIEG.Cells(i, "C").Value
        End If
    Next i
    
    ' Ordenar o dicionário por chaves (as datas)
    If dictDates.Count > 0 Then
        Dim sortedDates() As Variant
        sortedDates = dictDates.Items
        Call QuickSort(sortedDates, LBound(sortedDates), UBound(sortedDates))

        ' Obter a menor e a maior data
        minDate = sortedDates(LBound(sortedDates))
        maxDate = sortedDates(UBound(sortedDates))
    Else
        MsgBox "Não foram encontradas datas válidas na coluna C de 'SIEG'.", vbExclamation
        Exit Sub
    End If
    
    ' Preencher a coluna D em "Cont-CFe" com a menor data a partir de D3
    wsContCFe.Range("D3:D" & wsContCFe.Cells(wsContCFe.Rows.Count, "C").End(xlUp).Row).Value = minDate
    
    ' Preencher a coluna E em "Cont-CFe" com a maior data a partir de E3
    wsContCFe.Range("E3:E" & wsContCFe.Cells(wsContCFe.Rows.Count, "C").End(xlUp).Row).Value = maxDate

    ' Ajustar a largura das colunas para melhor visualização (opcional)
    ' wsContCFe.Columns("A:L").AutoFit

    ' Chamar a função CriarCompCFe (se necessário)
    RemoverLinhasComSomaZero
    
End Sub

' Função para ordenar o array usando QuickSort
Sub QuickSort(arr As Variant, ByVal low As Long, ByVal high As Long)
    Dim i As Long, j As Long
    Dim pivot As Variant, temp As Variant

    i = low
    j = high
    pivot = arr((low + high) \ 2)

    Do While i <= j
        Do While arr(i) < pivot
            i = i + 1
        Loop
        Do While arr(j) > pivot
            j = j - 1
        Loop
        If i <= j Then
            temp = arr(i)
            arr(i) = arr(j)
            arr(j) = temp
            i = i + 1
            j = j - 1
        End If
    Loop

    If low < j Then QuickSort arr, low, j
    If i < high Then QuickSort arr, i, high
End Sub


Private Sub RemoverLinhasComSomaZero()

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim soma As Double

    ' Desativa atualizações e alertas para melhor desempenho
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ' Definir a planilha "Cont-CFe"
    Set ws = ThisWorkbook.Sheets("Cont-CFe")

    ' Encontrar a última linha preenchida na coluna A
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    ' Loop para verificar cada linha a partir da linha 3
    For i = lastRow To 3 Step -1 ' Loop de baixo para cima para evitar problemas ao excluir linhas
        ' Somar os valores das colunas F a I
        soma = Application.WorksheetFunction.Sum(ws.Range("F" & i & ":I" & i))
        
        ' Se a soma for igual a 0, apagar a linha
        If soma = 0 Then
            ws.Rows(i).Delete
        End If
    Next i

    CriarCompCFs

End Sub






'Aqui começa notas faltantes

Private Sub CriarCompCFs()

    Dim wsCFe As Worksheet
    Dim wsCompCFs As Worksheet
    Dim dictCFs As Object
    Dim lastRowCFe As Long
    Dim i As Long, j As Long
    Dim key As String
    
    ' Desativa atualizações e alertas para melhor desempenho
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    ' Definir a planilha "CFe_Sieg"
    Set wsCFe = ThisWorkbook.Sheets("CFe_Sieg")
    
    ' Criar uma nova aba chamada "Comp-CFe"
    On Error Resume Next
    Set wsCompCFs = ThisWorkbook.Sheets("Comp-CFe")
    On Error GoTo 0
    
    If wsCompCFs Is Nothing Then
        Set wsCompCFs = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsCompCFs.Name = "Comp-CFe"
    End If

    ' Escrever os cabeçalhos na aba "Comp-CFe"
    With wsCompCFs
        .Cells(1, 1).Value = "Empresa"
        .Cells(1, 2).Value = "Descrição"
        .Cells(1, 3).Value = "CNPJ"
        .Cells(1, 4).Value = "Data"
        .Cells(1, 5).Value = "Nota"
        .Cells(1, 6).Value = "Espécie"
        .Cells(1, 7).Value = "Status"
        .Cells(1, 8).Value = "Valor Sieg"
        .Cells(1, 9).Value = "Valor Dom"
        .Cells(1, 10).Value = "Mensagem"
    End With
    
    ' Criar um dicionário para armazenar os cupons
    Set dictCFs = CreateObject("Scripting.Dictionary")
    
    ' Encontrar a última linha preenchida na aba "CFe_Sieg"
    lastRowCFe = wsCFe.Cells(wsCFe.Rows.Count, "A").End(xlUp).Row
    
    ' Preencher o dicionário com os valores da planilha "CFe_Sieg"
    j = 2 ' Iniciar na linha 2 da aba "Comp-CFe"
    
    ' Processar as linhas onde "AD" = "Sai"
    For i = 6 To lastRowCFe
        If wsCFe.Cells(i, "A").Value <> "" Then
            key = CStr(wsCFe.Cells(i, "D").Value)
            If Not dictCFs.Exists(key) Then
                dictCFs.Add key, Array(wsCFe.Cells(i, "D").Value, wsCFe.Cells(i, "C").Value, wsCFe.Cells(i, "A").Value, wsCFe.Cells(i, "I").Value)
            End If
            
            ' Copiar os valores para a aba "Comp-CFe"
            With wsCompCFs
                .Cells(j, 3).Value = wsCFe.Cells(i, "D").Value ' CNPJ
                .Cells(j, 4).Value = wsCFe.Cells(i, "C").Value ' Data
                .Cells(j, 5).Value = wsCFe.Cells(i, "A").Value ' Nota
                .Cells(j, 6).Value = "CFe"
                .Cells(j, 7).Value = wsCFe.Cells(i, "N").Value ' Status
                .Cells(j, 8).Value = wsCFe.Cells(i, "I").Value ' Valor Sieg
            End With
            j = j + 1
        End If
    Next i

    ' Ativar a aba "Comp-CFe"
    wsCompCFs.Activate
    
    PreencherCompCFsComValoresFaltantes

End Sub




Private Sub PreencherCompCFsComValoresFaltantes()
    Dim wsCompCFs As Worksheet
    Dim wsCFsDom As Worksheet
    Dim dictCompCFs As Object
    Dim dictCFsDom As Object
    Dim lastRowComp As Long, lastRowCFsDom As Long
    Dim i As Long, lastRowNew As Long
    Dim BValue As String, EValue As String, DValue As String, IValue As String
    Dim key As String
    
    ' Desativa atualizações e alertas para melhor desempenho
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    ' Definir as planilhas
    Set wsCompCFs = ThisWorkbook.Sheets("Comp-CFe")
    Set wsCFsDom = ThisWorkbook.Sheets("CFs_Dom")
    
    ' Criar dicionário para armazenar as combinações de "Comp-CFe"
    Set dictCompCFs = CreateObject("Scripting.Dictionary")
    
    ' Criar dicionário para armazenar as combinações únicas de "CFs_Dom"
    Set dictCFsDom = CreateObject("Scripting.Dictionary")
    
    ' Encontrar a última linha preenchida em "Comp-CFe"
    lastRowComp = wsCompCFs.Cells(wsCompCFs.Rows.Count, "C").End(xlUp).Row
    
    ' Preencher o dicionário com combinações de "C" e "E" em "Comp-CFe"
    For i = 2 To lastRowComp
        BValue = CStr(wsCompCFs.Cells(i, "C").Value)
        EValue = CStr(wsCompCFs.Cells(i, "D").Value)
        DValue = CStr(wsCompCFs.Cells(i, "E").Value)
        key = BValue & "|" & EValue & "|" & DValue
        
        If Not dictCompCFs.Exists(key) Then
            dictCompCFs.Add key, True
        End If
    Next i
    
    ' Encontrar a última linha preenchida em "CFs_Dom"
    lastRowCFsDom = wsCFsDom.Cells(wsCFsDom.Rows.Count, "B").End(xlUp).Row
    
    ' Preencher "Comp-CFe" com valores de "CFs_Dom" onde a combinação não é encontrada
    For i = 7 To lastRowCFsDom
        BValue = CStr(wsCFsDom.Cells(i, "B").Value)
        EValue = CStr(wsCFsDom.Cells(i, "C").Value)
        IValue = CStr(wsCFsDom.Cells(i, "D").Value)
        key = BValue & "|" & EValue & "|" & IValue
        
        ' Verificar se a combinação já foi adicionada ao dicionário de "CFs_Dom"
        If Not dictCFsDom.Exists(key) Then
            dictCFsDom.Add key, True
            
            ' Verificar se a combinação não existe em "Comp-CFe"
            If Not dictCompCFs.Exists(key) Then
                lastRowNew = wsCompCFs.Cells(wsCompCFs.Rows.Count, "C").End(xlUp).Row + 1
                wsCompCFs.Cells(lastRowNew, "C").Value = wsCFsDom.Cells(i, "B").Value
                wsCompCFs.Cells(lastRowNew, "D").Value = wsCFsDom.Cells(i, "C").Value
                wsCompCFs.Cells(lastRowNew, "E").Value = wsCFsDom.Cells(i, "D").Value
                wsCompCFs.Cells(lastRowNew, "F").Value = "CFe"
                wsCompCFs.Cells(lastRowNew, "G").Value = "Dominio"
                wsCompCFs.Cells(lastRowNew, "H").Value = "N"
            End If
        End If
    Next i

    FiltrarCompCFs

End Sub



Private Sub FiltrarCompCFs()

    Dim wsCompCFs As Worksheet
    Dim wsContCFs As Worksheet
    Dim dictContCFs As Object
    Dim lastRowComp As Long, lastRowCont As Long
    Dim i As Long, key As String
    
    ' Desativa atualizações e alertas para melhor desempenho
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ' Definir as planilhas
    Set wsCompCFs = ThisWorkbook.Sheets("Comp-CFe")
    Set wsContCFs = ThisWorkbook.Sheets("Cont-CFe")
    
    ' Criar dicionário para armazenar os valores da coluna C de "Cont-CFs"
    Set dictContCFs = CreateObject("Scripting.Dictionary")
    
    ' Encontrar as últimas linhas preenchidas em ambas as planilhas
    lastRowComp = wsCompCFs.Cells(wsCompCFs.Rows.Count, "C").End(xlUp).Row
    lastRowCont = wsContCFs.Cells(wsContCFs.Rows.Count, "C").End(xlUp).Row
    
    ' Preencher o dicionário com os valores da coluna C de "Cont-CFs"
    For i = 3 To lastRowCont ' Começa de C3 conforme solicitado
        key = CStr(wsContCFs.Cells(i, "C").Value)
        If Not dictContCFs.Exists(key) Then
            dictContCFs.Add key, True
        End If
    Next i
    
    ' Verificar e apagar linhas da aba "Comp-CFe" cujos valores de C não estão em "Cont-CFs"
    For i = lastRowComp To 2 Step -1 ' Percorrer de baixo para cima para evitar problemas ao excluir linhas
        key = CStr(wsCompCFs.Cells(i, "C").Value)
        If Not dictContCFs.Exists(key) Then
            wsCompCFs.Rows(i).Delete
        End If
    Next i

    'RemoverLinhasDuplicadas
    PreencherCompCFsComValores

End Sub




Private Sub PreencherCompCFsComValores()

    Dim wsCompCFs As Worksheet
    Dim wsCFsDom As Worksheet
    Dim dictCFsDom As Object
    Dim lastRowComp As Long, lastRowCFsDom As Long
    Dim i As Long, key As String
    
    ' Desativa atualizações e alertas para melhor desempenho
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ' Definir as planilhas
    Set wsCompCFs = ThisWorkbook.Sheets("Comp-CFe")
    Set wsCFsDom = ThisWorkbook.Sheets("CFs_Dom")
    
    ' Criar dicionário para armazenar as combinações de "CFs_Dom"
    Set dictCFsDom = CreateObject("Scripting.Dictionary")
    
    ' Encontrar as últimas linhas preenchidas em ambas as planilhas
    lastRowComp = wsCompCFs.Cells(wsCompCFs.Rows.Count, "C").End(xlUp).Row
    lastRowCFsDom = wsCFsDom.Cells(wsCFsDom.Rows.Count, "B").End(xlUp).Row
    
    ' Preencher o dicionário com as combinações de "B" e "E" em "CFs_Dom"
    For i = 7 To lastRowCFsDom ' Começa da linha 5 conforme solicitado
        key = CStr(wsCFsDom.Cells(i, "B").Value) & "|" & CStr(wsCFsDom.Cells(i, "C").Value) & "|" & CStr(wsCFsDom.Cells(i, "D").Value)
        If dictCFsDom.Exists(key) Then
            ' Se a chave já existir, soma o valor de "N" ao valor existente
            dictCFsDom(key) = dictCFsDom(key) + wsCFsDom.Cells(i, "I").Value
        Else
            dictCFsDom.Add key, wsCFsDom.Cells(i, "I").Value
        End If
    Next i
    
    ' Preencher a coluna H de "Comp-CFe" com os valores somados de "N" de "CFs_Dom" ou "N" se não encontrado
    For i = 2 To lastRowComp ' Começa da linha 2 conforme solicitado
        key = CStr(wsCompCFs.Cells(i, "C").Value) & "|" & CStr(wsCompCFs.Cells(i, "D").Value) & "|" & CStr(wsCompCFs.Cells(i, "E").Value)
        If dictCFsDom.Exists(key) Then
            wsCompCFs.Cells(i, "I").Value = dictCFsDom(key)
        Else
            wsCompCFs.Cells(i, "I").Value = "N"
        End If
    Next i

    ' Chamar a função para apagar linhas se necessário
    ApagarLinhasCompCFs

End Sub





Private Sub ApagarLinhasCompCFs()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim valorG As Double, valorH As Double
    Dim statusF As String

    ' Desativa atualizações e alertas para melhor desempenho
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ' Defina a planilha "Comp-CFe"
    Set ws = ThisWorkbook.Sheets("Comp-CFe")
    
    ' Encontre a última linha preenchida na planilha
    lastRow = ws.Cells(ws.Rows.Count, "H").End(xlUp).Row
    
    ' Desativa atualizações e alertas para melhor desempenho
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    ' Percorre as linhas de baixo para cima para evitar problemas ao deletar
    For i = lastRow To 2 Step -1
        statusF = ws.Cells(i, "G").Value
        
        ' Verifica se F não contém "Cancelamento" nem "Denegado"
        If InStr(1, statusF, "Cancelamento", vbTextCompare) = 0 And InStr(1, statusF, "Denegado", vbTextCompare) = 0 Then
            ' Tenta converter valores de G e H para números
            If IsNumeric(ws.Cells(i, "H").Value) And IsNumeric(ws.Cells(i, "I").Value) Then
                valorG = CDbl(ws.Cells(i, "H").Value)
                valorH = CDbl(ws.Cells(i, "I").Value)
                
                ' Verifica se a diferença entre G e H é 0, 0.1 ou 0.2
                If Abs(valorG - valorH) <= 0.2 Then
                    ws.Rows(i).Delete
                End If
            End If
        ' Verifica se H é 0 e se F contém "Cancelamento" ou "Denegado"
        ElseIf IsNumeric(ws.Cells(i, "I").Value) Then
            valorH = CDbl(ws.Cells(i, "I").Value)
            If valorH = 0 And (InStr(1, statusF, "Cancelamento", vbTextCompare) > 0 Or InStr(1, statusF, "Denegado", vbTextCompare) > 0) Then
                ws.Rows(i).Delete
            End If
        End If
    Next i
    
    ' Reativa atualizações e alertas
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    
    ExcluirLinhasCompCFeComBaseEmCFsDom
    
End Sub


Private Sub ExcluirLinhasCompCFeComBaseEmCFsDom()
    Dim wsComp As Worksheet, wsDom As Worksheet
    Dim lastRowComp As Long, lastRowDom As Long
    Dim i As Long
    Dim dictDom As Object
    Dim chaveDom As String, chaveComp As String

    ' Cria o dicionário
    Set dictDom = CreateObject("Scripting.Dictionary")

    ' Define as planilhas
    Set wsComp = ThisWorkbook.Sheets("Comp-CFe")
    Set wsDom = ThisWorkbook.Sheets("CFs_Dom")

    ' Encontra a última linha de cada planilha
    lastRowDom = wsDom.Cells(wsDom.Rows.Count, "B").End(xlUp).Row
    lastRowComp = wsComp.Cells(wsComp.Rows.Count, "C").End(xlUp).Row

    ' Desativa atualizações para desempenho
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ' Monta o dicionário com chaves B|D da aba CFs_Dom, onde F = -1
    For i = 2 To lastRowDom
        If Trim(wsDom.Cells(i, "B").Value) <> "" And _
           Trim(wsDom.Cells(i, "D").Value) <> "" And _
           Trim(wsDom.Cells(i, "F").Value) = "-1" Then
           
            chaveDom = Trim(wsDom.Cells(i, "B").Value) & "|" & Trim(wsDom.Cells(i, "D").Value)
            dictDom(chaveDom) = True
        End If
    Next i

    ' Percorre Comp-CFe de baixo para cima
    For i = lastRowComp To 2 Step -1
        If Trim(wsComp.Cells(i, "C").Value) <> "" And _
           Trim(wsComp.Cells(i, "E").Value) <> "" Then
           
            chaveComp = Trim(wsComp.Cells(i, "C").Value) & "|" & Trim(wsComp.Cells(i, "E").Value)
            
            If dictDom.Exists(chaveComp) Then
                wsComp.Rows(i).Delete
            End If
        End If
    Next i

    ' Reativa atualizações
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True


    CFPreencherColunaI
End Sub




Private Sub CFPreencherColunaI()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim valorG As String, valorH As String
    Dim statusF As String
    Dim valorD As Date
    Dim mensagem As String
    
    ' Desativa atualizações e alertas para melhor desempenho
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ' Defina a planilha "Comp-CFe"
    Set ws = ThisWorkbook.Sheets("Comp-CFe")
    
    ' Encontre a última linha preenchida na coluna C
    lastRow = ws.Cells(ws.Rows.Count, "C").End(xlUp).Row
    
    ' Percorre as linhas a partir da segunda linha até a última linha válida
    For i = 2 To lastRow
        valorG = ws.Cells(i, "H").Value
        valorH = ws.Cells(i, "I").Value
        statusF = ws.Cells(i, "G").Value
        valorD = ws.Cells(i, "D").Value

        ' Condições para preencher a coluna I
        If valorH = "N" And valorD > Date Then
            mensagem = "Não importada mas dentro do limite de ?? horas"
        ElseIf valorG = "N" Then
            mensagem = "Cupom não encontrado em SIEG"
        ElseIf valorH = "N" Then
            mensagem = "Cupom não encontrado em Dominio"
        ElseIf (InStr(1, statusF, "Cancelamento", vbTextCompare) > 0 Or InStr(1, statusF, "Denegado", vbTextCompare) > 0) And IsNumeric(valorH) And CDbl(valorH) <> 0 Then
            mensagem = "Cupom cancelado importado como não cancelado"
        ElseIf InStr(1, statusF, "Autorizado", vbTextCompare) > 0 And IsNumeric(valorH) And CDbl(valorH) = 0 Then
            mensagem = "Atualizar tag de cancelamento em SIEG"
        Else
            mensagem = "Cupom importado com erro"
        End If

        ' Preenche a coluna I com a mensagem apropriada
        ws.Cells(i, "J").Value = mensagem
    Next i
    
    PreencherColunasAEBCompCFs
    
End Sub




Private Sub PreencherColunasAEBCompCFs()
    Dim wsCompCFs As Worksheet
    Dim wsEmpresasDom As Worksheet
    Dim dictEmpresasDom As Object
    Dim lastRowComp As Long, lastRowEmpresas As Long
    Dim i As Long
    Dim valorC As String, chave As String
    Dim valorA As String, valorB As String
    
    ' Desativa atualizações e alertas para melhor desempenho
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    ' Definir as planilhas
    Set wsCompCFs = ThisWorkbook.Sheets("Comp-CFe")
    Set wsEmpresasDom = ThisWorkbook.Sheets("Empresas_Dom")
    
    ' Criar dicionário para armazenar os valores de "Empresas_Dom"
    Set dictEmpresasDom = CreateObject("Scripting.Dictionary")
    
    ' Encontrar a última linha preenchida em "Empresas_Dom"
    lastRowEmpresas = wsEmpresasDom.Cells(wsEmpresasDom.Rows.Count, "I").End(xlUp).Row
    
    ' Preencher o dicionário com valores de "Empresas_Dom" (colunas I, A e C)
    For i = 2 To lastRowEmpresas
        chave = CStr(wsEmpresasDom.Cells(i, "I").Value)
        valorA = wsEmpresasDom.Cells(i, "A").Value
        valorB = wsEmpresasDom.Cells(i, "G").Value
        
        ' Adicionar a chave e os valores ao dicionário
        If Not dictEmpresasDom.Exists(chave) Then
            dictEmpresasDom.Add chave, Array(valorA, valorB)
        End If
    Next i
    
    ' Encontrar a última linha preenchida em "Comp-CFe"
    lastRowComp = wsCompCFs.Cells(wsCompCFs.Rows.Count, "C").End(xlUp).Row
    
    ' Percorrer "Comp-CFe" e preencher colunas A e B com base no dicionário
    For i = 2 To lastRowComp
        valorC = CStr(wsCompCFs.Cells(i, "C").Value)
        
        ' Verificar se o valor de C existe no dicionário
        If dictEmpresasDom.Exists(valorC) Then
            wsCompCFs.Cells(i, "A").Value = dictEmpresasDom(valorC)(0) ' Preenche a coluna A
            wsCompCFs.Cells(i, "B").Value = dictEmpresasDom(valorC)(1) ' Preenche a coluna B
        End If
    Next i
    
    ' Formatando a coluna D como Data Abreviada
    wsCompCFs.Columns("D").NumberFormat = "dd/mm/yy"
    
    
    ' Ajustar a largura das colunas para melhor visualização
    wsCompCFs.Columns("A:J").AutoFit
    
   ' Ordenar os dados pela coluna A de maneira crescente, ignorando o cabeçalho
    wsCompCFs.Sort.SortFields.Clear
    wsCompCFs.Sort.SortFields.Add key:=wsCompCFs.Range("A2:A" & lastRowComp), Order:=xlAscending
    With wsCompCFs.Sort
        .SetRange wsCompCFs.Range("A1:J" & lastRowComp) ' Ajuste o intervalo conforme necessário
        .Header = xlYes ' Considerar o cabeçalho
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    
    
    ESequenciaSai.ENotasFaltantesSaidas
    
End Sub











