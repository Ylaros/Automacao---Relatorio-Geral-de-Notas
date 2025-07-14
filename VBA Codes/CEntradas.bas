Attribute VB_Name = "CEntradas"
Sub CEntrada()

    ' Desativa atualizações e alertas para melhor desempenho
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False


    Dim wsEmpresasDom As Worksheet
    Dim wsContEntSai As Worksheet
    Dim lastRow As Long
    Dim i As Long

    ' Adiciona uma nova aba chamada "Cont Ent-Sai"
    On Error Resume Next ' Evita erro se a aba já existir
    Set wsContEntSai = ThisWorkbook.Sheets("Cont-Entradas")
    On Error GoTo 0
    
    If wsContEntSai Is Nothing Then
        Set wsContEntSai = ThisWorkbook.Sheets.Add
        wsContEntSai.Name = "Cont-Entradas"
        ' Move a aba para o final
        wsContEntSai.Move After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
        
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
    wsEmpresasDom.Range("A2:A" & lastRow).Copy Destination:=wsContEntSai.Range("A3")
    wsEmpresasDom.Range("G2:G" & lastRow).Copy Destination:=wsContEntSai.Range("B3")
    wsEmpresasDom.Range("I2:I" & lastRow).Copy Destination:=wsContEntSai.Range("C3")

    
    'Grupo Dados
    wsContEntSai.Range("A1").Value = "Dados Empresa"
    
    wsContEntSai.Range("A2").Value = "Cód"
    wsContEntSai.Range("B2").Value = "Descrição"
    wsContEntSai.Range("C2").Value = "CNPJ"
    
    
    'Grupo Data
    wsContEntSai.Range("D1").Value = "Data Relatório"
    
    wsContEntSai.Range("D2").Value = "D. Inicial"
    wsContEntSai.Range("E2").Value = "D. Final"
    
    
    'Grupo Contagem
    wsContEntSai.Range("F1").Value = "Número de Notas"
    
    wsContEntSai.Range("F2").Value = "Sieg Válidas"
    wsContEntSai.Range("G2").Value = "Sieg Canceladas"
    wsContEntSai.Range("H2").Value = "Dom Válidas"
    wsContEntSai.Range("I2").Value = "Dom Canceladas"
    
    
    'Grupo Contabilização
    wsContEntSai.Range("J1").Value = "Contabilização"
    
    wsContEntSai.Range("J2").Value = "Sieg Válidas"
    wsContEntSai.Range("K2").Value = "Dom Válidas"
    wsContEntSai.Range("L2").Value = "Diferença"

    
    EApagarLinhasEspecificas
    
End Sub


Private Sub EApagarLinhasEspecificas()

    ' Desativa atualizações e alertas para melhor desempenho
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False


    Dim wsCont As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim valoresParaApagar As Variant

    ' Definindo a planilha
    Set wsCont = ThisWorkbook.Sheets("Cont-Entradas")

    ' Valores que devem ser apagados
    valoresParaApagar = Array(275, 507, 541, 977, 9990, 9991, 9992, 9993, 9994, 9995)

    ' Encontrar a última linha com dados na coluna A
    lastRow = wsCont.Cells(wsCont.Rows.Count, "A").End(xlUp).Row

    ' Percorrer a coluna A a partir da última linha até a linha 2
    For i = lastRow To 2 Step -1
        If Not IsError(Application.Match(wsCont.Cells(i, 1).Value, valoresParaApagar, 0)) Then
            wsCont.Rows(i).Delete
        End If
    Next i

    EContarValores
    
End Sub



Private Sub EContarValores()
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    Dim wsContEntradas As Worksheet, wsNFe As Worksheet, wsCTe As Worksheet
    Dim dictD As Object, dictG As Object, dictCTe As Object
    Dim lastRowC As Long, lastRowNFe As Long, lastRowCTe As Long
    Dim i As Long, key As String
    Dim countSai As Long, countDev As Long, countCTE As Long
    
    Set wsContEntradas = ThisWorkbook.Sheets("Cont-Entradas")
    Set wsNFe = ThisWorkbook.Sheets("NFe-NFCe_Sieg")
    Set wsCTe = ThisWorkbook.Sheets("CTe_Sieg")
    
    lastRowC = wsContEntradas.Cells(wsContEntradas.Rows.Count, "C").End(xlUp).Row
    lastRowNFe = wsNFe.Cells(wsNFe.Rows.Count, "B").End(xlUp).Row
    lastRowCTe = wsCTe.Cells(wsCTe.Rows.Count, "A").End(xlUp).Row
    
    Set dictD = CreateObject("Scripting.Dictionary")
    Set dictG = CreateObject("Scripting.Dictionary")
    Set dictCTe = CreateObject("Scripting.Dictionary")
    
    ' Preenche dictD com chaves para col D da NFe
    For i = 5 To lastRowNFe
        key = wsNFe.Cells(i, "D").Value & "_" & wsNFe.Cells(i, "AD").Value & "_" & wsNFe.Cells(i, "AB").Value
        If Not dictD.Exists(key) Then dictD(key) = 0
        dictD(key) = dictD(key) + 1
    Next i
    
    ' Preenche dictG com chaves para col G da NFe
    For i = 5 To lastRowNFe
        key = wsNFe.Cells(i, "G").Value & "_" & wsNFe.Cells(i, "AD").Value & "_" & wsNFe.Cells(i, "AB").Value
        If Not dictG.Exists(key) Then dictG(key) = 0
        dictG(key) = dictG(key) + 1
    Next i
    
    ' Preenche dictCTe com chaves da coluna P (CNPJ) e status normalizado
    For i = 6 To lastRowCTe
        Dim cnpjCTe As String
        Dim statusCTe As String
        Dim statusNormalizado As String
        
        cnpjCTe = Trim(wsCTe.Cells(i, "P").Value)
        statusCTe = Trim(wsCTe.Cells(i, "BE").Value)
        statusNormalizado = NormalizeText(statusCTe)
        
        If statusNormalizado Like "*autorizadoousodocte*" Then
            key = cnpjCTe & "_autorizadoousodocte"
            If Not dictCTe.Exists(key) Then dictCTe(key) = 0
            dictCTe(key) = dictCTe(key) + 1
        End If
    Next i
    
    ' Laço principal para preencher col F
    For i = 3 To lastRowC
        countSai = 0
        countDev = 0
        countCTE = 0
        
        Dim cnpjBusca As String
        cnpjBusca = wsContEntradas.Cells(i, "C").Value
        
        ' Entradas
        key = cnpjBusca & "_Ent_Autorizado o uso da NF-e"
        If dictD.Exists(key) Then countSai = countSai + dictD(key)
        
        ' Saídas
        key = cnpjBusca & "_Sai_Autorizado o uso da NF-e"
        If dictD.Exists(key) Then countSai = countSai + dictD(key)
        
        ' Devoluções (G)
        key = cnpjBusca & "_Dev_Autorizado o uso da NF-e"
        If dictG.Exists(key) Then countDev = countDev + dictG(key)
        
        ' Entradas (G)
        key = cnpjBusca & "_Ent_Autorizado o uso da NF-e"
        If dictG.Exists(key) Then countDev = countDev + dictG(key)
        
        ' CTe
        key = cnpjBusca & "_autorizadoousodocte"
        If dictCTe.Exists(key) Then countCTE = dictCTe(key)
        
        ' Preenche coluna F
        wsContEntradas.Cells(i, "F").Value = countSai + countDev + countCTE
    Next i
    
    EContarValoresComTexto
End Sub

Private Function NormalizeText(txt As String) As String
    Dim normalized As String
    normalized = LCase(txt)
    
    normalized = Replace(normalized, "á", "a")
    normalized = Replace(normalized, "à", "a")
    normalized = Replace(normalized, "ã", "a")
    normalized = Replace(normalized, "â", "a")
    normalized = Replace(normalized, "é", "e")
    normalized = Replace(normalized, "ê", "e")
    normalized = Replace(normalized, "í", "i")
    normalized = Replace(normalized, "ó", "o")
    normalized = Replace(normalized, "ô", "o")
    normalized = Replace(normalized, "õ", "o")
    normalized = Replace(normalized, "ú", "u")
    normalized = Replace(normalized, "ç", "c")
    normalized = Replace(normalized, "-", "")
    normalized = Replace(normalized, " ", "")
    
    NormalizeText = normalized
End Function




Private Sub EContarValoresComTexto()

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    Dim wsContEntradas As Worksheet, wsNFe As Worksheet, wsCTe As Worksheet
    Dim dictSai As Object, dictDev As Object, dictCTeCancel As Object
    Dim lastRowContEntradas As Long, lastRowNFe As Long, lastRowCTe As Long
    Dim i As Long, chave As String
    Dim valorSai As Long, valorDev As Long, valorCTe As Long

    Set wsContEntradas = ThisWorkbook.Sheets("Cont-Entradas")
    Set wsNFe = ThisWorkbook.Sheets("NFe-NFCe_Sieg")
    Set wsCTe = ThisWorkbook.Sheets("CTe_Sieg")

    lastRowContEntradas = wsContEntradas.Cells(wsContEntradas.Rows.Count, "C").End(xlUp).Row
    lastRowNFe = wsNFe.Cells(wsNFe.Rows.Count, "D").End(xlUp).Row
    lastRowCTe = wsCTe.Cells(wsCTe.Rows.Count, "P").End(xlUp).Row            ' pega a última linha real do CT-e

    Set dictSai = CreateObject("Scripting.Dictionary")
    Set dictDev = CreateObject("Scripting.Dictionary")
    Set dictCTeCancel = CreateObject("Scripting.Dictionary")

    ' -------- NFe / NFCe (Cancelamento ou Denegado) ----------
    For i = 5 To lastRowNFe

        '-------- Saídas (coluna D) ----------
        chave = wsNFe.Cells(i, "D").Value
        If wsNFe.Cells(i, "AD").Value <> "Dev" _
           And (InStr(1, wsNFe.Cells(i, "AB").Value, "Cancelamento") > 0 _
                Or InStr(1, wsNFe.Cells(i, "AB").Value, "Denegado") > 0) Then

            dictSai(chave) = dictSai(chave) + 1
        End If

        '-------- Devoluções / Entradas (coluna G) ----------
        chave = wsNFe.Cells(i, "G").Value
        If (wsNFe.Cells(i, "AD").Value = "Dev" Or wsNFe.Cells(i, "AD").Value = "Ent") _
           And (InStr(1, wsNFe.Cells(i, "AB").Value, "Cancelamento") > 0 _
                Or InStr(1, wsNFe.Cells(i, "AB").Value, "Denegado") > 0) Then

            dictDev(chave) = dictDev(chave) + 1
        End If
    Next i

    ' -------- CT-e  (somente Cancelamento Homologado) ----------
    For i = 6 To lastRowCTe
        Dim cnpjCTe As String, statusCTe As String, statusNorm As String

        cnpjCTe = Trim(wsCTe.Cells(i, "P").Value)
        statusCTe = Trim(wsCTe.Cells(i, "BE").Value)
        statusNorm = NormalizeText(statusCTe) ' usa a mesma função do módulo anterior

        ' Conta apenas "Cancelamento Homologado"
        If statusNorm Like "*cancelamentohomologado*" Then
            chave = cnpjCTe & "_cancelamentohomologado"
            dictCTeCancel(chave) = dictCTeCancel(chave) + 1
        ' Se ainda quiser contar Denegado, descomente abaixo
        'ElseIf statusNorm Like "*denegado*" Then
        '    chave = cnpjCTe & "_denegado"
        '    dictCTeCancel(chave) = dictCTeCancel(chave) + 1
        End If
    Next i

    ' -------- Escreve resultado na coluna G (Cont-Entradas) ----------
    Dim cell As Range
    For Each cell In wsContEntradas.Range("C3:C" & lastRowContEntradas)
        chave = cell.Value
        valorSai = dictSai(chave)
        valorDev = dictDev(chave)

        ' CT-e Cancelamento
        If dictCTeCancel.Exists(chave & "_cancelamentohomologado") Then
            valorCTe = dictCTeCancel(chave & "_cancelamentohomologado")
        Else
            valorCTe = 0
        End If

        wsContEntradas.Cells(cell.Row, "G").Value = valorSai + valorDev + valorCTe

        ' Se vazio, coloca "0"
        If wsContEntradas.Cells(cell.Row, "G").Value = "" Then
            wsContEntradas.Cells(cell.Row, "G").Value = "0"
        End If
    Next cell

    ESomarColunaJ          ' mantém sua rotina extra
End Sub



Private Sub ESomarColunaJ()
    
    ' Desativa atualizações e alertas para melhor desempenho
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    Dim wsEntradas As Worksheet
    Dim wsNFe As Worksheet
    Dim wsCTe As Worksheet
    Dim dictSai As Object, dictDev As Object, dictCTe As Object, dictCTe2 As Object
    Dim lastRowEntradas As Long, lastRowNFe As Long, lastRowCTe As Long
    Dim i As Long, key As String
    Dim soma As Double

    ' Definir as planilhas
    Set wsEntradas = ThisWorkbook.Sheets("Cont-Entradas")
    Set wsNFe = ThisWorkbook.Sheets("NFe-NFCe_Sieg")
    Set wsCTe = ThisWorkbook.Sheets("CTe_Sieg")

    ' Encontrar as últimas linhas das planilhas
    lastRowEntradas = wsEntradas.Cells(wsEntradas.Rows.Count, "C").End(xlUp).Row
    lastRowNFe = wsNFe.Cells(wsNFe.Rows.Count, "D").End(xlUp).Row
    lastRowCTe = wsCTe.Cells(wsCTe.Rows.Count, "A").End(xlUp).Row


    ' Criar dicionários para armazenar somas
    Set dictSai = CreateObject("Scripting.Dictionary")
    Set dictDev = CreateObject("Scripting.Dictionary")
    Set dictCTe = CreateObject("Scripting.Dictionary")
    Set dictCTe2 = CreateObject("Scripting.Dictionary")

    ' Preencher os dicionários com as somas das colunas D e G em NFe-NFCe_Sieg
    For i = 2 To lastRowNFe
        If wsNFe.Cells(i, "AD").Value = "Dev" And wsNFe.Cells(i, "AB").Value = "Autorizado o uso da NF-e" Then
            
            key = CStr(wsNFe.Cells(i, "G").Value)
            If Not dictSai.Exists(key) Then dictSai(key) = 0
            dictSai(key) = Round(dictSai(key) + wsNFe.Cells(i, "J").Value, 2)
        End If
        
        If wsNFe.Cells(i, "AD").Value = "Ent" And wsNFe.Cells(i, "AB").Value = "Autorizado o uso da NF-e" Then
            
            key = CStr(wsNFe.Cells(i, "G").Value)
            If Not dictSai.Exists(key) Then dictSai(key) = 0
            dictSai(key) = Round(dictSai(key) + wsNFe.Cells(i, "J").Value, 2)
        End If
        
        
        If wsNFe.Cells(i, "AD").Value <> "Dev" And wsNFe.Cells(i, "AB").Value = "Autorizado o uso da NF-e" Then
        
            key = CStr(wsNFe.Cells(i, "D").Value)
            If Not dictDev.Exists(key) Then dictDev(key) = 0
            dictDev(key) = Round(dictDev(key) + wsNFe.Cells(i, "J").Value, 2)
        End If
    Next i


    For i = 2 To lastRowCTe
            
        ' Verifica se a célula contém um número, não está vazia, e se o valor é maior ou igual a 0
'        If IsNumeric(wsCTe.Cells(i, "AA").value) And wsCTe.Cells(i, "BE").value = "Autorizado o uso do CT-e" Then
'            key = CStr(wsCTe.Cells(i, "N").value)
'
'            ' Verifica se a chave já existe no dicionário, se não, inicializa com 0
'            If Not dictCTe.exists(key) Then dictCTe(key) = 0
'
'            ' Soma o valor à chave correspondente, arredondando para 2 casas decimais
'            dictCTe(key) = Round(dictCTe(key) + wsCTe.Cells(i, "AA").value, 2)
'        End If
    
    
         ' Verifica se a célula contém um número, não está vazia, e se o valor é maior ou igual a 0
        If IsNumeric(wsCTe.Cells(i, "AA").Value) And wsCTe.Cells(i, "BE").Value = "Autorizado o uso do CT-e" Or wsCTe.Cells(i, "BE").Value = "Autorizado o uso do CTe" Or wsCTe.Cells(i, "BE").Value = "Autorizado o uso do CTe." Or wsCTe.Cells(i, "BE").Value = "Autorizado o uso do CT-e." Then
            key = CStr(wsCTe.Cells(i, "P").Value)
            
            ' Verifica se a chave já existe no dicionário, se não, inicializa com 0
            If Not dictCTe2.Exists(key) Then dictCTe2(key) = 0
            
            ' Soma o valor à chave correspondente, arredondando para 2 casas decimais
            dictCTe2(key) = Round(dictCTe2(key) + wsCTe.Cells(i, "AA").Value, 2)
        End If
       
    
    
    Next i




    ' Preencher a coluna J em Cont-Entradas com as somas
    For i = 3 To lastRowEntradas
        key = CStr(wsEntradas.Cells(i, "C").Value)
        soma = 0
        If dictSai.Exists(key) Then soma = Round(soma + dictSai(key), 2)
        If dictDev.Exists(key) Then soma = Round(soma + dictDev(key), 2)
        If dictCTe.Exists(key) Then soma = Round(soma + dictCTe(key), 2)
        If dictCTe2.Exists(key) Then soma = Round(soma + dictCTe2(key), 2)
        wsEntradas.Cells(i, "J").Value = IIf(soma <> 0, soma, 0)
    Next i

    ESomarValoresContEntradas

End Sub



Private Sub ESomarValoresContEntradas()

    Dim wsContEntradas As Worksheet
    Dim wsEntradasDom As Worksheet
    Dim lastRowContEntradas As Long
    Dim lastRowEntradasDom As Long
    Dim i As Long, j As Long
    Dim valorC As String
    Dim soma As Long
    
    
    ' Desativa atualizações e alertas para melhor desempenho
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ' Define as abas
    Set wsContEntradas = ThisWorkbook.Sheets("Cont-Entradas")
    Set wsEntradasDom = ThisWorkbook.Sheets("Entradas_Dom")

    ' Encontra a última linha preenchida em "Cont-Entradas" coluna C
    lastRowContEntradas = wsContEntradas.Cells(wsContEntradas.Rows.Count, "C").End(xlUp).Row

    ' Encontra a última linha preenchida em "Entradas_Dom" coluna B
    lastRowEntradasDom = wsEntradasDom.Cells(wsEntradasDom.Rows.Count, "B").End(xlUp).Row

    ' Loop para cada valor na coluna C de "Cont-Entradas" a partir da linha 3
    For i = 3 To lastRowContEntradas
        valorC = wsContEntradas.Cells(i, "C").Value
        soma = 0

        ' Loop para cada valor na coluna B de "Entradas_Dom"
        For j = 5 To lastRowEntradasDom
            If wsEntradasDom.Cells(j, "B").Value = valorC Then
                If (wsEntradasDom.Cells(j, "G").Value = 0 Or wsEntradasDom.Cells(j, "G").Value = 1) And _
                   (wsEntradasDom.Cells(j, "P").Value <> 2 And wsEntradasDom.Cells(j, "P").Value <> 7) Then
                    soma = soma + 1
                End If
            End If
        Next j

        ' Escreve o resultado na coluna H de "Cont-Entradas"
        wsContEntradas.Cells(i, "H").Value = soma
    Next i

    ESomarValoresContEntradasCan

End Sub

Private Sub ESomarValoresContEntradasCan()

    Dim wsContEntradas As Worksheet
    Dim wsEntradasDom As Worksheet
    Dim dictEntradas As Object
    Dim lastRowContEntradas As Long
    Dim lastRowEntradasDom As Long
    Dim i As Long
    Dim key As String
    Dim countValue As Long
    

    ' Desativa atualizações e alertas para melhor desempenho
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ' Define as planilhas
    Set wsContEntradas = ThisWorkbook.Sheets("Cont-Entradas")
    Set wsEntradasDom = ThisWorkbook.Sheets("Entradas_Dom")

    ' Cria um dicionário para armazenar as contagens
    Set dictEntradas = CreateObject("Scripting.Dictionary")
    
    ' Preenche o dicionário com os valores de Entradas_Dom
    lastRowEntradasDom = wsEntradasDom.Cells(wsEntradasDom.Rows.Count, "B").End(xlUp).Row
    For i = 5 To lastRowEntradasDom
        key = wsEntradasDom.Cells(i, "B").Value
        
        If (wsEntradasDom.Cells(i, "G").Value = "0" Or wsEntradasDom.Cells(i, "G").Value = "1") _
            And (wsEntradasDom.Cells(i, "P").Value = "2" Or wsEntradasDom.Cells(i, "P").Value = "7") Then

            If dictEntradas.Exists(key) Then
                dictEntradas(key) = dictEntradas(key) + 1
            Else
                dictEntradas.Add key, 1
            End If
        End If
    Next i

    ' Preenche a coluna H de "Cont-Entradas" com as contagens do dicionário
    lastRowContEntradas = wsContEntradas.Cells(wsContEntradas.Rows.Count, "C").End(xlUp).Row
    For i = 3 To lastRowContEntradas
        key = wsContEntradas.Cells(i, "C").Value
        If dictEntradas.Exists(key) Then
            wsContEntradas.Cells(i, "I").Value = dictEntradas(key)
        Else
            wsContEntradas.Cells(i, "I").Value = 0
        End If
    Next i

    ESomarColunaNEntradasDom

End Sub


Private Sub ESomarColunaNEntradasDom()

    Dim wsContEntradas As Worksheet
    Dim wsEntradasDom As Worksheet
    Dim dictSomas As Object
    Dim lastRowContEntradas As Long
    Dim lastRowEntradasDom As Long
    Dim i As Long
    Dim key As String
    Dim soma As Double


    ' Desativa atualizações e alertas para melhor desempenho
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ' Define as planilhas
    Set wsContEntradas = ThisWorkbook.Sheets("Cont-Entradas")
    Set wsEntradasDom = ThisWorkbook.Sheets("Entradas_Dom")

    ' Cria um dicionário para armazenar as somas
    Set dictSomas = CreateObject("Scripting.Dictionary")
    
    ' Preenche o dicionário com as somas da coluna N de Entradas_Dom
    lastRowEntradasDom = wsEntradasDom.Cells(wsEntradasDom.Rows.Count, "B").End(xlUp).Row
    For i = 5 To lastRowEntradasDom
        key = wsEntradasDom.Cells(i, "B").Value
        
        If dictSomas.Exists(key) Then
            dictSomas(key) = Round(dictSomas(key) + wsEntradasDom.Cells(i, "N").Value, 2)
        Else
            dictSomas.Add key, Round(wsEntradasDom.Cells(i, "N").Value, 2)
        End If
    Next i

    ' Preenche a coluna K de "Cont-Entradas" com as somas do dicionário
    lastRowContEntradas = wsContEntradas.Cells(wsContEntradas.Rows.Count, "C").End(xlUp).Row
    For i = 3 To lastRowContEntradas
        key = wsContEntradas.Cells(i, "C").Value
        If dictSomas.Exists(key) Then
            wsContEntradas.Cells(i, "K").Value = Round(dictSomas(key), 2)
        Else
            wsContEntradas.Cells(i, "K").Value = 0
        End If
    Next i

    ESubtrairJMenosK

End Sub


Private Sub ESubtrairJMenosK()

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    
    ' Desativa atualizações e alertas para melhor desempenho
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ' Define a planilha "Cont-Entradas"
    Set ws = ThisWorkbook.Sheets("Cont-Entradas")

    ' Encontra a última linha preenchida na coluna J
    lastRow = ws.Cells(ws.Rows.Count, "J").End(xlUp).Row

    ' Loop através das linhas a partir da linha 3
    For i = 3 To lastRow
        ' Realiza a subtração de J - K e armazena o resultado na coluna L
        ws.Cells(i, "L").Value = ws.Cells(i, "J").Value - ws.Cells(i, "K").Value
    Next i




    EPreencherDatasContEntradas

End Sub


Private Sub EPreencherDatasContEntradas()
    Dim wsSIEG As Worksheet
    Dim wsContEntradas As Worksheet
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
    Set wsContEntradas = ThisWorkbook.Sheets("Cont-Entradas")
    
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
        Call EQuickSort(sortedDates, LBound(sortedDates), UBound(sortedDates))

        ' Obter a menor e a maior data
        minDate = sortedDates(LBound(sortedDates))
        maxDate = sortedDates(UBound(sortedDates))
    Else
        MsgBox "Não foram encontradas datas válidas na coluna C de 'SIEG'.", vbExclamation
        Exit Sub
    End If
    
    ' Preencher a coluna D em "Cont-Entradas" com a menor data a partir de D3
    wsContEntradas.Range("D3:D" & wsContEntradas.Cells(wsContEntradas.Rows.Count, "C").End(xlUp).Row).Value = minDate
    
    ' Preencher a coluna E em "Cont-Entradas" com a maior data a partir de E3
    wsContEntradas.Range("E3:E" & wsContEntradas.Cells(wsContEntradas.Rows.Count, "C").End(xlUp).Row).Value = maxDate

    ' Ajustar a largura das colunas para melhor visualização (opcional)
    ' wsContEntradas.Columns("A:L").AutoFit

    ' Chamar a função CriarCompEntradas (se necessário)
    ERemoverLinhasComSomaZero
End Sub

' Função para ordenar o array usando QuickSort
Sub EQuickSort(arr As Variant, ByVal low As Long, ByVal high As Long)
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

    If low < j Then EQuickSort arr, low, j
    If i < high Then EQuickSort arr, i, high
End Sub



Private Sub ERemoverLinhasComSomaZero()

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim soma As Double
    
    ' Desativa atualizações e alertas para melhor desempenho
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ' Definir a planilha "Cont-Saidas"
    Set ws = ThisWorkbook.Sheets("Cont-Entradas")

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

    ECriarCompEntradas

End Sub



'COMPARAÇÕES DE NOTAS COMEÇA AQUI

Private Sub ECriarCompEntradas()

    Dim wsNFe As Worksheet
    Dim wsCompEntradas As Worksheet
    Dim wsCTe As Worksheet
    Dim dictEntradas As Object
    Dim dictCTe As Object, dictCTe2 As Object
    Dim dictSaidas As Object ' Declaração do dicionário dictSaidas
    Dim lastRowNFe As Long
    Dim lastRowCTe As Long
    Dim i As Long, j As Long
    Dim key As String
    
    
    ' Desativa atualizações e alertas para melhor desempenho
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    ' Definir a planilha "NFe-NFCe_Sieg"
    Set wsNFe = ThisWorkbook.Sheets("NFe-NFCe_Sieg")
    Set wsCTe = ThisWorkbook.Sheets("CTe_Sieg")

    ' Criar uma nova aba chamada "Comp-Entradas"
    On Error Resume Next
    Set wsCompEntradas = ThisWorkbook.Sheets("Comp-Entradas")
    On Error GoTo 0
    
    If wsCompEntradas Is Nothing Then
        Set wsCompEntradas = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsCompEntradas.Name = "Comp-Entradas"
    End If

    ' Escrever os cabeçalhos na aba "Comp-Entradas"
    With wsCompEntradas
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
    
    ' Criar dicionários para armazenar as Entradas, CTe e Saidas
    Set dictEntradas = CreateObject("Scripting.Dictionary")
    Set dictCTe = CreateObject("Scripting.Dictionary")
    Set dictCTe2 = CreateObject("Scripting.Dictionary")
    Set dictSaidas = CreateObject("Scripting.Dictionary") ' Inicialização do dicionário dictSaidas
    
    ' Encontrar a última linha preenchida na aba "NFe-NFCe_Sieg"
    lastRowNFe = wsNFe.Cells(wsNFe.Rows.Count, "A").End(xlUp).Row
    lastRowCTe = wsCTe.Cells(wsCTe.Rows.Count, "A").End(xlUp).Row
    
    ' Preencher o dicionário com os valores da planilha "NFe-NFCe_Sieg"
    j = 2 ' Iniciar na linha 2 da aba "Comp-Entradas"
    For i = 2 To lastRowNFe
        If wsNFe.Cells(i, "AD").Value = "Ent" Then
            key = CStr(wsNFe.Cells(i, "G").Value)
            If Not dictSaidas.Exists(key) Then
                dictSaidas.Add key, Array(wsNFe.Cells(i, "K").Value, wsNFe.Cells(i, "A").Value, wsNFe.Cells(i, "AB").Value, wsNFe.Cells(i, "J").Value)
            End If
            
            ' Copiar os valores para a aba "Comp-Saidas"
            With wsCompEntradas
                .Cells(j, 3).Value = wsNFe.Cells(i, "G").Value ' CNPJ
                .Cells(j, 4).Value = wsNFe.Cells(i, "K").Value ' Data
                .Cells(j, 5).Value = wsNFe.Cells(i, "A").Value ' Nota
                .Cells(j, 6).Value = wsNFe.Cells(i, "AE").Value ' Especie
                .Cells(j, 7).Value = wsNFe.Cells(i, "AB").Value ' Status
                .Cells(j, 8).Value = wsNFe.Cells(i, "J").Value ' Valor Sieg
            End With
            j = j + 1
        End If
    Next i

    ' Processar as linhas onde "AD" = "Dev"
    For i = 2 To lastRowNFe
        If wsNFe.Cells(i, "AD").Value = "Dev" Then
            key = CStr(wsNFe.Cells(i, "G").Value)
            If Not dictSaidas.Exists(key) Then
                dictSaidas.Add key, Array(wsNFe.Cells(i, "K").Value, wsNFe.Cells(i, "A").Value, wsNFe.Cells(i, "AB").Value, wsNFe.Cells(i, "J").Value)
            End If
            
            ' Copiar os valores para a aba "Comp-Saidas"
            With wsCompEntradas
                .Cells(j, 3).Value = wsNFe.Cells(i, "G").Value ' CNPJ
                .Cells(j, 4).Value = wsNFe.Cells(i, "K").Value ' Data
                .Cells(j, 5).Value = wsNFe.Cells(i, "A").Value ' Nota
                .Cells(j, 6).Value = wsNFe.Cells(i, "AE").Value ' Especie
                .Cells(j, 7).Value = wsNFe.Cells(i, "AB").Value ' Status
                .Cells(j, 8).Value = wsNFe.Cells(i, "J").Value ' Valor Sieg
            End With
            j = j + 1
        End If
    Next i
    
    For i = 2 To lastRowNFe
        If wsNFe.Cells(i, "AD").Value <> "Dev" And wsNFe.Cells(i, "AD").Value <> "" Then
            key = CStr(wsNFe.Cells(i, "D").Value)
            If Not dictSaidas.Exists(key) Then
                dictSaidas.Add key, Array(wsNFe.Cells(i, "K").Value, wsNFe.Cells(i, "A").Value, wsNFe.Cells(i, "AB").Value, wsNFe.Cells(i, "J").Value)
            End If
            
            ' Copiar os valores para a aba "Comp-Saidas"
            With wsCompEntradas
                .Cells(j, 3).Value = wsNFe.Cells(i, "D").Value ' CNPJ (modificado para a coluna D)
                .Cells(j, 4).Value = wsNFe.Cells(i, "K").Value ' Data
                .Cells(j, 5).Value = wsNFe.Cells(i, "A").Value ' Nota
                .Cells(j, 6).Value = wsNFe.Cells(i, "AE").Value ' Especie
                .Cells(j, 7).Value = wsNFe.Cells(i, "AB").Value ' Status
                .Cells(j, 8).Value = wsNFe.Cells(i, "J").Value ' Valor Sieg
            End With
            j = j + 1
        End If
    Next i

    
    ' Preencher o dicionário para CTe
'    For i = 2 To lastRowCTe
'        If IsNumeric(wsCTe.Cells(i, "AA").value) And wsCTe.Cells(i, "AA").value <> "" Then
'            key = CStr(wsCTe.Cells(i, "AA").value)
'            If Not dictCTe.exists(key) Then
'                dictCTe.Add key, Array(wsCTe.Cells(i, "N").value, wsCTe.Cells(i, "U").value, wsCTe.Cells(i, "AA").value)
'            End If
'
'            ' Copiar os valores para a aba "Comp-Entradas"
'            With wsCompEntradas
'                .Cells(j, 3).value = wsCTe.Cells(i, "N").value ' CNPJ
'                .Cells(j, 4).value = wsCTe.Cells(i, "D").value ' Data
'                .Cells(j, 5).value = wsCTe.Cells(i, "U").value ' Nota
'                .Cells(j, 6).value = 38
'                .Cells(j, 7).value = wsCTe.Cells(i, "BE").value ' Nota
'                .Cells(j, 8).value = wsCTe.Cells(i, "AA").value ' Valor Sieg
'            End With
'            j = j + 1
'        End If
'    Next i
    
    
    ' Preencher o dicionário para CTe
    For i = 2 To lastRowCTe
        If IsNumeric(wsCTe.Cells(i, "AA").Value) And wsCTe.Cells(i, "AA").Value <> "" Then
            key = CStr(wsCTe.Cells(i, "AA").Value)
            If Not dictCTe2.Exists(key) Then ' Corrigido para verificar o dictCTe2
                dictCTe2.Add key, Array(wsCTe.Cells(i, "P").Value, wsCTe.Cells(i, "U").Value, wsCTe.Cells(i, "AA").Value)
            End If
            
            ' Copiar os valores para a aba "Comp-Entradas"
            With wsCompEntradas
                .Cells(j, 3).Value = wsCTe.Cells(i, "P").Value ' CNPJ
                .Cells(j, 4).Value = wsCTe.Cells(i, "D").Value ' Data
                .Cells(j, 5).Value = wsCTe.Cells(i, "U").Value ' Nota
                .Cells(j, 6).Value = 38
                .Cells(j, 7).Value = wsCTe.Cells(i, "BE").Value ' Nota
                .Cells(j, 8).Value = wsCTe.Cells(i, "AA").Value ' Valor Sieg
            End With
            j = j + 1
        End If
    Next i


    ' Ativar a aba "Comp-Entradas"
    wsCompEntradas.Activate

    ' Chamar a sub-rotina para preencher valores faltantes
    EPreencherCompEntradasComValoresFaltantes

End Sub




Private Sub EPreencherCompEntradasComValoresFaltantes()
    Dim wsCompEntradas As Worksheet
    Dim wsEntradasDom As Worksheet
    Dim dictCompEntradas As Object
    Dim dictEntradasDom As Object
    Dim lastRowComp As Long, lastRowEntradasDom As Long
    Dim i As Long, lastRowNew As Long
    Dim BValue As String, EValue As String, DValue As String, FValue As String, IValue As String, TValue As String
    Dim key As String
    
    ' Desativa atualizações e alertas para melhor desempenho
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    ' Definir as planilhas
    Set wsCompEntradas = ThisWorkbook.Sheets("Comp-Entradas")
    Set wsEntradasDom = ThisWorkbook.Sheets("Entradas_Dom")
    
    ' Criar dicionário para armazenar as combinações de "Comp-Entradas"
    Set dictCompEntradas = CreateObject("Scripting.Dictionary")
    
    ' Criar dicionário para armazenar as combinações únicas de "Entradas_Dom"
    Set dictEntradasDom = CreateObject("Scripting.Dictionary")
    
    ' Encontrar a última linha preenchida em "Comp-Entradas"
    lastRowComp = wsCompEntradas.Cells(wsCompEntradas.Rows.Count, "C").End(xlUp).Row
    
    ' Preencher o dicionário com combinações de "C", "D" e "E" em "Comp-Entradas"
    For i = 2 To lastRowComp
        BValue = CStr(wsCompEntradas.Cells(i, "C").Value)
        EValue = CStr(wsCompEntradas.Cells(i, "E").Value)
        DValue = CStr(wsCompEntradas.Cells(i, "D").Value)
        FValue = CStr(wsCompEntradas.Cells(i, "F").Value)
        key = BValue & "|" & EValue & "|" & DValue & "|" & FValue
        
        If Not dictCompEntradas.Exists(key) Then
            dictCompEntradas.Add key, True
        End If
    Next i
    
    ' Encontrar a última linha preenchida em "Entradas_Dom"
    lastRowEntradasDom = wsEntradasDom.Cells(wsEntradasDom.Rows.Count, "B").End(xlUp).Row
    
    ' Preencher "Comp-Entradas" com valores de "Entradas_Dom" onde a combinação não é encontrada
    For i = 5 To lastRowEntradasDom
        BValue = CStr(wsEntradasDom.Cells(i, "B").Value)
        EValue = CStr(wsEntradasDom.Cells(i, "E").Value)
        IValue = CStr(wsEntradasDom.Cells(i, "I").Value)
        TValue = CStr(wsEntradasDom.Cells(i, "T").Value)
        key = BValue & "|" & EValue & "|" & IValue & "|" & TValue
        
        ' Verificar se a combinação já foi adicionada ao dicionário de "Entradas_Dom"
        If Not dictEntradasDom.Exists(key) Then
            dictEntradasDom.Add key, True
            
            ' Verificar se a combinação não existe em "Comp-Entradas"
            If Not dictCompEntradas.Exists(key) Then
                lastRowNew = wsCompEntradas.Cells(wsCompEntradas.Rows.Count, "C").End(xlUp).Row + 1
                wsCompEntradas.Cells(lastRowNew, "C").Value = wsEntradasDom.Cells(i, "B").Value
                wsCompEntradas.Cells(lastRowNew, "D").Value = wsEntradasDom.Cells(i, "I").Value
                wsCompEntradas.Cells(lastRowNew, "E").Value = wsEntradasDom.Cells(i, "E").Value
                wsCompEntradas.Cells(lastRowNew, "F").Value = wsEntradasDom.Cells(i, "T").Value
                wsCompEntradas.Cells(lastRowNew, "G").Value = "Dominio"
                wsCompEntradas.Cells(lastRowNew, "H").Value = "N"
            End If
        End If
    Next i

    EFiltrarCompEntradas

End Sub



Private Sub EFiltrarCompEntradas()

    Dim wsCompEntradas As Worksheet
    Dim wsContEntradas As Worksheet
    Dim dictContEntradas As Object
    Dim lastRowComp As Long, lastRowCont As Long
    Dim i As Long, key As String
    

    ' Desativa atualizações e alertas para melhor desempenho
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ' Definir as planilhas
    Set wsCompEntradas = ThisWorkbook.Sheets("Comp-Entradas")
    Set wsContEntradas = ThisWorkbook.Sheets("Cont-Entradas")
    
    ' Criar dicionário para armazenar os valores da coluna C de "Cont-Entradas"
    Set dictContEntradas = CreateObject("Scripting.Dictionary")
    
    ' Encontrar as últimas linhas preenchidas em ambas as planilhas
    lastRowComp = wsCompEntradas.Cells(wsCompEntradas.Rows.Count, "C").End(xlUp).Row
    lastRowCont = wsContEntradas.Cells(wsContEntradas.Rows.Count, "C").End(xlUp).Row
    
    ' Preencher o dicionário com os valores da coluna C de "Cont-Entradas"
    For i = 3 To lastRowCont ' Começa de C3 conforme solicitado
        key = CStr(wsContEntradas.Cells(i, "C").Value)
        If Not dictContEntradas.Exists(key) Then
            dictContEntradas.Add key, True
        End If
    Next i
    
    ' Verificar e apagar linhas da aba "Comp-Entradas" cujos valores de C não estão em "Cont-Entradas"
    For i = lastRowComp To 2 Step -1 ' Percorrer de baixo para cima para evitar problemas ao excluir linhas
        key = CStr(wsCompEntradas.Cells(i, "C").Value)
        If Not dictContEntradas.Exists(key) Then
            wsCompEntradas.Rows(i).Delete
        End If
    Next i

    EPreencherCompEntradasComValores

End Sub



Private Sub EPreencherCompEntradasComValores()

    Dim wsCompEntradas As Worksheet
    Dim wsEntradasDom As Worksheet
    Dim dictEntradasDom As Object
    Dim lastRowComp As Long, lastRowEntradasDom As Long
    Dim i As Long, key As String
    
    ' Desativa atualizações e alertas para melhor desempenho
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ' Definir as planilhas
    Set wsCompEntradas = ThisWorkbook.Sheets("Comp-Entradas")
    Set wsEntradasDom = ThisWorkbook.Sheets("Entradas_Dom")
    
    ' Criar dicionário para armazenar as combinações de "Entradas_Dom"
    Set dictEntradasDom = CreateObject("Scripting.Dictionary")
    
    ' Encontrar as últimas linhas preenchidas em ambas as planilhas
    lastRowComp = wsCompEntradas.Cells(wsCompEntradas.Rows.Count, "C").End(xlUp).Row
    lastRowEntradasDom = wsEntradasDom.Cells(wsEntradasDom.Rows.Count, "B").End(xlUp).Row
    
    ' Preencher o dicionário com as combinações de "B", "E" e "I" em "Entradas_Dom"
    For i = 5 To lastRowEntradasDom ' Começa da linha 5 conforme solicitado
        key = CStr(wsEntradasDom.Cells(i, "B").Value) & "|" & CStr(wsEntradasDom.Cells(i, "E").Value) & "|" & CStr(wsEntradasDom.Cells(i, "I").Value) & "|" & CStr(wsEntradasDom.Cells(i, "T").Value)
        If dictEntradasDom.Exists(key) Then
            ' Se a chave já existir, soma o valor de "N" ao valor existente
            dictEntradasDom(key) = dictEntradasDom(key) + wsEntradasDom.Cells(i, "N").Value
        Else
            dictEntradasDom.Add key, wsEntradasDom.Cells(i, "N").Value
        End If
    Next i
    
    ' Preencher a coluna I de "Comp-Entradas" com os valores somados de "N" de "Entradas_Dom" ou "N" se não encontrado
    For i = 2 To lastRowComp ' Começa da linha 2 conforme solicitado
        key = CStr(wsCompEntradas.Cells(i, "C").Value) & "|" & CStr(wsCompEntradas.Cells(i, "E").Value) & "|" & CStr(wsCompEntradas.Cells(i, "D").Value) & "|" & CStr(wsCompEntradas.Cells(i, "F").Value)
        If dictEntradasDom.Exists(key) Then
            wsCompEntradas.Cells(i, "I").Value = dictEntradasDom(key)
        Else
            wsCompEntradas.Cells(i, "I").Value = "N"
        End If
    Next i

    ' Chamar a função para apagar linhas se necessário
    EApagarLinhasCompEntradas

End Sub



Private Sub EApagarLinhasCompEntradas()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim valorG As Double, valorH As Double
    Dim statusG As String


    Application.ScreenUpdating = False
    Application.DisplayAlerts = False


    ' Defina a planilha "Comp-Entradas"
    Set ws = ThisWorkbook.Sheets("Comp-Entradas")
    
    ' Encontre a última linha preenchida na planilha
    lastRow = ws.Cells(ws.Rows.Count, "H").End(xlUp).Row
    
    ' Desativa atualizações e alertas para melhor desempenho
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    ' Percorre as linhas de baixo para cima para evitar problemas ao deletar
    For i = lastRow To 2 Step -1
        statusG = ws.Cells(i, "G").Value
        
        ' Verifica se G não contém "Cancelamento" nem "Denegado"
        If InStr(1, statusG, "Cancelamento", vbTextCompare) = 0 And InStr(1, statusG, "Denegado", vbTextCompare) = 0 Then
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
            If valorH = 0 And (InStr(1, statusG, "Cancelamento", vbTextCompare) > 0 Or InStr(1, statusG, "Denegado", vbTextCompare) > 0) Then
                ws.Rows(i).Delete
            End If
        End If
    Next i
    
    ' Reativa atualizações e alertas
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    
    EPreencherColunaI
    
End Sub


Private Sub EPreencherColunaI()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim valorG As String, valorH As String
    Dim statusF As String
    Dim mensagem As String
    

    ' Desativa atualizações e alertas para melhor desempenho
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ' Defina a planilha "Comp-Entradas"
    Set ws = ThisWorkbook.Sheets("Comp-Entradas")
    
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
            mensagem = "Nota não encontrada em SIEG"
        ElseIf valorH = "N" Then
            mensagem = "Nota não encontrada em Dominio"
        ElseIf (InStr(1, statusF, "Cancelamento", vbTextCompare) > 0 Or InStr(1, statusF, "Denegado", vbTextCompare) > 0) And IsNumeric(valorH) And CDbl(valorH) <> 0 Then
            mensagem = "Nota cancelada importada como não cancelada"
        ElseIf InStr(1, statusF, "Autorizado", vbTextCompare) > 0 And IsNumeric(valorH) And CDbl(valorH) = 0 Then
            mensagem = "Atualizar tag de cancelamento em SIEG"
        Else
            mensagem = "Nota importada com erro"
        End If

        ' Preenche a coluna I com a mensagem apropriada
        ws.Cells(i, "J").Value = mensagem
    Next i
    
    EPreencherColunasAEBCompEntradas
    
End Sub



Private Sub EPreencherColunasAEBCompEntradas()
    Dim wsCompEntradas As Worksheet
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
    Set wsCompEntradas = ThisWorkbook.Sheets("Comp-Entradas")
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
    
    ' Encontrar a última linha preenchida em "Comp-Entradas"
    lastRowComp = wsCompEntradas.Cells(wsCompEntradas.Rows.Count, "C").End(xlUp).Row
    
    ' Percorrer "Comp-Entradas" e preencher colunas A e B com base no dicionário
    For i = 2 To lastRowComp
        valorC = CStr(wsCompEntradas.Cells(i, "C").Value)
        
        ' Verificar se o valor de C existe no dicionário
        If dictEmpresasDom.Exists(valorC) Then
            wsCompEntradas.Cells(i, "A").Value = dictEmpresasDom(valorC)(0) ' Preenche a coluna A
            wsCompEntradas.Cells(i, "B").Value = dictEmpresasDom(valorC)(1) ' Preenche a coluna B
        End If
    Next i
    
    ' Formatando a coluna D como Data Abreviada
    wsCompEntradas.Columns("D").NumberFormat = "dd/mm/yy"
    
    
    ' Ajustar a largura das colunas para melhor visualização
    wsCompEntradas.Columns("A:J").AutoFit
    
   ' Ordenar os dados pela coluna A de maneira crescente, ignorando o cabeçalho
    wsCompEntradas.Sort.SortFields.Clear
    wsCompEntradas.Sort.SortFields.Add key:=wsCompEntradas.Range("A2:A" & lastRowComp), Order:=xlAscending
    With wsCompEntradas.Sort
        .SetRange wsCompEntradas.Range("A1:J" & lastRowComp) ' Ajuste o intervalo conforme necessário
        .Header = xlYes ' Considerar o cabeçalho
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    
    
    
    DCupom.DCupomFiscal
    
End Sub



