Attribute VB_Name = "BSaidas"
Sub BSaida()

    ' Desativa atualizações e alertas para melhor desempenho
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False


    Dim wsEmpresasDom As Worksheet
    Dim wsContEntSai As Worksheet
    Dim lastRow As Long
    Dim i As Long

    ' Adiciona uma nova aba chamada "Cont Ent-Sai"
    On Error Resume Next ' Evita erro se a aba já existir
    Set wsContEntSai = ThisWorkbook.Sheets("Cont-Saidas")
    On Error GoTo 0
    
    If wsContEntSai Is Nothing Then
        Set wsContEntSai = ThisWorkbook.Sheets.Add
        wsContEntSai.Name = "Cont-Saidas"
        ' Move a aba para o final
        wsContEntSai.Move After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
        
    End If

    ' Define a aba "Empresas_Dom"
    Set wsEmpresasDom = ThisWorkbook.Sheets("Empresas_Dom")

    ' Apaga as 5 primeiras linhas de "Empresas_Dom"
    wsEmpresasDom.Rows("1:5").Delete

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

    
    ApagarLinhasEspecificas
    
End Sub


Private Sub ApagarLinhasEspecificas()

    ' Desativa atualizações e alertas para melhor desempenho
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False


    Dim wsCont As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim valoresParaApagar As Variant

    ' Definindo a planilha
    Set wsCont = ThisWorkbook.Sheets("Cont-Saidas")

    ' Valores que devem ser apagados 416,
    valoresParaApagar = Array(275, 541, 977, 9990, 9991, 9992, 9993, 9994, 9995)

    ' Encontrar a última linha com dados na coluna A
    lastRow = wsCont.Cells(wsCont.Rows.Count, "A").End(xlUp).Row

    ' Percorrer a coluna A a partir da última linha até a linha 2
    For i = lastRow To 2 Step -1
        If Not IsError(Application.Match(wsCont.Cells(i, 1).Value, valoresParaApagar, 0)) Then
            wsCont.Rows(i).Delete
        End If
    Next i

    ContarValores
    
End Sub





Private Sub ContarValores()
    
    
    ' Desativa atualizações e alertas para melhor desempenho
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    
    Dim wsContSaidas As Worksheet, wsNFe As Worksheet
    Dim dictD As Object, dictG As Object, dictCTe As Object, dictCTe2 As Object
    Dim lastRowC As Long, lastRowNFe As Long, lastRowCTe As Long
    Dim i As Long, key As String
    Dim countSai As Long, countDev As Long, countCTE As Long
    
    ' Definir as planilhas
    Set wsContSaidas = ThisWorkbook.Sheets("Cont-Saidas")
    Set wsNFe = ThisWorkbook.Sheets("NFe-NFCe_Sieg")
    Set wsCTe = ThisWorkbook.Sheets("CTe_Sieg")
    
    ' Definir o intervalo de valores nas colunas C, D e G
    lastRowC = wsContSaidas.Cells(wsContSaidas.Rows.Count, "C").End(xlUp).Row
    lastRowNFe = wsNFe.Cells(wsNFe.Rows.Count, "B").End(xlUp).Row
    lastRowCTe = wsCTe.Cells(wsCTe.Rows.Count, "A").End(xlUp).Row
    
    ' Inicializar os dicionários
    Set dictD = CreateObject("Scripting.Dictionary")
    Set dictG = CreateObject("Scripting.Dictionary")
    Set dictCTe = CreateObject("Scripting.Dictionary")
    Set dictCTe2 = CreateObject("Scripting.Dictionary")
    
    ' Preencher o dicionário com os valores da coluna D da planilha NFe-NFCe_Sieg
    For i = 5 To lastRowNFe
        key = wsNFe.Cells(i, "D").Value & "_" & wsNFe.Cells(i, "AD").Value & "_" & wsNFe.Cells(i, "AB").Value
        If Not dictD.Exists(key) Then
            dictD(key) = 0
        End If
        dictD(key) = dictD(key) + 1
    Next i
    
    ' Preencher o dicionário com os valores da coluna G da planilha NFe-NFCe_Sieg
    For i = 5 To lastRowNFe
        key = wsNFe.Cells(i, "G").Value & "_" & wsNFe.Cells(i, "AD").Value & "_" & wsNFe.Cells(i, "AB").Value
        If Not dictG.Exists(key) Then
            dictG(key) = 0
        End If
        dictG(key) = dictG(key) + 1
    Next i
    
    
    ' Preencher o dicionário com os valores da coluna CNPJ da planilha CTE
    For i = 6 To lastRowCTe
        key = wsCTe.Cells(i, "N").Value & "_" & wsCTe.Cells(i, "BE").Value
        If Not dictCTe.Exists(key) Then
            dictCTe(key) = 0
        End If
        dictCTe(key) = dictCTe(key) + 1
    Next i
    
    
    
    
    
    
    ' Preencher a coluna F da planilha Cont-Saídas
    For i = 3 To lastRowC
        countSai = 0
        countDev = 0
        countCTE = 0
        key = wsContSaidas.Cells(i, "C").Value & "_Sai_Autorizado o uso da NF-e"
        
'        If dictD.exists(key) Then
'            countSai = dictD(key)
'        End If
        If dictG.Exists(key) Then
            countSai = countSai + dictG(key)
        End If
        
        key = wsContSaidas.Cells(i, "C").Value & "_Dev_Autorizado o uso da NF-e"
        
        If dictD.Exists(key) Then
            countDev = dictD(key)
        End If
        
        
         ' Verifica a terceira chave (CTE)
        key = wsContSaidas.Cells(i, "C").Value & "_Autorizado o uso do CT-e"
        If dictCTe.Exists(key) Then
            countCTE = countCTE + dictCTe(key)
        End If
        
        
        ' Verifica a terceira chave (CTE)
        key = wsContSaidas.Cells(i, "C").Value & "_Autorizado o uso do CTe"
        If dictCTe.Exists(key) Then
            countCTE = countCTE + dictCTe(key)
        End If
        
        
        ' Verifica a terceira chave (CTE)
        key = wsContSaidas.Cells(i, "C").Value & "_Autorizado o uso do CT-e"
        If dictCTe2.Exists(key) Then
            countCTE = countCTE + dictCTe2(key)
        End If
  
        
        
        wsContSaidas.Cells(i, "F").Value = countSai + countDev + countCTE
    Next i
    


    
    
    ContarValoresComTexto
    
End Sub

Private Sub ContarValoresComTexto()
    
    ' Desativa atualizações e alertas para melhor desempenho
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    
    Dim wsContEntradas As Worksheet
    Dim wsNFe As Worksheet
    Dim wsCTe As Worksheet
    Dim lastRowContEntradas As Long
    Dim lastRowNFe As Long
    Dim lastRowCTe As Long
    Dim dictSai As Object
    Dim dictDev As Object
    Dim dictCTe As Object
    Dim dictCTe2 As Object
    Dim i As Long
    Dim chave As String
    Dim valorSai As Long
    Dim valorDev As Long
    Dim valorCTe As Long
    Dim valorCTe2 As Long
    Dim rng As Range
    Dim cell As Range

    Set wsContSaidas = ThisWorkbook.Sheets("Cont-Saidas")
    Set wsNFe = ThisWorkbook.Sheets("NFe-NFCe_Sieg")
    Set wsCTe = ThisWorkbook.Sheets("CTe_Sieg")

    ' Encontrar a última linha da coluna C em Cont-Saídas e NFe-NFCe_Sieg
    lastRowContSaidas = wsContSaidas.Cells(wsContSaidas.Rows.Count, "C").End(xlUp).Row
    lastRowNFe = wsNFe.Cells(wsNFe.Rows.Count, "D").End(xlUp).Row
    lastRowCTe = wsNFe.Cells(wsNFe.Rows.Count, "D").End(xlUp).Row

    ' Criar dicionários para armazenar as contagens
    Set dictSai = CreateObject("Scripting.Dictionary")
    Set dictDev = CreateObject("Scripting.Dictionary")
    Set dictCTe = CreateObject("Scripting.Dictionary")
    Set dictCTe2 = CreateObject("Scripting.Dictionary")

    ' Preencher os dicionários com as contagens
    For i = 5 To lastRowNFe
        chave = wsNFe.Cells(i, "D").Value
'        If wsNFe.Cells(i, "AD").Value = "Sai" And _
'           (InStr(1, wsNFe.Cells(i, "AB").Value, "Cancelamento") > 0 Or _
'            InStr(1, wsNFe.Cells(i, "AB").Value, "Denegado") > 0) Then
'            If dictSai.exists(chave) Then
'                dictSai(chave) = dictSai(chave) + 1
'            Else
'                dictSai(chave) = 1
'            End If
        If wsNFe.Cells(i, "AD").Value = "Dev" And _
           (InStr(1, wsNFe.Cells(i, "AB").Value, "Cancelamento") > 0 Or _
            InStr(1, wsNFe.Cells(i, "AB").Value, "Denegado") > 0) Then
            If dictDev.Exists(chave) Then
                dictDev(chave) = dictDev(chave) + 1
            Else
                dictDev(chave) = 1
            End If
        End If

        chave = wsNFe.Cells(i, "G").Value
        If wsNFe.Cells(i, "AD").Value = "Sai" And _
           (InStr(1, wsNFe.Cells(i, "AB").Value, "Cancelamento") > 0 Or _
            InStr(1, wsNFe.Cells(i, "AB").Value, "Denegado") > 0) Then
            If dictSai.Exists(chave) Then
                dictSai(chave) = dictSai(chave) + 1
            Else
                dictSai(chave) = 1
            End If
'        ElseIf wsNFe.Cells(i, "AD").Value = "Dev" And _
'           (InStr(1, wsNFe.Cells(i, "AB").Value, "Cancelamento") > 0 Or _
'            InStr(1, wsNFe.Cells(i, "AB").Value, "Denegado") > 0) Then
'            If dictDev.exists(chave) Then
'                dictDev(chave) = dictDev(chave) + 1
'            Else
'                dictDev(chave) = 1
'            End If
        End If
    Next i



    ' Preencher os dicionários com as contagens
    For i = 6 To lastRowCTe
        chave = wsCTe.Cells(i, "N").Value
        If (InStr(1, wsCTe.Cells(i, "BE").Value, "Cancelamento") > 0 Or _
            InStr(1, wsCTe.Cells(i, "BE").Value, "Denegado") > 0) Then
            If dictCTe.Exists(chave) Then
                dictCTe(chave) = dictCTe(chave) + 1
            Else
                dictCTe(chave) = 1
            End If
        End If
        
        
'        chave = wsCTe.Cells(i, "P").value
'        If (InStr(1, wsCTe.Cells(i, "BE").value, "Cancelamento") > 0 Or _
'            InStr(1, wsCTe.Cells(i, "BE").value, "Denegado") > 0) Then
'            If dictCTe2.exists(chave) Then
'                dictCTe2(chave) = dictCTe2(chave) + 1
'            Else
'                dictCTe2(chave) = 1
'            End If
'        End If
        
        
    Next i




    ' Preencher a coluna G em Cont-Saídas com as contagens
    For Each cell In wsContSaidas.Range("C3:C" & lastRowContSaidas)
        chave = cell.Value
        valorSai = 0
        valorDev = 0
        valorCTe = 0
        
        If dictSai.Exists(chave) Then valorSai = dictSai(chave)
        If dictDev.Exists(chave) Then valorDev = dictDev(chave)
        If dictCTe.Exists(chave) Then valorCTe = dictCTe(chave) + dictCTe2(chave)
        
        wsContSaidas.Cells(cell.Row, "G").Value = valorSai + valorDev + valorCTe
        
        ' Se a soma for zero, preencher com "0"
        If wsContSaidas.Cells(cell.Row, "G").Value = 0 Then
            wsContSaidas.Cells(cell.Row, "G").Value = "0"
        End If
    Next cell
    
    SomarColunaJ
    
End Sub


Private Sub SomarColunaJ()
    
    ' Desativa atualizações e alertas para melhor desempenho
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    Dim wsSaidas As Worksheet
    Dim wsNFe As Worksheet
    Dim wsCTe As Worksheet
    Dim dictSai As Object, dictDev As Object, dictCTe As Object, dictCTe2 As Object
    Dim lastRowSaidas As Long, lastRowNFe As Long, lastRowCTe As Long
    Dim i As Long, key As String
    Dim soma As Double

    ' Definir as planilhas
    Set wsSaidas = ThisWorkbook.Sheets("Cont-Saidas")
    Set wsNFe = ThisWorkbook.Sheets("NFe-NFCe_Sieg")
    Set wsCTe = ThisWorkbook.Sheets("CTe_Sieg")

    ' Encontrar as últimas linhas das planilhas
    lastRowSaidas = wsSaidas.Cells(wsSaidas.Rows.Count, "C").End(xlUp).Row
    lastRowNFe = wsNFe.Cells(wsNFe.Rows.Count, "D").End(xlUp).Row
    lastRowCTe = wsCTe.Cells(wsCTe.Rows.Count, "A").End(xlUp).Row

    ' Criar dicionários para armazenar somas
    Set dictSai = CreateObject("Scripting.Dictionary")
    Set dictDev = CreateObject("Scripting.Dictionary")
    Set dictCTe = CreateObject("Scripting.Dictionary")
    Set dictCTe2 = CreateObject("Scripting.Dictionary")

    ' Preencher os dicionários com as somas das colunas D e G em NFe-NFCe_Sieg
    For i = 2 To lastRowNFe
        If wsNFe.Cells(i, "AD").Value = "Sai" And wsNFe.Cells(i, "AB").Value = "Autorizado o uso da NF-e" Then
            
            key = CStr(wsNFe.Cells(i, "G").Value)
            If Not dictSai.Exists(key) Then dictSai(key) = 0
            dictSai(key) = Round(dictSai(key) + wsNFe.Cells(i, "J").Value, 2)
        End If
        
        If wsNFe.Cells(i, "AD").Value = "Dev" And wsNFe.Cells(i, "AB").Value = "Autorizado o uso da NF-e" Then
            key = CStr(wsNFe.Cells(i, "D").Value)
            If Not dictDev.Exists(key) Then dictDev(key) = 0
            dictDev(key) = Round(dictDev(key) + wsNFe.Cells(i, "J").Value, 2)
        End If
    Next i

    For i = 2 To lastRowCTe
            
        ' Verifica se a célula contém um número, não está vazia, e se o valor é maior ou igual a 0
        If IsNumeric(wsCTe.Cells(i, "AA").Value) And wsCTe.Cells(i, "BE").Value = "Autorizado o uso do CT-e" Then
            key = CStr(wsCTe.Cells(i, "N").Value)

            ' Verifica se a chave já existe no dicionário, se não, inicializa com 0
            If Not dictCTe.Exists(key) Then dictCTe(key) = 0

            ' Soma o valor à chave correspondente, arredondando para 2 casas decimais
            dictCTe(key) = Round(dictCTe(key) + wsCTe.Cells(i, "AA").Value, 2)
        End If
    
    
         ' Verifica se a célula contém um número, não está vazia, e se o valor é maior ou igual a 0
'        If IsNumeric(wsCTe.Cells(i, "AA").value) And wsCTe.Cells(i, "BE").value = "Autorizado o uso do CT-e" Then
'            key = CStr(wsCTe.Cells(i, "P").value)
'
'            ' Verifica se a chave já existe no dicionário, se não, inicializa com 0
'            If Not dictCTe2.exists(key) Then dictCTe2(key) = 0
'
'            ' Soma o valor à chave correspondente, arredondando para 2 casas decimais
'            dictCTe2(key) = Round(dictCTe2(key) + wsCTe.Cells(i, "AA").value, 2)
'        End If
       
    
    
    Next i



    ' Preencher a coluna J em Cont-Entradas com as somas
    For i = 3 To lastRowSaidas
        key = CStr(wsSaidas.Cells(i, "C").Value)
        soma = 0
        If dictSai.Exists(key) Then soma = Round(soma + dictSai(key), 2)
        If dictDev.Exists(key) Then soma = Round(soma + dictDev(key), 2)
        If dictCTe.Exists(key) Then soma = Round(soma + dictCTe(key), 2)
        If dictCTe2.Exists(key) Then soma = Round(soma + dictCTe2(key), 2)
        wsSaidas.Cells(i, "J").Value = IIf(soma <> 0, soma, 0)
    Next i

    SomarValoresContSaidas

End Sub



Private Sub SomarValoresContSaidas()

    Dim wsContSaidas As Worksheet
    Dim wsSaidasDom As Worksheet
    Dim lastRowContSaidas As Long
    Dim lastRowSaidasDom As Long
    Dim i As Long
    Dim valorC As String
    Dim dict As Object
    Dim chave As String
    
    ' Desativa atualizações e alertas para melhor desempenho
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ' Define as abas
    Set wsContSaidas = ThisWorkbook.Sheets("Cont-Saidas")
    Set wsSaidasDom = ThisWorkbook.Sheets("Saidas_Dom")

    ' Encontra a última linha preenchida em "Cont-Saídas" coluna C
    lastRowContSaidas = wsContSaidas.Cells(wsContSaidas.Rows.Count, "C").End(xlUp).Row

    ' Encontra a última linha preenchida em "Saidas_Dom" coluna B
    lastRowSaidasDom = wsSaidasDom.Cells(wsSaidasDom.Rows.Count, "B").End(xlUp).Row

    ' Criar o dicionário
    Set dict = CreateObject("Scripting.Dictionary")

    ' Preencher o dicionário com as somas de "Saidas_Dom"
    For i = 5 To lastRowSaidasDom
        If (wsSaidasDom.Cells(i, "G").Value = 0 Or wsSaidasDom.Cells(i, "G").Value = 1) And _
           (wsSaidasDom.Cells(i, "P").Value <> 2 And wsSaidasDom.Cells(i, "P").Value <> 7) Then
            
            chave = wsSaidasDom.Cells(i, "B").Value
            If Not dict.Exists(chave) Then
                dict(chave) = 0
            End If
            dict(chave) = dict(chave) + 1
        End If
    Next i

    ' Preencher a coluna H em "Cont-Saídas" com as somas do dicionário
    For i = 3 To lastRowContSaidas
        valorC = wsContSaidas.Cells(i, "C").Value
        If dict.Exists(valorC) Then
            wsContSaidas.Cells(i, "H").Value = dict(valorC)
        Else
            wsContSaidas.Cells(i, "H").Value = 0
        End If
    Next i
    
    SomarValoresContSaidasCan
    
End Sub


Private Sub SomarValoresContSaidasCan()

    Dim wsContSaidas As Worksheet
    Dim wsSaidasDom As Worksheet
    Dim dictSaidas As Object
    Dim lastRowContSaidas As Long
    Dim lastRowSaidasDom As Long
    Dim i As Long
    Dim key As String
    Dim countValue As Long

    ' Desativa atualizações e alertas para melhor desempenho
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ' Define as planilhas
    Set wsContSaidas = ThisWorkbook.Sheets("Cont-Saidas")
    Set wsSaidasDom = ThisWorkbook.Sheets("Saidas_Dom")

    ' Cria um dicionário para armazenar as contagens
    Set dictSaidas = CreateObject("Scripting.Dictionary")
    
    ' Preenche o dicionário com os valores de Saidas_Dom
    lastRowSaidasDom = wsSaidasDom.Cells(wsSaidasDom.Rows.Count, "B").End(xlUp).Row
    For i = 5 To lastRowSaidasDom
        key = wsSaidasDom.Cells(i, "B").Value
        
        If (wsSaidasDom.Cells(i, "G").Value = "0" Or wsSaidasDom.Cells(i, "G").Value = "1") _
            And (wsSaidasDom.Cells(i, "P").Value = "2" Or wsSaidasDom.Cells(i, "P").Value = "7") Then

            If dictSaidas.Exists(key) Then
                dictSaidas(key) = dictSaidas(key) + 1
            Else
                dictSaidas.Add key, 1
            End If
        End If
    Next i

    ' Preenche a coluna H de "Cont-Saídas" com as contagens do dicionário
    lastRowContSaidas = wsContSaidas.Cells(wsContSaidas.Rows.Count, "C").End(xlUp).Row
    For i = 3 To lastRowContSaidas
        key = wsContSaidas.Cells(i, "C").Value
        If dictSaidas.Exists(key) Then
            wsContSaidas.Cells(i, "I").Value = dictSaidas(key)
        Else
            wsContSaidas.Cells(i, "I").Value = 0
        End If
    Next i

    SomarColunaNSaidasDom

End Sub


Private Sub SomarColunaNSaidasDom()

    Dim wsContSaidas As Worksheet
    Dim wsSaidasDom As Worksheet
    Dim dictSomas As Object
    Dim lastRowContSaidas As Long
    Dim lastRowSaidasDom As Long
    Dim i As Long
    Dim key As String
    Dim soma As Double


    ' Desativa atualizações e alertas para melhor desempenho
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False


    ' Define as planilhas
    Set wsContSaidas = ThisWorkbook.Sheets("Cont-Saidas")
    Set wsSaidasDom = ThisWorkbook.Sheets("Saidas_Dom")

    ' Cria um dicionário para armazenar as somas
    Set dictSomas = CreateObject("Scripting.Dictionary")
    
    ' Preenche o dicionário com as somas da coluna N de Saidas_Dom
    lastRowSaidasDom = wsSaidasDom.Cells(wsSaidasDom.Rows.Count, "B").End(xlUp).Row
    For i = 5 To lastRowSaidasDom
        key = wsSaidasDom.Cells(i, "B").Value
        
        If dictSomas.Exists(key) Then
            dictSomas(key) = Round(dictSomas(key) + wsSaidasDom.Cells(i, "N").Value, 2)
        Else
            dictSomas.Add key, Round(wsSaidasDom.Cells(i, "N").Value, 2)
        End If
    Next i

    ' Preenche a coluna K de "Cont-Saídas" com as somas do dicionário
    lastRowContSaidas = wsContSaidas.Cells(wsContSaidas.Rows.Count, "C").End(xlUp).Row
    For i = 3 To lastRowContSaidas
        key = wsContSaidas.Cells(i, "C").Value
        If dictSomas.Exists(key) Then
            wsContSaidas.Cells(i, "K").Value = Round(dictSomas(key), 2)
        Else
            wsContSaidas.Cells(i, "K").Value = 0
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
    Set ws = ThisWorkbook.Sheets("Cont-Saidas")

    ' Encontra a última linha preenchida na coluna J
    lastRow = ws.Cells(ws.Rows.Count, "J").End(xlUp).Row

    ' Loop através das linhas a partir da linha 3
    For i = 3 To lastRow
        ' Realiza a subtração de J - K e armazena o resultado na coluna L
        ws.Cells(i, "L").Value = ws.Cells(i, "J").Value - ws.Cells(i, "K").Value
    Next i




    PreencherDatasContSaidas

End Sub


Private Sub PreencherDatasContSaidas()
    Dim wsSIEG As Worksheet
    Dim wsContSaidas As Worksheet
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
    Set wsContSaidas = ThisWorkbook.Sheets("Cont-Saidas")
    
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
    
    ' Preencher a coluna D em "Cont-Saidas" com a menor data a partir de D3
    wsContSaidas.Range("D3:D" & wsContSaidas.Cells(wsContSaidas.Rows.Count, "C").End(xlUp).Row).Value = minDate
    
    ' Preencher a coluna E em "Cont-Saidas" com a maior data a partir de E3
    wsContSaidas.Range("E3:E" & wsContSaidas.Cells(wsContSaidas.Rows.Count, "C").End(xlUp).Row).Value = maxDate

    ' Ajustar a largura das colunas para melhor visualização (opcional)
    ' wsContSaidas.Columns("A:L").AutoFit

    ' Chamar a função CriarCompSaidas (se necessário)
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

    ' Definir a planilha "Cont-Saidas"
    Set ws = ThisWorkbook.Sheets("Cont-Saidas")

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

    CriarCompSaidas

End Sub




'Aqui começa notas faltantes

Private Sub CriarCompSaidas()

    Dim wsNFe As Worksheet
    Dim wsCompSaidas As Worksheet
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
    
    ' Criar uma nova aba chamada "Comp-Saidas"
    On Error Resume Next
    Set wsCompSaidas = ThisWorkbook.Sheets("Comp-Saidas")
    On Error GoTo 0
    
    If wsCompSaidas Is Nothing Then
        Set wsCompSaidas = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsCompSaidas.Name = "Comp-Saidas"
    End If

    ' Escrever os cabeçalhos na aba "Comp-Saidas"
    With wsCompSaidas
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
    
    ' Criar um dicionário para armazenar as saídas
    Set dictSaidas = CreateObject("Scripting.Dictionary")
    Set dictCTe = CreateObject("Scripting.Dictionary")
    Set dictCTe2 = CreateObject("Scripting.Dictionary")
    Set dictSaidas = CreateObject("Scripting.Dictionary") ' Inicialização do dicionário dictSaidas
    
    
    
    ' Encontrar a última linha preenchida na aba "NFe-NFCe_Sieg"
    lastRowNFe = wsNFe.Cells(wsNFe.Rows.Count, "A").End(xlUp).Row
    lastRowCTe = wsCTe.Cells(wsCTe.Rows.Count, "A").End(xlUp).Row
    
    ' Preencher o dicionário com os valores da planilha "NFe-NFCe_Sieg"
    j = 2 ' Iniciar na linha 2 da aba "Comp-Saidas"
    
    ' Processar as linhas onde "AD" = "Sai"
    For i = 2 To lastRowNFe
        If wsNFe.Cells(i, "AD").Value = "Sai" Then
            key = CStr(wsNFe.Cells(i, "G").Value)
            If Not dictSaidas.Exists(key) Then
                dictSaidas.Add key, Array(wsNFe.Cells(i, "K").Value, wsNFe.Cells(i, "A").Value, wsNFe.Cells(i, "AB").Value, wsNFe.Cells(i, "J").Value)
            End If
            
            ' Copiar os valores para a aba "Comp-Saidas"
            With wsCompSaidas
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
            With wsCompSaidas
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
    For i = 2 To lastRowCTe
        If IsNumeric(wsCTe.Cells(i, "AA").Value) And wsCTe.Cells(i, "AA").Value <> "" Then
            key = CStr(wsCTe.Cells(i, "AA").Value)
            If Not dictCTe.Exists(key) Then
                dictCTe.Add key, Array(wsCTe.Cells(i, "N").Value, wsCTe.Cells(i, "U").Value, wsCTe.Cells(i, "AA").Value)
            End If

            ' Copiar os valores para a aba "Comp-Entradas"
            With wsCompSaidas
                .Cells(j, 3).Value = wsCTe.Cells(i, "N").Value ' CNPJ
                .Cells(j, 4).Value = wsCTe.Cells(i, "D").Value ' Data
                .Cells(j, 5).Value = wsCTe.Cells(i, "U").Value ' Nota
                .Cells(j, 6).Value = 38
                .Cells(j, 7).Value = wsCTe.Cells(i, "BE").Value ' Nota
                .Cells(j, 8).Value = wsCTe.Cells(i, "AA").Value ' Valor Sieg
            End With
            j = j + 1
        End If
    Next i




    ' Ativar a aba "Comp-Saidas"
    wsCompSaidas.Activate
    
    PreencherCompSaidasComValoresFaltantes

End Sub




Private Sub PreencherCompSaidasComValoresFaltantes()
    Dim wsCompSaidas As Worksheet
    Dim wsSaidasDom As Worksheet
    Dim dictCompSaidas As Object
    Dim dictSaidasDom As Object
    Dim lastRowComp As Long, lastRowSaidasDom As Long
    Dim i As Long, lastRowNew As Long
    Dim BValue As String, EValue As String, DValue As String, FValue As String, IValue As String, TValue As String
    Dim key As String
    
    ' Desativa atualizações e alertas para melhor desempenho
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    ' Definir as planilhas
    Set wsCompSaidas = ThisWorkbook.Sheets("Comp-Saidas")
    Set wsSaidasDom = ThisWorkbook.Sheets("Saidas_Dom")
    
    ' Criar dicionário para armazenar as combinações de "Comp-Saidas"
    Set dictCompSaidas = CreateObject("Scripting.Dictionary")
    
    ' Criar dicionário para armazenar as combinações únicas de "Saidas_Dom"
    Set dictSaidasDom = CreateObject("Scripting.Dictionary")
    
    ' Encontrar a última linha preenchida em "Comp-Saidas"
    lastRowComp = wsCompSaidas.Cells(wsCompSaidas.Rows.Count, "C").End(xlUp).Row
    
    ' Preencher o dicionário com combinações de "C" e "E" em "Comp-Saidas"
    For i = 2 To lastRowComp
        BValue = CStr(wsCompSaidas.Cells(i, "C").Value)
        EValue = CStr(wsCompSaidas.Cells(i, "E").Value)
        DValue = CStr(wsCompSaidas.Cells(i, "D").Value)
        FValue = CStr(wsCompSaidas.Cells(i, "F").Value)
        key = BValue & "|" & EValue & "|" & DValue & "|" & FValue
        
        If Not dictCompSaidas.Exists(key) Then
            dictCompSaidas.Add key, True
        End If
    Next i
    
    ' Encontrar a última linha preenchida em "Saidas_Dom"
    lastRowSaidasDom = wsSaidasDom.Cells(wsSaidasDom.Rows.Count, "B").End(xlUp).Row
    
    ' Preencher "Comp-Saidas" com valores de "Saidas_Dom" onde a combinação não é encontrada
    For i = 5 To lastRowSaidasDom
        BValue = CStr(wsSaidasDom.Cells(i, "B").Value)
        EValue = CStr(wsSaidasDom.Cells(i, "E").Value)
        IValue = CStr(wsSaidasDom.Cells(i, "I").Value)
        TValue = CStr(wsSaidasDom.Cells(i, "T").Value)
        key = BValue & "|" & EValue & "|" & IValue & "|" & TValue
        
        ' Verificar se a combinação já foi adicionada ao dicionário de "Saidas_Dom"
        If Not dictSaidasDom.Exists(key) Then
            dictSaidasDom.Add key, True
            
            ' Verificar se a combinação não existe em "Comp-Saidas"
            If Not dictCompSaidas.Exists(key) Then
                lastRowNew = wsCompSaidas.Cells(wsCompSaidas.Rows.Count, "C").End(xlUp).Row + 1
                wsCompSaidas.Cells(lastRowNew, "C").Value = wsSaidasDom.Cells(i, "B").Value
                wsCompSaidas.Cells(lastRowNew, "D").Value = wsSaidasDom.Cells(i, "I").Value
                wsCompSaidas.Cells(lastRowNew, "E").Value = wsSaidasDom.Cells(i, "E").Value
                wsCompSaidas.Cells(lastRowNew, "F").Value = wsSaidasDom.Cells(i, "T").Value
                wsCompSaidas.Cells(lastRowNew, "G").Value = "Dominio"
                wsCompSaidas.Cells(lastRowNew, "H").Value = "N"
            End If
        End If
    Next i

    FiltrarCompSaidas

End Sub



Private Sub FiltrarCompSaidas()

    Dim wsCompSaidas As Worksheet
    Dim wsContSaidas As Worksheet
    Dim dictContSaidas As Object
    Dim lastRowComp As Long, lastRowCont As Long
    Dim i As Long, key As String

    ' Desativa atualizações e alertas para melhor desempenho
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ' Definir as planilhas
    Set wsCompSaidas = ThisWorkbook.Sheets("Comp-Saidas")
    Set wsContSaidas = ThisWorkbook.Sheets("Cont-Saidas")
    
    ' Criar dicionário para armazenar os valores da coluna C de "Cont-Saidas"
    Set dictContSaidas = CreateObject("Scripting.Dictionary")
    
    ' Encontrar as últimas linhas preenchidas em ambas as planilhas
    lastRowComp = wsCompSaidas.Cells(wsCompSaidas.Rows.Count, "C").End(xlUp).Row
    lastRowCont = wsContSaidas.Cells(wsContSaidas.Rows.Count, "C").End(xlUp).Row
    
    ' Preencher o dicionário com os valores da coluna C de "Cont-Saidas"
    For i = 3 To lastRowCont ' Começa de C3 conforme solicitado
        key = CStr(wsContSaidas.Cells(i, "C").Value)
        If Not dictContSaidas.Exists(key) Then
            dictContSaidas.Add key, True
        End If
    Next i
    
    ' Verificar e apagar linhas da aba "Comp-Saidas" cujos valores de C não estão em "Cont-Saidas"
    For i = lastRowComp To 2 Step -1 ' Percorrer de baixo para cima para evitar problemas ao excluir linhas
        key = CStr(wsCompSaidas.Cells(i, "C").Value)
        If Not dictContSaidas.Exists(key) Then
            wsCompSaidas.Rows(i).Delete
        End If
    Next i

    'RemoverLinhasDuplicadas
    PreencherCompSaidasComValores

End Sub





Private Sub PreencherCompSaidasComValores()

    Dim wsCompSaidas As Worksheet
    Dim wsSaidasDom As Worksheet
    Dim dictSaidasDom As Object
    Dim lastRowComp As Long, lastRowSaidasDom As Long
    Dim i As Long, key As String

    ' Desativa atualizações e alertas para melhor desempenho
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ' Definir as planilhas
    Set wsCompSaidas = ThisWorkbook.Sheets("Comp-Saidas")
    Set wsSaidasDom = ThisWorkbook.Sheets("Saidas_Dom")
    
    ' Criar dicionário para armazenar as combinações de "Saidas_Dom"
    Set dictSaidasDom = CreateObject("Scripting.Dictionary")
    
    ' Encontrar as últimas linhas preenchidas em ambas as planilhas
    lastRowComp = wsCompSaidas.Cells(wsCompSaidas.Rows.Count, "C").End(xlUp).Row
    lastRowSaidasDom = wsSaidasDom.Cells(wsSaidasDom.Rows.Count, "B").End(xlUp).Row
    
    ' Preencher o dicionário com as combinações de "B" e "E" em "Saidas_Dom"
    For i = 5 To lastRowSaidasDom ' Começa da linha 5 conforme solicitado
        key = CStr(wsSaidasDom.Cells(i, "B").Value) & "|" & CStr(wsSaidasDom.Cells(i, "E").Value) & "|" & CStr(wsSaidasDom.Cells(i, "I").Value) & "|" & CStr(wsSaidasDom.Cells(i, "T").Value)
        If dictSaidasDom.Exists(key) Then
            ' Se a chave já existir, soma o valor de "N" ao valor existente
            dictSaidasDom(key) = dictSaidasDom(key) + wsSaidasDom.Cells(i, "N").Value
        Else
            dictSaidasDom.Add key, wsSaidasDom.Cells(i, "N").Value
        End If
    Next i
    
    ' Preencher a coluna H de "Comp-Saidas" com os valores somados de "N" de "Saidas_Dom" ou "N" se não encontrado
    For i = 2 To lastRowComp ' Começa da linha 2 conforme solicitado
        key = CStr(wsCompSaidas.Cells(i, "C").Value) & "|" & CStr(wsCompSaidas.Cells(i, "E").Value) & "|" & CStr(wsCompSaidas.Cells(i, "D").Value) & "|" & CStr(wsCompSaidas.Cells(i, "F").Value)
        If dictSaidasDom.Exists(key) Then
            wsCompSaidas.Cells(i, "I").Value = dictSaidasDom(key)
        Else
            wsCompSaidas.Cells(i, "I").Value = "N"
        End If
    Next i

    ' Chamar a função para apagar linhas se necessário
    ApagarLinhasCompSaidas

End Sub








Private Sub ApagarLinhasCompSaidas()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim valorG As Double, valorH As Double
    Dim statusF As String

    ' Desativa atualizações e alertas para melhor desempenho
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False



    ' Defina a planilha "Comp-Saidas"
    Set ws = ThisWorkbook.Sheets("Comp-Saidas")
    
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
    
    
    PreencherColunaI
    
End Sub


Private Sub PreencherColunaI()
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

    ' Defina a planilha "Comp-Saidas"
    Set ws = ThisWorkbook.Sheets("Comp-Saidas")
    
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
    
    PreencherColunasAEBCompSaidas
    
End Sub




Private Sub PreencherColunasAEBCompSaidas()
    Dim wsCompSaidas As Worksheet
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
    Set wsCompSaidas = ThisWorkbook.Sheets("Comp-Saidas")
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
    
    ' Encontrar a última linha preenchida em "Comp-Saidas"
    lastRowComp = wsCompSaidas.Cells(wsCompSaidas.Rows.Count, "C").End(xlUp).Row
    
    ' Percorrer "Comp-Saidas" e preencher colunas A e B com base no dicionário
    For i = 2 To lastRowComp
        valorC = CStr(wsCompSaidas.Cells(i, "C").Value)
        
        ' Verificar se o valor de C existe no dicionário
        If dictEmpresasDom.Exists(valorC) Then
            wsCompSaidas.Cells(i, "A").Value = dictEmpresasDom(valorC)(0) ' Preenche a coluna A
            wsCompSaidas.Cells(i, "B").Value = dictEmpresasDom(valorC)(1) ' Preenche a coluna B
        End If
    Next i
    
    ' Formatando a coluna D como Data Abreviada
    wsCompSaidas.Columns("D").NumberFormat = "dd/mm/yy"
    
    
    ' Ajustar a largura das colunas para melhor visualização
    wsCompSaidas.Columns("A:J").AutoFit
    
   ' Ordenar os dados pela coluna A de maneira crescente, ignorando o cabeçalho
    wsCompSaidas.Sort.SortFields.Clear
    wsCompSaidas.Sort.SortFields.Add key:=wsCompSaidas.Range("A2:A" & lastRowComp), Order:=xlAscending
    With wsCompSaidas.Sort
        .SetRange wsCompSaidas.Range("A1:J" & lastRowComp) ' Ajuste o intervalo conforme necessário
        .Header = xlYes ' Considerar o cabeçalho
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    
    
    CEntradas.CEntrada
    
End Sub

