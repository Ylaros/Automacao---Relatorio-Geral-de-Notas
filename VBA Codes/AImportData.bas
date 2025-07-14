Attribute VB_Name = "AImportData"
Sub AIniciar()
Attribute AIniciar.VB_ProcData.VB_Invoke_Func = "m\n14"

    Dim ws As Worksheet

    ' Desativa alertas para confirmação de exclusão
    Application.DisplayAlerts = False

    ' Loop através de todas as abas, exceto "Planilha1"
    For Each ws In ThisWorkbook.Sheets
        If ws.Name <> "Plan1" Then
            ws.Delete
        End If
    Next ws

    ' Reativa alertas
    Application.DisplayAlerts = True
    
    ImportarDadosDominio

End Sub


Private Sub ImportarDadosDominio()

Dim filePath As String
Dim fileName As String
Dim infoA3 As String

'Automação Excel para Cópia e Colagem das planilhas domínio
'


    ' Desativa atualizações e alertas para melhor desempenho
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ' Verifica e cria a aba Saidas_Dom se não existir
    On Error Resume Next
    Sheets("Saidas_Dom").Activate

    If Err.Number <> 0 Then
        Sheets.Add(After:=Sheets(Sheets.Count)).Name = "Saidas_Dom"
        Err.Clear
    End If

    ' Passo 2 - Abre e copia o conteúdo do primeiro arquivo (apenas valores)
    Workbooks.Open "Z:\18 - T.I\Relatório Geral de Notas\Resumido - Relatório de Saídas.xls"
    Sheets(1).UsedRange.Copy
    ThisWorkbook.Sheets("Saidas_Dom").Cells(ThisWorkbook.Sheets("Saidas_Dom").Rows.Count, 1).End(xlUp).Offset(1, 0).PasteSpecial xlPasteValues
    Application.CutCopyMode = False
    Workbooks("Resumido - Relatório de Saídas.xls").Close False

    ' Verifica e cria a aba Entradas_Dom se não existir
    On Error Resume Next
    Sheets("Entradas_Dom").Activate

    If Err.Number <> 0 Then
        Sheets.Add(After:=Sheets(Sheets.Count)).Name = "Entradas_Dom"
        Err.Clear
    End If

    ' Passo 4 - Repete o processo para o segundo arquivo (apenas valores)
    Workbooks.Open "Z:\18 - T.I\Relatório Geral de Notas\Resumido - Relatório de Entradas.xls"
    Sheets(1).UsedRange.Copy
    ThisWorkbook.Sheets("Entradas_Dom").Cells(ThisWorkbook.Sheets("Entradas_Dom").Rows.Count, 1).End(xlUp).Offset(1, 0).PasteSpecial xlPasteValues
    Application.CutCopyMode = False
    Workbooks("Resumido - Relatório de Entradas.xls").Close False

    ' Verifica e cria a aba CFs_Dom se não existir
    On Error Resume Next
    Sheets("CFs_Dom").Activate

    If Err.Number <> 0 Then
        Sheets.Add(After:=Sheets(Sheets.Count)).Name = "CFs_Dom"
        Err.Clear
    End If
    
    ' Passo 6 - Repete o processo para o terceiro arquivo (apenas valores)
    Workbooks.Open "Z:\18 - T.I\Relatório Geral de Notas\Resumido - Relatório de Cupons Fiscais.xls"
    Sheets(1).UsedRange.Copy
    ThisWorkbook.Sheets("CFs_Dom").Cells(ThisWorkbook.Sheets("CFs_Dom").Rows.Count, 1).End(xlUp).Offset(1, 0).PasteSpecial xlPasteValues
    Application.CutCopyMode = False
    Workbooks("Resumido - Relatório de Cupons Fiscais.xls").Close False
    Rows("2:2").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove

    ' Verifica e cria a aba Empresas se não existir
    On Error Resume Next
    Sheets("Empresas_Dom").Activate

    If Err.Number <> 0 Then
        Sheets.Add(After:=Sheets(Sheets.Count)).Name = "Empresas_Dom"
        Err.Clear
    End If
    
    ' Passo 8 - Repete o processo para o terceiro arquivo (apenas valores)
    Workbooks.Open "Z:\18 - T.I\Relatório Geral de Notas\Empresas.xls"
    Sheets(1).UsedRange.Copy
    ThisWorkbook.Sheets("Empresas_Dom").Cells(ThisWorkbook.Sheets("CFs_Dom").Rows.Count, 1).End(xlUp).Offset(1, 0).PasteSpecial xlPasteValues
    Application.CutCopyMode = False
    Workbooks("Empresas.xls").Close False
    
    
'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx




    ' Verifica e cria a aba se não existir
    On Error Resume Next
    Sheets("NFe-NFCe_Sieg").Activate

    If Err.Number <> 0 Then
        Sheets.Add(After:=Sheets(Sheets.Count)).Name = "NFe-NFCe_Sieg"
        Err.Clear
    End If
    
    ' Define o caminho da pasta onde o arquivo está localizado
    filePath = "Z:\18 - T.I\Relatório Geral de Notas\"
    
    ' Usa o método Dir para encontrar o primeiro arquivo que comece com "RelatorioNFe"
    fileName = Dir(filePath & "RelatorioNFe*")
    
    
    ' Passo 6 - Repete o processo para o terceiro arquivo (apenas valores)
    Workbooks.Open filePath & fileName
    Sheets(1).UsedRange.Copy
    ThisWorkbook.Sheets("NFe-NFCe_Sieg").Cells(ThisWorkbook.Sheets("NFe-NFCe_Sieg").Rows.Count, 1).End(xlUp).Offset(1, 0).PasteSpecial xlPasteValues
    Application.CutCopyMode = False
    Workbooks(fileName).Close False
    Rows("3:3").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove





    ' Verifica e cria a aba se não existir
    On Error Resume Next
    Sheets("CFe_Sieg").Activate

    If Err.Number <> 0 Then
        Sheets.Add(After:=Sheets(Sheets.Count)).Name = "CFe_Sieg"
        Err.Clear
    End If
    
    ' Define o caminho da pasta onde o arquivo está localizado
    'filePath = "Z:\18 - T.I\Automações\Automações de conferência\Conferência de Saídas\"
    
    ' Usa o método Dir para encontrar o primeiro arquivo que comece com "RelatorioNFe"
    fileName = Dir(filePath & "RelatorioCFe*")
    
    
    ' Passo 6 - Repete o processo para o terceiro arquivo (apenas valores)
    Workbooks.Open filePath & fileName
    Sheets(1).UsedRange.Copy
    ThisWorkbook.Sheets("CFe_Sieg").Cells(ThisWorkbook.Sheets("CFe_Sieg").Rows.Count, 1).End(xlUp).Offset(1, 0).PasteSpecial xlPasteValues
    Application.CutCopyMode = False
    Workbooks(fileName).Close False
    Rows("3:3").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove





     ' Verifica e cria a aba se não existir
    On Error Resume Next
    Sheets("CTe_Sieg").Activate

    If Err.Number <> 0 Then
        Sheets.Add(After:=Sheets(Sheets.Count)).Name = "CTe_Sieg"
        Err.Clear
    End If
    
    ' Define o caminho da pasta onde o arquivo está localizado
    'filePath = "Z:\18 - T.I\Automações\Automações de conferência\Conferência de Saídas\"
    
    ' Usa o método Dir para encontrar o primeiro arquivo que comece com "RelatorioNFe"
    fileName = Dir(filePath & "RelatorioCTe*")
    
    
    ' Passo 6 - Repete o processo para o terceiro arquivo (apenas valores)
    Workbooks.Open filePath & fileName
    Sheets(1).UsedRange.Copy
    ThisWorkbook.Sheets("CTe_Sieg").Cells(ThisWorkbook.Sheets("CTe_Sieg").Rows.Count, 1).End(xlUp).Offset(1, 0).PasteSpecial xlPasteValues
    Application.CutCopyMode = False
    Workbooks(fileName).Close False
    Rows("2:2").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove

    
    
    
  
    fileName = Dir(filePath & "Relatorio Xml Cofre SIEG*.xlsx")
        
    'Verifica e cria a aba SIEG se não existir
    On Error Resume Next
    Sheets("SIEG").Activate
    If Err.Number <> 0 Then
        Sheets.Add(After:=Sheets(Sheets.Count)).Name = "SIEG"
        Err.Clear
    End If
    

    
    Do While fileName <> ""
        ' Abre cada arquivo com prefixo "Relatorio Xml Cofre SIEG" e copia somente valores
        Set wb = Workbooks.Open(filePath & fileName)
        wb.Sheets(1).UsedRange.Copy
        ThisWorkbook.Sheets("SIEG").Cells(ThisWorkbook.Sheets("SIEG").Rows.Count, 1).End(xlUp).Offset(1, 0).PasteSpecial xlPasteValues
        Application.CutCopyMode = False
        wb.Close False
        fileName = Dir
    Loop
    
    
    
    
    fileName = Dir(filePath & "CTE*.xlsx")
        
    'Verifica e cria a aba SIEG se não existir
    On Error Resume Next
    Sheets("SIEG2").Activate
    If Err.Number <> 0 Then
        Sheets.Add(After:=Sheets(Sheets.Count)).Name = "SIEG2"
        Err.Clear
    End If
    

    
    Do While fileName <> ""
        ' Abre cada arquivo com prefixo "Relatorio Xml Cofre SIEG" e copia somente valores
        Set wb = Workbooks.Open(filePath & fileName)
        wb.Sheets(1).UsedRange.Copy
        ThisWorkbook.Sheets("SIEG2").Cells(ThisWorkbook.Sheets("SIEG2").Rows.Count, 1).End(xlUp).Offset(1, 0).PasteSpecial xlPasteValues
        Application.CutCopyMode = False
        wb.Close False
        fileName = Dir
    Loop
    
    

    ' Reativa atualizações e alertas
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

    PreencherColunaR

End Sub


Private Sub PreencherColunaR()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim currentValue As Long
    
    
    ' Desativa atualizações e alertas para melhor desempenho
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ' Definir a planilha e a última linha preenchida na coluna A
    Set ws = ThisWorkbook.Sheets("SIEG")
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    ' Inicializar a variável
    currentValue = 0
    
    ' Percorrer a coluna A
    For i = 1 To lastRow
        Select Case ws.Cells(i, 1).Value
            Case "Num NFCe"
                currentValue = 41
            Case "Num NFe"
                currentValue = 36
            Case "Num CFe"
                currentValue = 99
        End Select
        
        ' Preencher a coluna R com o valor atual
        If currentValue <> 0 Then
            ws.Cells(i, 18).Value = currentValue
        End If
    Next i
    
    PadronizarColunas
    
End Sub





Private Sub PadronizarColunas()


    ' Desativa atualizações e alertas para melhor desempenho
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ' Padroniza as colunas conforme especificado
    Call RemoverCaracteres("Empresas_Dom", "I")
    Call RemoverCaracteres("NFe-NFCe_Sieg", "D")
    Call RemoverCaracteres("NFe-NFCe_Sieg", "G")
    Call RemoverCaracteres("CFe_Sieg", "D")
    Call RemoverCaracteres("CFe_Sieg", "F")
    Call RemoverCaracteres("CTe_Sieg", "G")
    Call RemoverCaracteres("CTe_Sieg", "N")
    Call RemoverCaracteres("CTe_Sieg", "P")
    Call RemoverCaracteres("SIEG", "D")
    Call RemoverCaracteres("SIEG", "G")
    
    ConvertGeneral
    
End Sub

Sub RemoverCaracteres(sheetName As String, col As String)
    Dim ws As Worksheet
    Dim rng As Range

    ' Define a aba a ser processada
    Set ws = ThisWorkbook.Sheets(sheetName)
    
    ' Define o intervalo de células na coluna especificada
    Set rng = ws.Columns(col)
    
    ' Aplicar Replace nos caracteres especificados em todo o intervalo
    rng.Replace What:=".", Replacement:="", LookAt:=xlPart, MatchCase:=False
    rng.Replace What:="/", Replacement:="", LookAt:=xlPart, MatchCase:=False
    rng.Replace What:="-", Replacement:="", LookAt:=xlPart, MatchCase:=False
    rng.Replace What:=" ", Replacement:="", LookAt:=xlPart, MatchCase:=False

End Sub




Private Sub ConvertGeneral()

    ' Desativa atualizações e alertas para melhor desempenho
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    ' Converter as colunas especificadas nas abas para o formato Geral
    Call ForcarEdicao("NFe-NFCe_Sieg", "A")
    Call ForcarEdicao("CFe_Sieg", "A")
    Call ForcarEdicao("CTe_Sieg", "U")
    Call ForcarEdicao("CTe_Sieg", "AA")
    Call ForcarEdicao("Empresas_Dom", "I")
    
    SubstituirPontosPorVirgulas
    
End Sub

Sub ForcarEdicao(sheetName As String, col As String)
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim cell As Range

    ' Define a aba a ser processada
    Set ws = ThisWorkbook.Sheets(sheetName)

    ' Encontra a última linha preenchida na coluna especificada
    lastRow = ws.Cells(ws.Rows.Count, col).End(xlUp).Row
    
    ' Loop através de cada célula na coluna especificada
    For Each cell In ws.Range(col & "2:" & col & lastRow)
        If Not IsEmpty(cell.Value) Then
            ' Força a edição do conteúdo da célula e define o formato para Geral
            cell.Value = cell.Value ' Reforça a edição do conteúdo
            cell.NumberFormat = "General"
        End If
    Next cell
End Sub

Private Sub SubstituirPontosPorVirgulas()

    ' Desativa atualizações e alertas para melhor desempenho
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ' Seleciona a planilha "CFe_Sieg"
    With ThisWorkbook.Sheets("CFe_Sieg")
        ' Seleciona a coluna I e realiza a substituição
        .Columns("I").Replace What:=".", Replacement:=",", LookAt:=xlPart, _
                              SearchOrder:=xlByRows, MatchCase:=False
    End With



    ConvertCFeNumeros

End Sub



Private Sub ConvertCFeNumeros()


    ' Desativa atualizações e alertas para melhor desempenho
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ' Converter os valores de texto para números na coluna D de "CFe_Sieg"
    Call ForcarNumeros("CFe_Sieg", "A", "D")
    Call ForcarNumeros("CFe_Sieg", "A", "I")

    
   ConvertNFeNumeros
    
End Sub

Sub ForcarNumeros(sheetName As String, colCheck As String, colConvert As String)
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim cell As Range

    ' Define a aba a ser processada
    Set ws = ThisWorkbook.Sheets(sheetName)

    ' Encontra a última linha preenchida na coluna de verificação (colCheck)
    lastRow = ws.Cells(ws.Rows.Count, colCheck).End(xlUp).Row
    
    ' Loop através de cada célula na coluna de verificação (colCheck)
    For Each cell In ws.Range(colCheck & "6:" & colCheck & lastRow)
        If Not IsEmpty(cell.Value) Then
            ' Converte o valor correspondente na coluna a ser convertida (colConvert) para número
            With ws.Cells(cell.Row, colConvert)
                If IsNumeric(.Value) And Not IsEmpty(.Value) Then
                    .Value = CDbl(.Value)
                End If
            End With
        End If
    Next cell
End Sub




Private Sub ConvertNFeNumeros()

    ' Desativa atualizações e alertas para melhor desempenho
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ' Converter os valores de texto para números na coluna J de "NFe-NFCe_Sieg"
    Call ForcarNumeros3("NFe-NFCe_Sieg", "A", "J")

    SubstituirHifenPorBarra

End Sub

Sub ForcarNumeros3(sheetName As String, colCheck As String, colConvert As String)
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim cell As Range

    ' Define a aba a ser processada
    Set ws = ThisWorkbook.Sheets(sheetName)

    ' Encontra a última linha preenchida na coluna de verificação (colCheck)
    lastRow = ws.Cells(ws.Rows.Count, colCheck).End(xlUp).Row
    
    ' Loop através de cada célula na coluna de verificação (colCheck)
    For Each cell In ws.Range(colCheck & "6:" & colCheck & lastRow)
        If Not IsEmpty(cell.Value) Then
            ' Converte o valor correspondente na coluna a ser convertida (colConvert) para número
            With ws.Cells(cell.Row, colConvert)
                If IsNumeric(.Value) And Not IsEmpty(.Value) Then
                    .Value = CDbl(.Value)
                End If
            End With
        End If
    Next cell
End Sub




Private Sub SubstituirHifenPorBarra()


    ' Desativa atualizações e alertas para melhor desempenho
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False


    Dim ws As Worksheet
    Dim ws2 As Worksheet
    Dim ws3 As Worksheet
    Dim wsSIEG As Worksheet
    Dim rng As Range

    ' Define a aba "NFe-NFCe_Sieg"
    Set ws = ThisWorkbook.Sheets("NFe-NFCe_Sieg")
    Set ws2 = ThisWorkbook.Sheets("CFe_Sieg")
    Set ws3 = ThisWorkbook.Sheets("CTe_Sieg")

    ' Define o intervalo como toda a coluna
    Set rng = ws.Columns("K")

    ' Substitui "-" por "/" em toda a coluna K
    rng.Replace What:="-", Replacement:="/", LookAt:=xlPart, MatchCase:=False
    
    
    ' Define o intervalo como toda a coluna
    Set rng = ws2.Columns("C")

    ' Substitui "-" por "/" em toda a coluna K
    rng.Replace What:="-", Replacement:="/", LookAt:=xlPart, MatchCase:=False
    
        
     ' Define o intervalo como toda a coluna
    Set rng = ws3.Columns("D")

    ' Substitui "-" por "/" em toda a coluna K
    rng.Replace What:="-", Replacement:="/", LookAt:=xlPart, MatchCase:=False
        

    RemoverLinhasNaoNumericas

End Sub


Private Sub RemoverLinhasNaoNumericas()


    ' Desativa atualizações e alertas para melhor desempenho
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False


    Dim ws As Worksheet
    Dim i As Long
    Dim lastRow As Long

    ' Define a aba "SIEG"
    Set ws = ThisWorkbook.Sheets("SIEG")

    ' Encontra a última linha preenchida na coluna A
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    ' Percorre a coluna A a partir da linha 5 até a última linha
    For i = lastRow To 5 Step -1
        ' Verifica se o valor na célula não é numérico
        If Not IsNumeric(ws.Cells(i, "A").Value) Then
            ' Apaga a linha se o valor não for numérico
            ws.Rows(i).Delete
        End If
    Next i
    
    ConvertS
    
End Sub





 
 Private Sub ConvertS()

    ' Desativa atualizações e alertas para melhor desempenho
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False


    ' Converter as colunas especificadas nas abas
    Call ForcarEdicao2("SIEG", "C")
    Call ForcarEdicao2("SIEG", "J")
    Call ForcarEdicao2("Entradas_Dom", "I")
    Call ForcarEdicao2("Saidas_Dom", "I")
    Call ForcarEdicao5("CFs_Dom", "B")

    
    VerificarValoresEPreencherAA
    
End Sub

Sub ForcarEdicao2(sheetName As String, col As String)
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim cell As Range

    ' Define a aba a ser processada
    Set ws = ThisWorkbook.Sheets(sheetName)

    ' Encontra a última linha preenchida na coluna especificada
    lastRow = ws.Cells(ws.Rows.Count, col).End(xlUp).Row
    
    ' Loop através de cada célula na coluna especificada
    For Each cell In ws.Range(col & "5:" & col & lastRow)
        If Not IsEmpty(cell.Value) Then
            ' Força a edição do conteúdo da célula e converte o valor para data
            cell.Value = CDate(cell.Value)
        End If
    Next cell
End Sub

Sub ForcarEdicao5(sheetName As String, col As String)
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim cell As Range

    ' Define a aba a ser processada
    Set ws = ThisWorkbook.Sheets(sheetName)

    ' Encontra a última linha preenchida na coluna especificada
    lastRow = ws.Cells(ws.Rows.Count, col).End(xlUp).Row
    
    ' Loop através de cada célula na coluna especificada
    For Each cell In ws.Range(col & "7:" & col & lastRow)
        If Not IsEmpty(cell.Value) Then
            ' Força a edição do conteúdo da célula e converte o valor para data
            cell.Value = CDate(cell.Value)
        End If
    Next cell
End Sub





'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx


Private Sub VerificarValoresEPreencherAA()

    ' Desativa atualizações e alertas para melhor desempenho
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False


    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim valor As String

    ' Define a aba "NFe-NFCe_Sieg"
    Set ws = ThisWorkbook.Sheets("NFe-NFCe_Sieg")

    ' Encontra a última linha com valor na coluna B
    lastRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row

    ' Loop da linha 5 até a última linha válida
    For i = 6 To lastRow
        ' Verifica se a célula em AA está vazia
        If ws.Cells(i, "AA").Value = "" Then
            ' Pega o valor duas linhas abaixo na coluna AA
            valor = ws.Cells(i + 2, "AA").Value

            ' Verifica se o valor começa com 5, 6 ou 7
            If Left(valor, 1) = "5" Or Left(valor, 1) = "6" Or Left(valor, 1) = "7" Then
                ws.Cells(i, "AA").Value = "Saída"
            Else
                ws.Cells(i, "AA").Value = "Entrada"
            End If
        End If
    Next i

    RemoverValoresNaoNumericos

End Sub


Private Sub RemoverValoresNaoNumericos()


    ' Desativa atualizações e alertas para melhor desempenho
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False


    Dim ws As Worksheet
    Dim i As Long
    Dim lastRowD As Long
    Dim lastRowG As Long

    ' Define a aba "SIEG"
    Set ws = ThisWorkbook.Sheets("SIEG")

    ' Encontra a última linha preenchida nas colunas D e G
    lastRowD = ws.Cells(ws.Rows.Count, "D").End(xlUp).Row
    lastRowG = ws.Cells(ws.Rows.Count, "G").End(xlUp).Row

    ' Percorre a coluna D a partir da linha 5 até a última linha
    For i = 5 To lastRowD
        ' Verifica se o valor na célula não é numérico
        If Not IsNumeric(ws.Cells(i, "D").Value) Then
            ' Apaga o valor se não for numérico
            ws.Cells(i, "D").ClearContents
        End If
    Next i

    ' Percorre a coluna G a partir da linha 5 até a última linha
    For i = 5 To lastRowG
        ' Verifica se o valor na célula não é numérico
        If Not IsNumeric(ws.Cells(i, "G").Value) Then
            ' Apaga o valor se não for numérico
            ws.Cells(i, "G").ClearContents
        End If
    Next i

    PreencherColunaT

End Sub




Private Sub PreencherColunaT()


    ' Desativa atualizações e alertas para melhor desempenho
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False


    Dim wsNFe As Worksheet
    Dim wsCTe As Worksheet
    Dim wsSIEG As Worksheet
    Dim wsSieg2 As Worksheet
    Dim lastRowNFe As Long, lastRowSIEG As Long, lastRowSieg2 As Long, lastRowCTe As Long
    Dim i As Long
    Dim key As String
    Dim dic As Object
    Dim dict As Object
    Dim dict2 As Object

    ' Define as abas
    Set wsNFe = ThisWorkbook.Sheets("NFe-NFCe_Sieg")
    Set wsCTe = ThisWorkbook.Sheets("CTe_Sieg")
    Set wsSIEG = ThisWorkbook.Sheets("SIEG")
    Set wsSieg2 = ThisWorkbook.Sheets("SIEG2")

    ' Encontra a última linha preenchida em ambas as abas
    lastRowNFe = wsNFe.Cells(wsNFe.Rows.Count, "A").End(xlUp).Row
    lastRowCTe = wsCTe.Cells(wsCTe.Rows.Count, "A").End(xlUp).Row
    lastRowSIEG = wsSIEG.Cells(wsSIEG.Rows.Count, "A").End(xlUp).Row
    lastRowSieg2 = wsSieg2.Cells(wsSieg2.Rows.Count, "A").End(xlUp).Row

    ' Cria um dicionário para armazenar as combinações da aba "SIEG"
    Set dic = CreateObject("Scripting.Dictionary")
    Set dict = CreateObject("Scripting.Dictionary")
    Set dict2 = CreateObject("Scripting.Dictionary")

    ' Preenche o dicionário com combinações únicas das colunas A, C, D, F da aba "SIEG"
    For i = 2 To lastRowSIEG
        key = wsSIEG.Cells(i, "A").Value & "|" & wsSIEG.Cells(i, "C").Value & "|" & wsSIEG.Cells(i, "G").Value & "|" & wsSIEG.Cells(i, "D").Value
        If Not dict.Exists(key) Then
            dict.Add key, wsSIEG.Cells(i, "O").Value
            dic.Add key, wsSIEG.Cells(i, "R").Value
        End If
    Next i
    

    ' Percorre as linhas da aba "NFe-NFCe_Sieg" e busca a combinação no dicionário
    For i = 2 To lastRowNFe
        key = wsNFe.Cells(i, "A").Value & "|" & wsNFe.Cells(i, "K").Value & "|" & wsNFe.Cells(i, "D").Value & "|" & wsNFe.Cells(i, "G").Value
        If dict.Exists(key) Then
            wsNFe.Cells(i, "AB").Value = dict(key)
            wsNFe.Cells(i, "AE").Value = dic(key)
        Else
            wsNFe.Cells(i, "AB").Value = "" ' Não preenche nada se não encontrar
        End If
    Next i
    
    
    ' Preenche o dicionário com combinações únicas das colunas A da aba "SIEG2"
    For i = 2 To lastRowSieg2
        key = wsSieg2.Cells(i, "A").Value
        If Not dict2.Exists(key) Then
            dict2.Add key, wsSieg2.Cells(i, "I").Value
        End If
    Next i
    
    ' Percorre as linhas da aba "CTe" e busca a combinação no dicionário
    For i = 2 To lastRowCTe
        key = wsCTe.Cells(i, "K").Value
        If dict2.Exists(key) Then
            wsCTe.Cells(i, "BE").Value = dict2(key)
            wsCTe.Cells(i, "BF").Value = 38
        Else
            wsCTe.Cells(i, "BE").Value = "" ' Não preenche nada se não encontrar
        End If
    Next i
    
    
    PreencherColunaN
    
    
End Sub


Private Sub PreencherColunaN()


    ' Desativa atualizações e alertas para melhor desempenho
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False


    Dim wsNFe As Worksheet
    Dim wsSIEG As Worksheet
    Dim lastRowNFe As Long, lastRowSIEG As Long
    Dim i As Long
    Dim key As String
    Dim dict As Object

    ' Define as abas
    Set wsNFe = ThisWorkbook.Sheets("CFe_Sieg")
    Set wsSIEG = ThisWorkbook.Sheets("SIEG")

    ' Encontra a última linha preenchida em ambas as abas
    lastRowNFe = wsNFe.Cells(wsNFe.Rows.Count, "A").End(xlUp).Row
    lastRowSIEG = wsSIEG.Cells(wsSIEG.Rows.Count, "A").End(xlUp).Row

    ' Cria um dicionário para armazenar as combinações da aba "SIEG"
    Set dict = CreateObject("Scripting.Dictionary")

    ' Preenche o dicionário com combinações únicas das colunas A, C, D, F da aba "SIEG"
    For i = 2 To lastRowSIEG
        key = wsSIEG.Cells(i, "A").Value & "|" & wsSIEG.Cells(i, "C").Value & "|" & wsSIEG.Cells(i, "D").Value
        If Not dict.Exists(key) Then
            dict.Add key, wsSIEG.Cells(i, "M").Value
        End If
    Next i

    ' Percorre as linhas da aba "NFe-NFCe_Sieg" e busca a combinação no dicionário
    For i = 2 To lastRowNFe
        key = wsNFe.Cells(i, "A").Value & "|" & wsNFe.Cells(i, "C").Value & "|" & wsNFe.Cells(i, "D").Value
        If dict.Exists(key) Then
            wsNFe.Cells(i, "N").Value = dict(key)
        Else
            wsNFe.Cells(i, "N").Value = "" ' Não preenche nada se não encontrar
        End If
    Next i
    
    PreencherColunaAC
    
    
End Sub



Private Sub PreencherColunaAC()

    ' Desativa atualizações e alertas para melhor desempenho
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False


    Dim ws As Worksheet
    Dim lastRowB As Long
    Dim i As Long

    ' Define a aba "NFe-NFCe_Sieg"
    Set ws = ThisWorkbook.Sheets("NFe-NFCe_Sieg")

    ' Encontra a última linha preenchida na coluna B
    lastRowB = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row

    ' Percorre as linhas até a última linha válida na coluna B
    For i = 5 To lastRowB
        ' Verifica se AB não é vazio
        If ws.Cells(i, "AB").Value <> "" Then
            ' Preenche AC com o valor de duas linhas abaixo de AA
            ws.Cells(i, "AC").Value = ws.Cells(i + 2, "AA").Value
        Else
            ' Deixa AC vazio se AB for vazio
            ws.Cells(i, "AC").Value = ""
        End If
    Next i
    
    ClassificarValoresAC

End Sub

Private Sub ClassificarValoresAC()

    ' Desativa atualizações e alertas para melhor desempenho
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False


    Dim ws As Worksheet
    Dim lastRowB As Long
    Dim i As Long
    Dim valorAC As String
    Dim devValores As Variant
    
    ' Lista de valores para "Dev"
    devValores = Array("1202", "1203", "1204", "1208", "1209", "1212", "1213", "1214", "1410", _
                       "1915", "1411", "1503", "1504", "1505", "1506", "1553", "1660", "1661", "1662", _
                       "1918", "1919", "2201", "2202", "2203", "2204", "2208", "2209", "2212", "2213", _
                       "2214", "2410", "2411", "2503", "2504", "2505", "2506", "2553", "2660", "2661", _
                       "2662", "2918", "2919", "3201", "3202", "3211", "3212", "3503", "3553", "1201", _
                       "5208", "5209", "5213", "5214", "5215", "5410", "2909", "1664", "2606", "2921", _
                       "5503", "5660", "5661", "5662", "5919", "1909", "2206", "1917", "1949", _
                       "6208", "6209", "6213", "6214", "6215", "6410", "1102", _
                       "6412", "6413", "6503", "6553", "6555", "6660", "6661", "6662", _
                       "6919", "7201", "7202", "7210", "7211", "7212", "7553", "7556", _
                       "7930")
                       
                       
'"6202", "5202", "5556", "6556", "5949", "5411", "6201", "5921", "6921", "5413", "5201", "1201", "6411", "5412", "5553", "5918", "6918", "5210", "6210", "5555"

    ' Define a aba "NFe-NFCe_Sieg"
    Set ws = ThisWorkbook.Sheets("NFe-NFCe_Sieg")

    ' Encontra a última linha preenchida na coluna B
    lastRowB = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row

    ' Percorre as linhas até a última linha válida na coluna B
    For i = 5 To lastRowB
        valorAC = ws.Cells(i, "AC").Value

        ' Verifica se AC está vazio
        If valorAC <> "" Then
            ' Verifica se o valor está na lista de "Dev"
            If Not IsError(Application.Match(valorAC, devValores, 0)) Then
                ws.Cells(i, "AD").Value = "Dev"
            ' Verifica se começa com 1, 2 ou 3 para "Ent"
            ElseIf Left(valorAC, 1) = "1" Or Left(valorAC, 1) = "2" Or Left(valorAC, 1) = "3" Then
                ws.Cells(i, "AD").Value = "Ent"
            ' Verifica se começa com 5, 6 ou 7 para "Sai"
            ElseIf Left(valorAC, 1) = "5" Or Left(valorAC, 1) = "6" Or Left(valorAC, 1) = "7" Then
                ws.Cells(i, "AD").Value = "Sai"
            End If
        End If
    Next i
    
    PreencherValores
    
End Sub


Private Sub PreencherValores()


    ' Desativa atualizações e alertas para melhor desempenho
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    Dim wsEmpresasDom As Worksheet
    Dim wsSaidasDom As Worksheet
    Dim wsEntradasDom As Worksheet
    Dim wsCFsDom As Worksheet
    Dim dictEmpresas As Object
    Dim lastRow As Long
    Dim i As Long
    Dim key As String

    ' Define as abas
    Set wsEmpresasDom = ThisWorkbook.Sheets("Empresas_Dom")
    Set wsSaidasDom = ThisWorkbook.Sheets("Saidas_Dom")
    Set wsEntradasDom = ThisWorkbook.Sheets("Entradas_Dom")
    Set wsCFsDom = ThisWorkbook.Sheets("CFs_Dom")

    ' Cria um dicionário para armazenar os valores da aba "Empresas_Dom"
    Set dictEmpresas = CreateObject("Scripting.Dictionary")
    
    ' Preenche o dicionário com os valores da coluna A e I de "Empresas_Dom"
    lastRow = wsEmpresasDom.Cells(wsEmpresasDom.Rows.Count, "A").End(xlUp).Row
    For i = 2 To lastRow
        key = wsEmpresasDom.Cells(i, "A").Value
        If Not dictEmpresas.Exists(key) Then
            dictEmpresas.Add key, wsEmpresasDom.Cells(i, "I").Value
        End If
    Next i

    ' Preenche a coluna B de "Saidas_Dom"
    lastRow = wsSaidasDom.Cells(wsSaidasDom.Rows.Count, "A").End(xlUp).Row
    For i = 5 To lastRow
        key = wsSaidasDom.Cells(i, "A").Value
        If dictEmpresas.Exists(key) Then
            wsSaidasDom.Cells(i, "B").Value = dictEmpresas(key)
        End If
    Next i

    ' Preenche a coluna B de "Entradas_Dom"
    lastRow = wsEntradasDom.Cells(wsEntradasDom.Rows.Count, "A").End(xlUp).Row
    For i = 5 To lastRow
        key = wsEntradasDom.Cells(i, "A").Value
        If dictEmpresas.Exists(key) Then
            wsEntradasDom.Cells(i, "B").Value = dictEmpresas(key)
        End If
    Next i

    ' Adiciona uma nova coluna entre A e B em "CFs_Dom"
    wsCFsDom.Columns("B").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove

    ' Preenche a nova coluna B de "CFs_Dom"
    lastRow = wsCFsDom.Cells(wsCFsDom.Rows.Count, "A").End(xlUp).Row
    For i = 7 To lastRow
        key = wsCFsDom.Cells(i, "A").Value
        If dictEmpresas.Exists(key) Then
            wsCFsDom.Cells(i, "B").Value = dictEmpresas(key)
        End If
    Next i

    MarcarRepeticoes

    
End Sub


Private Sub MarcarRepeticoes()

    Dim ws As Worksheet
    Dim ws2 As Worksheet
    Dim ws3 As Worksheet
    Dim lastRow As Long
    Dim lastRow2 As Long
    Dim lastRow3 As Long
    Dim i As Long
    Dim key As String
    Dim dictComb As Object
    Dim dictComb2 As Object
    Dim dictComb3 As Object
    
    
    ' Desativa atualizações e alertas para melhor desempenho
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ' Define as planilhas
    Set ws = ThisWorkbook.Sheets("NFe-NFCe_Sieg")
    Set ws2 = ThisWorkbook.Sheets("CFe_Sieg")
    Set ws3 = ThisWorkbook.Sheets("CTe_Sieg")

    ' Cria os dicionários para armazenar as combinações
    Set dictComb = CreateObject("Scripting.Dictionary")
    Set dictComb2 = CreateObject("Scripting.Dictionary")
    Set dictComb3 = CreateObject("Scripting.Dictionary")

    ' Encontra as últimas linhas preenchidas em cada planilha
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    lastRow2 = ws2.Cells(ws2.Rows.Count, "A").End(xlUp).Row
    lastRow3 = ws3.Cells(ws3.Rows.Count, "A").End(xlUp).Row

    ' Loop através das linhas na planilha "NFe-NFCe_Sieg"
    For i = 2 To lastRow
        ' Verifica se a célula em A não está vazia
        If ws.Cells(i, "A").Value <> "" Then
            ' Cria a chave de combinação baseada nos valores de A, D, G, J
            key = ws.Cells(i, "A").Value & "|" & ws.Cells(i, "D").Value & "|" & ws.Cells(i, "G").Value & "|" & ws.Cells(i, "J").Value
            
            ' Verifica se a combinação já existe no dicionário
            If dictComb.Exists(key) Then
                ' Se existir, adiciona "R" aos valores em D e G
                ws.Cells(i, "D").Value = "R" & ws.Cells(i, "D").Value
                ws.Cells(i, "G").Value = "R" & ws.Cells(i, "G").Value
            Else
                ' Se não existir, adiciona a combinação ao dicionário
                dictComb.Add key, 1
            End If
        End If
    Next i

    ' Loop através das linhas na planilha "CFe_Sieg"
    For i = 2 To lastRow2
        ' Verifica se a célula em A não está vazia
        If ws2.Cells(i, "A").Value <> "" Then
            ' Cria a chave de combinação baseada nos valores de A, C, D, F
            key = ws2.Cells(i, "A").Value & "|" & ws2.Cells(i, "C").Value & "|" & ws2.Cells(i, "D").Value & "|" & ws2.Cells(i, "F").Value
            
            ' Verifica se a combinação já existe no dicionário
            If dictComb2.Exists(key) Then
                ' Se existir, adiciona "R" aos valores em D e F
                ws2.Cells(i, "D").Value = "R" & ws2.Cells(i, "D").Value
                ws2.Cells(i, "F").Value = "R" & ws2.Cells(i, "F").Value
            Else
                ' Se não existir, adiciona a combinação ao dicionário
                dictComb2.Add key, 1
            End If
        End If
    Next i

    ' Loop através das linhas na planilha "CTe_Sieg"
    For i = 2 To lastRow3
        ' Verifica se a célula em A não está vazia
        If ws3.Cells(i, "A").Value <> "" Then
            ' Cria a chave de combinação baseada nos valores de A, N, P, AA
            key = ws3.Cells(i, "A").Value & "|" & ws3.Cells(i, "N").Value & "|" & ws3.Cells(i, "P").Value & "|" & ws3.Cells(i, "AA").Value
            
            ' Verifica se a combinação já existe no dicionário
            If dictComb3.Exists(key) Then
                ' Se existir, adiciona "R" aos valores em N e P
                ws3.Cells(i, "N").Value = "R" & ws3.Cells(i, "N").Value
                ws3.Cells(i, "P").Value = "R" & ws3.Cells(i, "P").Value
            Else
                ' Se não existir, adiciona a combinação ao dicionário
                dictComb3.Add key, 1
            End If
        End If
    Next i

    MarcarColunaPComR

End Sub





Private Sub MarcarColunaPComR()

    Dim ws As Worksheet
    Dim dict As Object
    Dim lastRow As Long
    Dim i As Long
    Dim valorColunaN As String
    Dim valorColunaP As String
    Dim valorColunaG As String

    ' Desativa atualizações e alertas para melhor desempenho
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    ' Definir a planilha "CTe_Sieg"
    Set ws = ThisWorkbook.Sheets("CTe_Sieg")
    
    ' Criar o dicionário
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' Encontrar a última linha preenchida na coluna N
    lastRow = ws.Cells(ws.Rows.Count, "N").End(xlUp).Row
    
    ' Preencher o dicionário com os valores da coluna N
    For i = 5 To lastRow
        valorColunaN = CStr(ws.Cells(i, "N").Value)
        valorColunaP = CStr(ws.Cells(i, "P").Value)
        valorColunaG = CStr(ws.Cells(i, "G").Value)
        
        ' Se os valores de N e P forem iguais, adicione "R" ao valor da coluna P
        If valorColunaN = valorColunaP Then 'Or valorColunaN <> valorColunaP Then
            ws.Cells(i, "N").Value = "R" & valorColunaN
        End If
        
         If valorColunaN <> valorColunaG Then 'Or valorColunaN <> valorColunaP Then
            ws.Cells(i, "N").Value = "R" & valorColunaN
        End If
       
        
        
    Next i

    
    Substituir37Por36
    
End Sub


Private Sub Substituir37Por36()

    Dim wsEntradasDom As Worksheet
    Dim wsSaidasDom As Worksheet
    

    ' Desativa atualizações e alertas para melhor desempenho
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False


    ' Define as planilhas
    Set wsEntradasDom = ThisWorkbook.Sheets("Entradas_Dom")
    Set wsSaidasDom = ThisWorkbook.Sheets("Saidas_Dom")

    ' Substitui o valor 37 por 36 na coluna T de "Entradas_Dom"
    With wsEntradasDom.Columns("T")
        .Replace What:="37", Replacement:="36", LookAt:=xlWhole, _
                 SearchOrder:=xlByRows, MatchCase:=False, _
                 SearchFormat:=False, ReplaceFormat:=False
    End With

    ' Substitui o valor 46 por 38 na coluna T de "Entradas_Dom"
    With wsEntradasDom.Columns("T")
        .Replace What:="46", Replacement:="38", LookAt:=xlWhole, _
                 SearchOrder:=xlByRows, MatchCase:=False, _
                 SearchFormat:=False, ReplaceFormat:=False
    End With

    ' Substitui o valor 37 por 36 na coluna T de "Saidas_Dom"
    With wsSaidasDom.Columns("T")
        .Replace What:="37", Replacement:="36", LookAt:=xlWhole, _
                 SearchOrder:=xlByRows, MatchCase:=False, _
                 SearchFormat:=False, ReplaceFormat:=False
    End With

    AdicionarSNaColunaBSeValorTNaoFor36Ou38

End Sub



Private Sub AdicionarSNaColunaBSeValorTNaoFor36Ou38()

    Dim wsEntradasDom As Worksheet, wsSaidasDom As Worksheet
    Dim dictValoresPermitidos As Object
    Dim lastRow As Long, lastRow2 As Long
    Dim i As Long
    Dim valorT As String
    Dim valorB As String
    
    ' Desativa atualizações e alertas para melhor desempenho
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    
    ' Definir a planilha
    Set wsEntradasDom = ThisWorkbook.Sheets("Entradas_Dom")
    Set wsSaidasDom = ThisWorkbook.Sheets("Saidas_Dom")

    
    ' Criar dicionário para os valores permitidos
    Set dictValoresPermitidos = CreateObject("Scripting.Dictionary")
    
    ' Adicionar os valores 36 e 38 ao dicionário
    dictValoresPermitidos.Add "36", True
    dictValoresPermitidos.Add "38", True
    dictValoresPermitidos.Add "41", True
    dictValoresPermitidos.Add "46", True
    
    ' Encontrar a última linha preenchida na coluna T
    lastRow = wsEntradasDom.Cells(wsEntradasDom.Rows.Count, "T").End(xlUp).Row
    lastRow2 = wsSaidasDom.Cells(wsSaidasDom.Rows.Count, "T").End(xlUp).Row
    
    ' Percorrer a coluna T
    For i = 5 To lastRow ' Começa da linha 5 conforme solicitado
        valorT = CStr(wsEntradasDom.Cells(i, "T").Value)
        
        ' Verificar se o valor de T não está no dicionário de valores permitidos
        If Not dictValoresPermitidos.Exists(valorT) Then
            valorB = CStr(wsEntradasDom.Cells(i, "B").Value)
            ' Adicionar "S" na frente do valor de B
            wsEntradasDom.Cells(i, "B").Value = "S" & valorB
        End If
    Next i

    ' Percorrer a coluna T
    For i = 5 To lastRow2 ' Começa da linha 5 conforme solicitado
        valorT = CStr(wsSaidasDom.Cells(i, "T").Value)
        
        ' Verificar se o valor de T não está no dicionário de valores permitidos
        If Not dictValoresPermitidos.Exists(valorT) Then
            valorB = CStr(wsSaidasDom.Cells(i, "B").Value)
            ' Adicionar "S" na frente do valor de B
            wsSaidasDom.Cells(i, "B").Value = "S" & valorB
        End If
    Next i


    ModificarValoresColunasDG

End Sub




Private Sub ModificarValoresColunasDG()

    Dim wsNFeNFCeSieg As Worksheet
    Dim dictColunaD As Object
    Dim dictColunaG As Object
    Dim dictColunaAD As Object
    Dim lastRow As Long
    Dim i As Long, valorD As String, valorG As String, valorAD As String
    

    ' Desativa atualizações e alertas para melhor desempenho
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    
    ' Definir a planilha
    Set wsNFeNFCeSieg = ThisWorkbook.Sheets("NFe-NFCe_Sieg")
    
    ' Criar dicionários para as colunas D, G e AD
    Set dictColunaD = CreateObject("Scripting.Dictionary")
    Set dictColunaG = CreateObject("Scripting.Dictionary")
    Set dictColunaAD = CreateObject("Scripting.Dictionary")
    
    ' Encontrar a última linha preenchida na coluna A
    lastRow = wsNFeNFCeSieg.Cells(wsNFeNFCeSieg.Rows.Count, "A").End(xlUp).Row
    
    ' Percorrer a planilha e preencher os dicionários com os valores das colunas D, G e AD
    For i = 5 To lastRow ' Assume que a primeira linha é cabeçalho
        If wsNFeNFCeSieg.Cells(i, "A").Value <> "" Then
            valorD = CStr(wsNFeNFCeSieg.Cells(i, "D").Value)
            valorG = CStr(wsNFeNFCeSieg.Cells(i, "G").Value)
            valorAD = CStr(wsNFeNFCeSieg.Cells(i, "AD").Value)
            
            ' Armazenar os valores nos dicionários
            dictColunaD.Add i, valorD
            dictColunaG.Add i, valorG
            dictColunaAD.Add i, valorAD
        End If
    Next i
    
    ' Percorrer as linhas novamente para aplicar as regras
    For i = 5 To lastRow
        If dictColunaD.Exists(i) Then
            valorD = dictColunaD(i)
            valorG = dictColunaG(i)
            valorAD = dictColunaAD(i)
            
            ' Verificar as condições e modificar os valores conforme necessário
            If valorD = valorG Then
                If valorAD = "Dev" Then
                    wsNFeNFCeSieg.Cells(i, "D").Value = "R" & valorG 'G
                Else
                    wsNFeNFCeSieg.Cells(i, "D").Value = "R" & valorD 'D
                End If
            End If
        End If
    Next i
    'ALTER

    PreencherColunaNComDicionario

End Sub


Private Sub PreencherColunaNComDicionario()
    Dim wsCFe As Worksheet, wsSIEG As Worksheet
    Dim lastRowCFe As Long, lastRowSIEG As Long
    Dim i As Long, chave As String
    Dim dict As Object

    ' Desativa atualizações e alertas para melhor desempenho
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    
    ' Definir as planilhas
    Set wsCFe = ThisWorkbook.Sheets("CFe_Sieg")
    Set wsSIEG = ThisWorkbook.Sheets("SIEG")
    
    ' Obter a última linha preenchida em cada planilha
    lastRowCFe = wsCFe.Cells(wsCFe.Rows.Count, "A").End(xlUp).Row
    lastRowSIEG = wsSIEG.Cells(wsSIEG.Rows.Count, "A").End(xlUp).Row
    
    ' Criar o dicionário
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' Preencher o dicionário com os valores de SIEG
    For i = 5 To lastRowSIEG
            chave = wsSIEG.Cells(i, 1).Value & "|" & wsSIEG.Cells(i, 3).Value & "|" & 99
            dict(chave) = wsSIEG.Cells(i, 13).Value ' Coluna M em SIEG
    Next i
    
    ' Percorrer a coluna A de CFe_Sieg a partir de A6
    For i = 6 To lastRowCFe
        If wsCFe.Cells(i, 1).Value <> "" Then
            chave = wsCFe.Cells(i, 1).Value & "|" & wsCFe.Cells(i, 4).Value & "|" & 99  ' Coluna A e D em CFe_Sieg
            
            ' Verificar se a chave existe no dicionário
            If dict.Exists(chave) Then
                wsCFe.Cells(i, 14).Value = dict(chave) ' Preencher a coluna N em CFe_Sieg
            End If
        End If
    Next i
    
    
    ' Percorrer a coluna A de CFe_Sieg a partir de A6
    For i = 6 To lastRowCFe
        If wsCFe.Cells(i, 1).Value <> "" And wsCFe.Cells(i, 14) = "" Or wsCFe.Cells(i, 14) = "210210" Then
                wsCFe.Cells(i, 14).Value = "Autorizado o uso do CFe"
        End If
    Next i
    
    BSaidas.BSaida
    
End Sub




