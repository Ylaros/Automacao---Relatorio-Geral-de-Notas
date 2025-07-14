Attribute VB_Name = "FSequenciaCF"
'Aqui come�a notas faltantes

Sub FVerificarNotasFaltantesCF()

    Dim wsCFe As Worksheet
    Dim wsNNLCFs As Worksheet
    Dim dictCFs As Object
    Dim lastRowCFe As Long
    Dim i As Long, j As Long
    Dim key As String
    
    ' Desativa atualiza��es e alertas para melhor desempenho
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    ' Definir a planilha "CFe_Sieg"
    Set wsCFe = ThisWorkbook.Sheets("CFe_Sieg")
    
    ' Criar uma nova aba chamada "NNL-CFe"
    On Error Resume Next
    Set wsNNLCFs = ThisWorkbook.Sheets("NNL-CFe")
    On Error GoTo 0
    
    If wsNNLCFs Is Nothing Then
        Set wsNNLCFs = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsNNLCFs.Name = "NNL-CFe"
    End If

    ' Escrever os cabe�alhos na aba "NNL-CFe"
    With wsNNLCFs
        .Cells(1, 1).Value = "Empresa"
        .Cells(1, 2).Value = "Descri��o"
        .Cells(1, 3).Value = "CNPJ"
        .Cells(1, 4).Value = "Data"
        .Cells(1, 5).Value = "Nota"
        .Cells(1, 6).Value = "Esp�cie"
        .Cells(1, 7).Value = "Status"
        .Cells(1, 8).Value = "Valor Sieg"
        .Cells(1, 9).Value = "Valor Dom"
        .Cells(1, 10).Value = "Mensagem"
    End With
    
    ' Criar um dicion�rio para armazenar as sa�das
    Set dictCFs = CreateObject("Scripting.Dictionary")
    
    ' Encontrar a �ltima linha preenchida na aba "CFe_Sieg"
    lastRowCFe = wsCFe.Cells(wsCFe.Rows.Count, "A").End(xlUp).Row
    
    ' Preencher o dicion�rio com os valores da planilha "CFe_Sieg"
    j = 2 ' Iniciar na linha 2 da aba "NNL-CFe"
    
    ' Processar as linhas onde "AD" = "Sai"
    For i = 6 To lastRowCFe
        If wsCFe.Cells(i, "A").Value <> "" Then
            key = CStr(wsCFe.Cells(i, "D").Value)
            If Not dictCFs.Exists(key) Then
                dictCFs.Add key, Array(wsCFe.Cells(i, "D").Value, wsCFe.Cells(i, "C").Value, wsCFe.Cells(i, "A").Value, wsCFe.Cells(i, "I").Value)
            End If
            
            ' Copiar os valores para a aba "NNL-CFe"
            With wsNNLCFs
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

    ' Ativar a aba "NNL-CFe"
    wsNNLCFs.Activate
    
    PreencherNNLCFsComValoresFaltantes

End Sub




Private Sub PreencherNNLCFsComValoresFaltantes()
    Dim wsNNLCFs As Worksheet
    Dim wsCFsDom As Worksheet
    Dim dictNNLCFs As Object
    Dim dictCFsDom As Object
    Dim lastRowNNL As Long, lastRowCFsDom As Long
    Dim i As Long, lastRowNew As Long
    Dim BValue As String, EValue As String, DValue As String, IValue As String
    Dim key As String
    
    ' Desativa atualiza��es e alertas para melhor desempenho
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    ' Definir as planilhas
    Set wsNNLCFs = ThisWorkbook.Sheets("NNL-CFe")
    Set wsCFsDom = ThisWorkbook.Sheets("CFs_Dom")
    
    ' Criar dicion�rio para armazenar as combina��es de "NNL-CFe"
    Set dictNNLCFs = CreateObject("Scripting.Dictionary")
    
    ' Criar dicion�rio para armazenar as combina��es �nicas de "CFs_Dom"
    Set dictCFsDom = CreateObject("Scripting.Dictionary")
    
    ' Encontrar a �ltima linha preenchida em "NNL-CFe"
    lastRowNNL = wsNNLCFs.Cells(wsNNLCFs.Rows.Count, "C").End(xlUp).Row
    
    ' Preencher o dicion�rio com combina��es de "C" e "E" em "NNL-CFe"
    For i = 2 To lastRowNNL
        BValue = CStr(wsNNLCFs.Cells(i, "C").Value)
        EValue = CStr(wsNNLCFs.Cells(i, "D").Value)
        DValue = CStr(wsNNLCFs.Cells(i, "E").Value)
        key = BValue & "|" & EValue & "|" & DValue
        
        If Not dictNNLCFs.Exists(key) Then
            dictNNLCFs.Add key, True
        End If
    Next i
    
    ' Encontrar a �ltima linha preenchida em "CFs_Dom"
    lastRowCFsDom = wsCFsDom.Cells(wsCFsDom.Rows.Count, "B").End(xlUp).Row
    
    ' Preencher "NNL-CFe" com valores de "CFs_Dom" onde a combina��o n�o � encontrada
    For i = 5 To lastRowCFsDom
        BValue = CStr(wsCFsDom.Cells(i, "B").Value)
        EValue = CStr(wsCFsDom.Cells(i, "C").Value)
        IValue = CStr(wsCFsDom.Cells(i, "D").Value)
        key = BValue & "|" & EValue & "|" & IValue
        
        ' Verificar se a combina��o j� foi adicionada ao dicion�rio de "CFs_Dom"
        If Not dictCFsDom.Exists(key) Then
            dictCFsDom.Add key, True
            
            ' Verificar se a combina��o n�o existe em "NNL-CFe"
            If Not dictNNLCFs.Exists(key) Then
                lastRowNew = wsNNLCFs.Cells(wsNNLCFs.Rows.Count, "C").End(xlUp).Row + 1
                wsNNLCFs.Cells(lastRowNew, "C").Value = wsCFsDom.Cells(i, "B").Value
                wsNNLCFs.Cells(lastRowNew, "D").Value = wsCFsDom.Cells(i, "C").Value
                wsNNLCFs.Cells(lastRowNew, "E").Value = wsCFsDom.Cells(i, "D").Value
                wsNNLCFs.Cells(lastRowNew, "F").Value = "CFe"
                wsNNLCFs.Cells(lastRowNew, "G").Value = "Dominio"
                wsNNLCFs.Cells(lastRowNew, "H").Value = "N"
            End If
        End If
    Next i

    FiltrarNNLCFs

End Sub



Private Sub FiltrarNNLCFs()

    Dim wsNNLCFs As Worksheet
    Dim wsContCFs As Worksheet
    Dim dictContCFs As Object
    Dim lastRowNNL As Long, lastRowCont As Long
    Dim i As Long, key As String
    
    ' Desativa atualiza��es e alertas para melhor desempenho
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ' Definir as planilhas
    Set wsNNLCFs = ThisWorkbook.Sheets("NNL-CFe")
    Set wsContCFs = ThisWorkbook.Sheets("Cont-CFe")
    
    ' Criar dicion�rio para armazenar os valores da coluna C de "Cont-CFs"
    Set dictContCFs = CreateObject("Scripting.Dictionary")
    
    ' Encontrar as �ltimas linhas preenchidas em ambas as planilhas
    lastRowNNL = wsNNLCFs.Cells(wsNNLCFs.Rows.Count, "C").End(xlUp).Row
    lastRowCont = wsContCFs.Cells(wsContCFs.Rows.Count, "C").End(xlUp).Row
    
    ' Preencher o dicion�rio com os valores da coluna C de "Cont-CFs"
    For i = 3 To lastRowCont ' Come�a de C3 conforme solicitado
        key = CStr(wsContCFs.Cells(i, "C").Value)
        If Not dictContCFs.Exists(key) Then
            dictContCFs.Add key, True
        End If
    Next i
    
    ' Verificar e apagar linhas da aba "NNL-CFe" cujos valores de C n�o est�o em "Cont-CFs"
    For i = lastRowNNL To 2 Step -1 ' Percorrer de baixo para cima para evitar problemas ao excluir linhas
        key = CStr(wsNNLCFs.Cells(i, "C").Value)
        If Not dictContCFs.Exists(key) Then
            wsNNLCFs.Rows(i).Delete
        End If
    Next i

    'RemoverLinhasDuplicadas
    CFVerificarNotasFaltantes

End Sub

Sub CFVerificarNotasFaltantes()
    Dim wsOrigem As Worksheet
    Dim wsDestino As Worksheet
    Dim rngDados As Range
    Dim celula As Range
    Dim cnpjAtual As String
    Dim ultimaLinha As Long
    Dim dictNotas As Object
    Dim notaAtual As Long
    Dim notaProxima As Long
    Dim i As Long
    Dim linhaDestino As Long
    Dim cnpjAnterior As String
    Dim notasLista As Variant
    Dim j As Long
    
    ' Desativa atualiza��es e alertas para melhor desempenho
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    ' Definir a planilha de origem e criar a de destino
    Set wsOrigem = ThisWorkbook.Sheets("NNL-CFe")
    
    ' Criar a planilha de destino se n�o existir
    On Error Resume Next
    Set wsDestino = ThisWorkbook.Sheets("NNLs-CFe")
    On Error GoTo 0
    If wsDestino Is Nothing Then
        Set wsDestino = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsDestino.Name = "NNLs-CFe"
    Else
        wsDestino.Cells.Clear
    End If
    
    ' Definir a �rea de dados
    ultimaLinha = wsOrigem.Cells(wsOrigem.Rows.Count, "C").End(xlUp).Row
    Set rngDados = wsOrigem.Range("C2:E" & ultimaLinha)
    
    ' Configurar a planilha de destino
    wsDestino.Cells(1, 1).Value = "Empresa"
    wsDestino.Cells(1, 2).Value = "Descri��o"
    wsDestino.Cells(1, 3).Value = "CNPJ"
    wsDestino.Cells(1, 4).Value = "Nota Faltante"
    
    linhaDestino = 2
    
    ' Configurar o dicion�rio para armazenar as notas
    Set dictNotas = CreateObject("Scripting.Dictionary")
    
    ' Inicializar vari�veis
    cnpjAnterior = ""
    
    ' Percorrer os dados
    For Each celula In rngDados.Columns(1).Cells
        If celula.Row > 1 Then
            If celula.Value <> cnpjAnterior Then
                ' Novo CNPJ encontrado
                If cnpjAnterior <> "" Then
                    ' Processar o CNPJ anterior
                    If dictNotas.Count > 0 Then
                        ' Obter as notas para o CNPJ atual
                        notasLista = dictNotas.keys
                        Call SortArray(notasLista)
                        
                        For i = LBound(notasLista) To UBound(notasLista) - 1
                            notaAtual = CLng(notasLista(i))
                            notaProxima = CLng(notasLista(i + 1))
                            
                            ' Verificar se a diferen�a entre notas � menor ou igual a 150
                            If notaProxima - notaAtual > 1 And notaProxima - notaAtual <= 150 Then
                                For j = notaAtual + 1 To notaProxima - 1
                                    ' Registrar nota faltante
                                    wsDestino.Cells(linhaDestino, 3).Value = cnpjAnterior
                                    wsDestino.Cells(linhaDestino, 4).Value = j
                                    linhaDestino = linhaDestino + 1
                                Next j
                            End If
                        Next i
                    End If
                End If
                
                ' Inicializar o dicion�rio para o novo CNPJ
                dictNotas.RemoveAll
                cnpjAnterior = celula.Value
            End If
            
            ' Adicionar a nota atual ao dicion�rio
            notaAtual = celula.Offset(0, 2).Value
            dictNotas(CStr(notaAtual)) = True
        End If
    Next celula
    
    ' Processar o �ltimo CNPJ
    If dictNotas.Count > 0 Then
        ' Obter as notas para o CNPJ atual
        notasLista = dictNotas.keys
        Call SortArray(notasLista)
        
        For i = LBound(notasLista) To UBound(notasLista) - 1
            notaAtual = CLng(notasLista(i))
            notaProxima = CLng(notasLista(i + 1))
            
            ' Verificar se a diferen�a entre notas � menor ou igual a 150
            If notaProxima - notaAtual > 1 And notaProxima - notaAtual <= 150 Then
                For j = notaAtual + 1 To notaProxima - 1
                    ' Registrar nota faltante
                    wsDestino.Cells(linhaDestino, 3).Value = cnpjAnterior
                    wsDestino.Cells(linhaDestino, 4).Value = j
                    linhaDestino = linhaDestino + 1
                Next j
            End If
        Next i
    End If
    
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets("NNL-CFe").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0


    'CFOrganizarNotasSequenciais
    ExcluirLinhasNNLsComBaseEmCFsDom

End Sub

' Fun��o para ordenar um array
Sub SortArray(ByRef arr As Variant)
    Dim i As Long, j As Long
    Dim temp As Variant
    
    For i = LBound(arr) To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If CLng(arr(i)) > CLng(arr(j)) Then
                temp = arr(i)
                arr(i) = arr(j)
                arr(j) = temp
            End If
        Next j
    Next i
End Sub


Sub ExcluirLinhasNNLsComBaseEmCFsDom()
    Dim wsNNL As Worksheet, wsDom As Worksheet
    Dim lastRowNNL As Long, lastRowDom As Long
    Dim i As Long
    Dim dictDom As Object
    Dim chaveDom As String, chaveNNL As String

    ' Cria o dicion�rio
    Set dictDom = CreateObject("Scripting.Dictionary")

    ' Define as planilhas
    Set wsNNL = ThisWorkbook.Sheets("NNLs-CFe")
    Set wsDom = ThisWorkbook.Sheets("CFs_Dom")

    ' Encontra a �ltima linha de cada planilha
    lastRowNNL = wsNNL.Cells(wsNNL.Rows.Count, "C").End(xlUp).Row
    lastRowDom = wsDom.Cells(wsDom.Rows.Count, "B").End(xlUp).Row

    ' Desativa atualiza��es para desempenho
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ' Monta o dicion�rio com chaves B|D da aba CFs_Dom
    For i = 2 To lastRowDom
        If Trim(wsDom.Cells(i, "B").Value) <> "" And Trim(wsDom.Cells(i, "D").Value) <> "" Then
            chaveDom = Trim(wsDom.Cells(i, "B").Value) & "|" & Trim(wsDom.Cells(i, "D").Value)
            dictDom(chaveDom) = True
        End If
    Next i

    ' Percorre NNLs-CFe de baixo para cima
    For i = lastRowNNL To 2 Step -1
        If Trim(wsNNL.Cells(i, "C").Value) <> "" And Trim(wsNNL.Cells(i, "D").Value) <> "" Then
            chaveNNL = Trim(wsNNL.Cells(i, "C").Value) & "|" & Trim(wsNNL.Cells(i, "D").Value)
            If dictDom.Exists(chaveNNL) Then
                wsNNL.Rows(i).Delete
            End If
        End If
    Next i

    ' Reativa atualiza��es
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

    OrganizarNotasSequenciaisCorrigido
End Sub



Private Sub OrganizarNotasSequenciaisCorrigido()
    Dim ws As Worksheet
    Dim ultimaLinha As Long
    Dim cnpjAtual As String
    Dim cnpjAnterior As String
    Dim primeiraNota As Long
    Dim ultimaNota As Long
    Dim notaAtual As Long
    Dim i As Long
    Dim linhaDestino As Long
    Dim linhasParaDeletar As Range
    
    ' Definir a planilha "NNLS-Saidas"
    Set ws = ThisWorkbook.Sheets("NNLS-CFe")
    
    ' Definir a �ltima linha com dados na coluna C (CNPJs)
    ultimaLinha = ws.Cells(ws.Rows.Count, "C").End(xlUp).Row
    
    ' Inicializar vari�veis
    cnpjAnterior = ""
    linhaDestino = 2 ' Come�a na linha 2 (Assumindo que a linha 1 tem os cabe�alhos)
    Set linhasParaDeletar = Nothing ' Vari�vel para armazenar as linhas a serem deletadas
    
    ' Percorrer as linhas da planilha
    For i = 2 To ultimaLinha
        cnpjAtual = ws.Cells(i, "C").Value
        notaAtual = ws.Cells(i, "D").Value
        
        ' Se mudar o CNPJ ou chegar na �ltima linha
        If cnpjAtual <> cnpjAnterior Or i = ultimaLinha Then
            ' Se houver uma sequ�ncia para o CNPJ anterior, escrever a sequ�ncia
            If cnpjAnterior <> "" Then
                If primeiraNota = ultimaNota Then
                    ws.Cells(linhaDestino, "D").Value = primeiraNota
                Else
                    ws.Cells(linhaDestino, "D").Value = primeiraNota & "-" & ultimaNota
                End If
            End If
            
            ' Reiniciar a sequ�ncia para o novo CNPJ
            primeiraNota = notaAtual
            ultimaNota = notaAtual
            cnpjAnterior = cnpjAtual
            linhaDestino = i
        Else
            ' Se o CNPJ for o mesmo
            If notaAtual = ultimaNota + 1 Then
                ' Se a nota for sequencial, atualizar a �ltima nota
                ultimaNota = notaAtual
                ' Marcar a linha para deletar (exceto a primeira da sequ�ncia)
                If linhasParaDeletar Is Nothing Then
                    Set linhasParaDeletar = ws.Rows(i)
                Else
                    Set linhasParaDeletar = Union(linhasParaDeletar, ws.Rows(i))
                End If
            Else
                ' Se a sequ�ncia foi interrompida, escrever a sequ�ncia
                If primeiraNota = ultimaNota Then
                    ws.Cells(linhaDestino, "D").Value = primeiraNota
                Else
                    ws.Cells(linhaDestino, "D").Value = primeiraNota & "-" & ultimaNota
                End If
                
                ' Iniciar uma nova sequ�ncia
                linhaDestino = i
                primeiraNota = notaAtual
                ultimaNota = notaAtual
            End If
        End If
    Next i
    
    ' Ap�s o loop, capturar a �ltima sequ�ncia (se houver)
    If cnpjAnterior <> "" Then
        If primeiraNota = ultimaNota Then
            ws.Cells(linhaDestino, "D").Value = primeiraNota
        Else
            ws.Cells(linhaDestino, "D").Value = primeiraNota & "-" & ultimaNota
        End If
    End If
    
    ' Deletar as linhas marcadas como parte de sequ�ncias j� consolidadas
    If Not linhasParaDeletar Is Nothing Then
        linhasParaDeletar.Delete
    End If
    
    CFPreencherColunasAEBNNLs
    
End Sub



Private Sub CFPreencherColunasAEBNNLs()
    Dim wsCompCFe As Worksheet
    Dim wsEmpresasDom As Worksheet
    Dim dictEmpresasDom As Object
    Dim lastRowComp As Long, lastRowEmpresas As Long
    Dim i As Long
    Dim valorC As String, chave As String
    Dim valorA As String, valorB As String
    
    ' Desativa atualiza��es e alertas para melhor desempenho
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    ' Definir as planilhas
    Set wsCompCFe = ThisWorkbook.Sheets("NNLs-CFe")
    Set wsEmpresasDom = ThisWorkbook.Sheets("Empresas_Dom")
    
    ' Criar dicion�rio para armazenar os valores de "Empresas_Dom"
    Set dictEmpresasDom = CreateObject("Scripting.Dictionary")
    
    ' Encontrar a �ltima linha preenchida em "Empresas_Dom"
    lastRowEmpresas = wsEmpresasDom.Cells(wsEmpresasDom.Rows.Count, "I").End(xlUp).Row
    
    ' Preencher o dicion�rio com valores de "Empresas_Dom" (colunas I, A e C)
    For i = 2 To lastRowEmpresas
        chave = CStr(wsEmpresasDom.Cells(i, "I").Value)
        valorA = wsEmpresasDom.Cells(i, "A").Value
        valorB = wsEmpresasDom.Cells(i, "G").Value
        
        ' Adicionar a chave e os valores ao dicion�rio
        If Not dictEmpresasDom.Exists(chave) Then
            dictEmpresasDom.Add chave, Array(valorA, valorB)
        End If
    Next i
    
    ' Encontrar a �ltima linha preenchida em "Comp-CFe"
    lastRowComp = wsCompCFe.Cells(wsCompCFe.Rows.Count, "C").End(xlUp).Row
    
    ' Percorrer "Comp-CFe" e preencher colunas A e B com base no dicion�rio
    For i = 2 To lastRowComp
        valorC = CStr(wsCompCFe.Cells(i, "C").Value)
        
        ' Verificar se o valor de C existe no dicion�rio
        If dictEmpresasDom.Exists(valorC) Then
            wsCompCFe.Cells(i, "A").Value = dictEmpresasDom(valorC)(0) ' Preenche a coluna A
            wsCompCFe.Cells(i, "B").Value = dictEmpresasDom(valorC)(1) ' Preenche a coluna B
        End If
    Next i
    
   
    ' Ajustar a largura das colunas para melhor visualiza��o
    wsCompCFe.Columns("A:J").AutoFit
    
   ' Ordenar os dados pela coluna A de maneira crescente, ignorando o cabe�alho
    wsCompCFe.Sort.SortFields.Clear
    wsCompCFe.Sort.SortFields.Add key:=wsCompCFe.Range("A2:A" & lastRowComp), Order:=xlAscending
    With wsCompCFe.Sort
        .SetRange wsCompCFe.Range("A1:D" & lastRowComp) ' Ajuste o intervalo conforme necess�rio
        .Header = xlYes ' Considerar o cabe�alho
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    ZFormatar.ZFormatarAbas
    
End Sub











