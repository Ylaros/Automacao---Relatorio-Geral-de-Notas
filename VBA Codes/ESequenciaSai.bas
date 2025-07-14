Attribute VB_Name = "ESequenciaSai"
'Aqui come�a notas faltantes

Sub ENotasFaltantesSaidas()

    Dim wsNFe As Worksheet
    Dim wsNNLSaidas As Worksheet
    Dim dictSaidas As Object
    Dim lastRowNFe As Long
    Dim i As Long, j As Long
    Dim key As String
    
    ' Desativa atualiza��es e alertas para melhor desempenho
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    ' Definir a planilha "NFe-NFCe_Sieg"
    Set wsNFe = ThisWorkbook.Sheets("NFe-NFCe_Sieg")
    
    ' Criar uma nova aba chamada "NNL-Saidas"
    On Error Resume Next
    Set wsNNLSaidas = ThisWorkbook.Sheets("NNL-Saidas")
    On Error GoTo 0
    
    If wsNNLSaidas Is Nothing Then
        Set wsNNLSaidas = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsNNLSaidas.Name = "NNL-Saidas"
    End If

    ' Escrever os cabe�alhos na aba "NNL-Saidas"
    With wsNNLSaidas
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
    Set dictSaidas = CreateObject("Scripting.Dictionary")
    
    ' Encontrar a �ltima linha preenchida na aba "NFe-NFCe_Sieg"
    lastRowNFe = wsNFe.Cells(wsNFe.Rows.Count, "A").End(xlUp).Row
    
    ' Preencher o dicion�rio com os valores da planilha "NFe-NFCe_Sieg"
    j = 2 ' Iniciar na linha 2 da aba "NNL-Saidas"
    
    ' Processar as linhas onde "AD" = "Sai"
    For i = 2 To lastRowNFe
        If wsNFe.Cells(i, "AD").Value <> "XXX" And wsNFe.Cells(i, "A") <> "" Then
            key = CStr(wsNFe.Cells(i, "G").Value)
            If Not dictSaidas.Exists(key) Then
                dictSaidas.Add key, Array(wsNFe.Cells(i, "K").Value, wsNFe.Cells(i, "A").Value, wsNFe.Cells(i, "AB").Value, wsNFe.Cells(i, "J").Value)
            End If
            
            ' Copiar os valores para a aba "NNL-Saidas"
            With wsNNLSaidas
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

    ' Ativar a aba "NNL-Saidas"
    wsNNLSaidas.Activate
    
    SeqPreencherNNLSaidasComValoresFaltantes

End Sub




Private Sub SeqPreencherNNLSaidasComValoresFaltantes()
    Dim wsNNLSaidas As Worksheet
    Dim wsSaidasDom As Worksheet
    Dim dictNNLSaidas As Object
    Dim dictSaidasDom As Object
    Dim lastRowNNL As Long, lastRowSaidasDom As Long
    Dim i As Long, lastRowNew As Long
    Dim BValue As String, EValue As String, DValue As String, FValue As String, IValue As String, TValue As String
    Dim key As String
    
    ' Desativa atualiza��es e alertas para melhor desempenho
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    ' Definir as planilhas
    Set wsNNLSaidas = ThisWorkbook.Sheets("NNL-Saidas")
    Set wsSaidasDom = ThisWorkbook.Sheets("Saidas_Dom")
    
    ' Criar dicion�rio para armazenar as combina��es de "NNL-Saidas"
    Set dictNNLSaidas = CreateObject("Scripting.Dictionary")
    
    ' Criar dicion�rio para armazenar as combina��es �nicas de "Saidas_Dom"
    Set dictSaidasDom = CreateObject("Scripting.Dictionary")
    
    ' Encontrar a �ltima linha preenchida em "NNL-Saidas"
    lastRowNNL = wsNNLSaidas.Cells(wsNNLSaidas.Rows.Count, "C").End(xlUp).Row
    
    ' Preencher o dicion�rio com combina��es de "C" e "E" em "NNL-Saidas"
    For i = 2 To lastRowNNL
        BValue = CStr(wsNNLSaidas.Cells(i, "C").Value)
        EValue = CStr(wsNNLSaidas.Cells(i, "E").Value)
        DValue = CStr(wsNNLSaidas.Cells(i, "D").Value)
        FValue = CStr(wsNNLSaidas.Cells(i, "F").Value)
        key = BValue & "|" & EValue & "|" & DValue & "|" & FValue
        
        If Not dictNNLSaidas.Exists(key) Then
            dictNNLSaidas.Add key, True
        End If
    Next i
    
    ' Encontrar a �ltima linha preenchida em "Saidas_Dom"
    lastRowSaidasDom = wsSaidasDom.Cells(wsSaidasDom.Rows.Count, "B").End(xlUp).Row
    
    ' Preencher "NNL-Saidas" com valores de "Saidas_Dom" onde a combina��o n�o � encontrada
    For i = 5 To lastRowSaidasDom
        BValue = CStr(wsSaidasDom.Cells(i, "B").Value)
        EValue = CStr(wsSaidasDom.Cells(i, "E").Value)
        IValue = CStr(wsSaidasDom.Cells(i, "I").Value)
        TValue = CStr(wsSaidasDom.Cells(i, "T").Value)
        key = BValue & "|" & EValue & "|" & IValue & "|" & TValue
        
        ' Verificar se a combina��o j� foi adicionada ao dicion�rio de "Saidas_Dom"
        If Not dictSaidasDom.Exists(key) Then
            dictSaidasDom.Add key, True
            
            ' Verificar se a combina��o n�o existe em "NNL-Saidas"
            If Not dictNNLSaidas.Exists(key) Then
                lastRowNew = wsNNLSaidas.Cells(wsNNLSaidas.Rows.Count, "C").End(xlUp).Row + 1
                wsNNLSaidas.Cells(lastRowNew, "C").Value = wsSaidasDom.Cells(i, "B").Value
                wsNNLSaidas.Cells(lastRowNew, "D").Value = wsSaidasDom.Cells(i, "I").Value
                wsNNLSaidas.Cells(lastRowNew, "E").Value = wsSaidasDom.Cells(i, "E").Value
                wsNNLSaidas.Cells(lastRowNew, "F").Value = wsSaidasDom.Cells(i, "T").Value
                wsNNLSaidas.Cells(lastRowNew, "G").Value = "Dominio"
                wsNNLSaidas.Cells(lastRowNew, "H").Value = "N"
            End If
        End If
    Next i

    SeqFiltrarNNLSaidas

End Sub



Private Sub SeqFiltrarNNLSaidas()

    Dim wsCompSaidas As Worksheet
    Dim wsContSaidas As Worksheet
    Dim dictContSaidas As Object
    Dim lastRowComp As Long, lastRowCont As Long
    Dim i As Long, key As String

    ' Desativa atualiza��es e alertas para melhor desempenho
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ' Definir as planilhas
    Set wsCompSaidas = ThisWorkbook.Sheets("NNL-Saidas")
    Set wsContSaidas = ThisWorkbook.Sheets("Cont-Saidas")
    
    ' Criar dicion�rio para armazenar os valores da coluna C de "Cont-Saidas"
    Set dictContSaidas = CreateObject("Scripting.Dictionary")
    
    ' Encontrar as �ltimas linhas preenchidas em ambas as planilhas
    lastRowComp = wsCompSaidas.Cells(wsCompSaidas.Rows.Count, "C").End(xlUp).Row
    lastRowCont = wsContSaidas.Cells(wsContSaidas.Rows.Count, "C").End(xlUp).Row
    
    ' Preencher o dicion�rio com os valores da coluna C de "Cont-Saidas"
    For i = 3 To lastRowCont ' Come�a de C3 conforme solicitado
        key = CStr(wsContSaidas.Cells(i, "C").Value)
        If Not dictContSaidas.Exists(key) Then
            dictContSaidas.Add key, True
        End If
    Next i
    
    ' Verificar e apagar linhas da aba "Comp-Saidas" cujos valores de C n�o est�o em "Cont-Saidas"
    For i = lastRowComp To 2 Step -1 ' Percorrer de baixo para cima para evitar problemas ao excluir linhas
        key = CStr(wsCompSaidas.Cells(i, "C").Value)
        If Not dictContSaidas.Exists(key) Then
            wsCompSaidas.Rows(i).Delete
        End If
    Next i

    EVerificarNotasFaltantes
    

End Sub



Private Sub EVerificarNotasFaltantes()
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
    Set wsOrigem = ThisWorkbook.Sheets("NNL-Saidas")
    
    ' Criar a planilha de destino se n�o existir
    On Error Resume Next
    Set wsDestino = ThisWorkbook.Sheets("NNLs-Saidas")
    On Error GoTo 0
    If wsDestino Is Nothing Then
        Set wsDestino = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsDestino.Name = "NNLs-Saidas"
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
    ThisWorkbook.Sheets("NNL-Saidas").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0


    OrganizarNotasSequenciais

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

Private Sub OrganizarNotasSequenciais()
    Dim ws As Worksheet
    Dim ultimaLinha As Long
    Dim cnpjAtual As String
    Dim cnpjAnterior As String
    Dim primeiraNota As Long
    Dim ultimaNota As Long
    Dim notaAtual As Long
    Dim i As Long
    Dim sequencia As String
    Dim linhaDestino As Long
    Dim linhasParaDeletar As Range
    
    ' Definir a planilha "NNLS-Saidas"
    Set ws = ThisWorkbook.Sheets("NNLS-Saidas")
    
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
    
        PreencherColunasAEBNNLs
End Sub




Private Sub PreencherColunasAEBNNLs()
    Dim wsCompSaidas As Worksheet
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
    Set wsCompSaidas = ThisWorkbook.Sheets("NNLs-Saidas")
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
    
    ' Encontrar a �ltima linha preenchida em "Comp-Saidas"
    lastRowComp = wsCompSaidas.Cells(wsCompSaidas.Rows.Count, "C").End(xlUp).Row
    
    ' Percorrer "Comp-Saidas" e preencher colunas A e B com base no dicion�rio
    For i = 2 To lastRowComp
        valorC = CStr(wsCompSaidas.Cells(i, "C").Value)
        
        ' Verificar se o valor de C existe no dicion�rio
        If dictEmpresasDom.Exists(valorC) Then
            wsCompSaidas.Cells(i, "A").Value = dictEmpresasDom(valorC)(0) ' Preenche a coluna A
            wsCompSaidas.Cells(i, "B").Value = dictEmpresasDom(valorC)(1) ' Preenche a coluna B
        End If
    Next i
    
   
    ' Ajustar a largura das colunas para melhor visualiza��o
    wsCompSaidas.Columns("A:J").AutoFit
    
   ' Ordenar os dados pela coluna A de maneira crescente, ignorando o cabe�alho
    wsCompSaidas.Sort.SortFields.Clear
    wsCompSaidas.Sort.SortFields.Add key:=wsCompSaidas.Range("A2:A" & lastRowComp), Order:=xlAscending
    With wsCompSaidas.Sort
        .SetRange wsCompSaidas.Range("A1:D" & lastRowComp) ' Ajuste o intervalo conforme necess�rio
        .Header = xlYes ' Considerar o cabe�alho
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    FSequenciaCF.FVerificarNotasFaltantesCF
    
End Sub



