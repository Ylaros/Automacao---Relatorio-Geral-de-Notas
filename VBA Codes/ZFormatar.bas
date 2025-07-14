Attribute VB_Name = "ZFormatar"
Sub ZFormatarAbas()
    Dim ws As Worksheet
    Dim celula As Range
    Dim ultimaLinha As Long
    Dim valorCNPJ As String
    Dim cnpjComMascara As String
    Dim planilhas As Variant
    Dim i As Integer

    ' Desativa atualizações de tela
    Application.ScreenUpdating = False


    ' Definir as planilhas onde a máscara será aplicada
    planilhas = Array("Cont-Saidas", "Cont-Entradas", "Cont-CFe")
    
    ' Percorrer cada planilha da lista
    For i = LBound(planilhas) To UBound(planilhas)
        Set ws = ThisWorkbook.Sheets(planilhas(i))
        
        ' Determinar a última linha da coluna C (a partir da linha 3)
        ultimaLinha = ws.Cells(ws.Rows.Count, "C").End(xlUp).Row
        
        ' Aplicar a máscara para cada célula da coluna C, a partir da linha 3
        For Each celula In ws.Range("C3:C" & ultimaLinha)
            valorCNPJ = CStr(celula.Value)
            
            ' Completar com zeros à esquerda para garantir 14 dígitos
            valorCNPJ = Right("00000000000000" & valorCNPJ, 14)
            
            ' Aplicar a máscara de CNPJ (XX.XXX.XXX/XXXX-XX)
            cnpjComMascara = Left(valorCNPJ, 2) & "." & Mid(valorCNPJ, 3, 3) & "." & Mid(valorCNPJ, 6, 3) & "/" & Mid(valorCNPJ, 9, 4) & "-" & Right(valorCNPJ, 2)
            
            ' Atualizar a célula com o CNPJ formatado
            celula.Value = cnpjComMascara
        Next celula
    Next i
    
    AplicarMascaraCNPJCorretoVariasAbas
    
    
    
End Sub

Private Sub AplicarMascaraCNPJCorretoVariasAbas()
    Dim ws As Worksheet
    Dim celula As Range
    Dim ultimaLinha As Long
    Dim valorCNPJ As String
    Dim cnpjComMascara As String
    Dim planilhas As Variant
    Dim i As Integer

    ' Definir as planilhas onde a máscara será aplicada
    planilhas = Array("Comp-Saidas", "Comp-Entradas", "Comp-CFe", "NNLs-Saidas", "NNLs-CFe")
    
    ' Percorrer cada planilha da lista
    For i = LBound(planilhas) To UBound(planilhas)
        Set ws = ThisWorkbook.Sheets(planilhas(i))
        
        ' Verificar se a célula C2 não está vazia antes de aplicar a máscara
        If ws.Range("C2").Value <> "" Then
            ' Determinar a última linha da coluna C (a partir da linha 2)
            ultimaLinha = ws.Cells(ws.Rows.Count, "C").End(xlUp).Row
            
            ' Aplicar a máscara para cada célula da coluna C, a partir da linha 2
            For Each celula In ws.Range("C2:C" & ultimaLinha)
                valorCNPJ = CStr(celula.Value)
                
                ' Completar com zeros à esquerda para garantir 14 dígitos
                valorCNPJ = Right("00000000000000" & valorCNPJ, 14)
                
                ' Aplicar a máscara de CNPJ (XX.XXX.XXX/XXXX-XX)
                cnpjComMascara = Left(valorCNPJ, 2) & "." & Mid(valorCNPJ, 3, 3) & "." & Mid(valorCNPJ, 6, 3) & "/" & Mid(valorCNPJ, 9, 4) & "-" & Right(valorCNPJ, 2)
                
                ' Atualizar a célula com o CNPJ formatado
                celula.Value = cnpjComMascara
            Next celula
        End If
    Next i
    
    ZFormatarCS
    
End Sub



Private Sub ZFormatarCS()

    ' Desativa atualizações de tela
    Application.ScreenUpdating = False

    Dim wsContabilizacao As Worksheet
    Dim lastRowContabilizacao As Long
    Dim lastRowContabilizacaoM As Long
    Dim i As Long
    
    Set wsContabilizacao = ThisWorkbook.Sheets("Cont-Saidas")
    
   ' Ativar a planilha "Comp-Saidas"
    wsContabilizacao.Activate
    
    
    ' Encontra a última linha na coluna A de "Contabilização"
    lastRowContabilizacao = wsContabilizacao.Cells(wsContabilizacao.Rows.Count, "A").End(xlUp).Row - 2
    lastRowContabilizacaoM = lastRowContabilizacao + 2

    ' Mesclando conjunto de 3 células horizontalmente
    wsContabilizacao.Range("A1:C1").Merge
    wsContabilizacao.Range("D1:E1").Merge
    wsContabilizacao.Range("F1:I1").Merge
    wsContabilizacao.Range("J1:L1").Merge
    
    ' Negrito nas duas primeiras linhas
    With wsContabilizacao.Rows("1:2").Font
        .Bold = True
    End With
    
    ' Inserindo formato de moeda nas células
    wsContabilizacao.Range("J3:L" & lastRowContabilizacaoM).NumberFormat = _
        "_-[$R$-pt-BR] * #,##0.00_-;-[$R$-pt-BR] * #,##0.00_-;_-[$R$-pt-BR] * ""-""??_-;_-@_-"
    
    
    ' Inserindo bordas finas por toda a tabela
    With wsContabilizacao.Range("A1:L" & lastRowContabilizacaoM).Borders
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    
    ' Inserindo bordas grossas nas extremidades verticais e horizontais
    Dim ranges As Variant
    ranges = Array("A3:C", "D3:E", "F3:I", "J3:L")
    
    For Each rng In ranges
        With wsContabilizacao.Range(rng & "3:" & rng & lastRowContabilizacaoM).Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlMedium
        End With
        With wsContabilizacao.Range(rng & "3:" & rng & lastRowContabilizacaoM).Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlMedium
        End With
        With wsContabilizacao.Range(rng & "3:" & rng & lastRowContabilizacaoM).Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlMedium
        End With
        With wsContabilizacao.Range(rng & "3:" & rng & lastRowContabilizacaoM).Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlMedium
        End With
    Next rng

    ' Inserindo bordas grossas nas extremidades verticais e horizontais do cabeçalho
    ranges = Array("A1:C2", "D1:E2", "F1:I2", "J1:L2")
    
    For Each rng In ranges
        With wsContabilizacao.Range(rng).Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlMedium
        End With
        With wsContabilizacao.Range(rng).Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlMedium
        End With
        With wsContabilizacao.Range(rng).Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlMedium
        End With
        With wsContabilizacao.Range(rng).Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlMedium
        End With
    Next rng
    
    
     ' Centraliza o texto nas linhas 1 e 2
    With wsContabilizacao.Rows("1:2")
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
   


    
    ' Inserindo cores do cabeçalho
    With wsContabilizacao.Range("F1:I2").Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = -0.249977111117893
        .PatternTintAndShade = 0
    End With
    
    With wsContabilizacao.Range("D1:E2").Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = 0.399975585192419
        .PatternTintAndShade = 0
    End With
    

    ' Aplica o preenchimento sólido com a cor tema "Accent6" no intervalo A1:C2
    With Range("A1:C2").Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent2
        .TintAndShade = 0.399975585192419
        .PatternTintAndShade = 0
    End With

    ' Aplica o preenchimento sólido com a cor tema "Accent6" e sombreamento no intervalo J1:L2
    With Range("J1:L2").Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent4
        .TintAndShade = 0.399975585192419
        .PatternTintAndShade = 0
    End With



    ' Loop para formatar linhas alternadamente
    For i = 3 To lastRowContabilizacaoM
        ' Se i for ímpar, aplica a primeira cor, senão, aplica a segunda cor
        If i Mod 2 = 1 Then
            wsContabilizacao.Range("A" & i & ":L" & i).Interior.Color = RGB(255, 255, 255)
        Else
            wsContabilizacao.Range("A" & i & ":L" & i).Interior.Color = RGB(220, 220, 220)
        End If
    Next i

    ' Ajustar a largura das colunas para melhor visualização
    wsContabilizacao.Columns("A:I").AutoFit
    wsContabilizacao.Columns("F:I").ColumnWidth = 14.8
    
    ' Ativa atualizações de tela
    Application.ScreenUpdating = True
        
    FormatCompSaidas
    
End Sub



Private Sub FormatCompSaidas()

    Dim wsNotasFaltantes As Worksheet
    Dim lastRow As Long
    Dim i As Long

    ' Desativa atualizações de tela
    Application.ScreenUpdating = False

    ' Defina a planilha de trabalho
    Set wsNotasFaltantes = ThisWorkbook.Sheets("Comp-Saidas")

   ' Ativar a planilha "Comp-Saidas"
    wsNotasFaltantes.Activate


    'Inserindo formato de moeda nas células
    Columns("H:I").Select
    Selection.NumberFormat = _
        "_-[$R$-pt-BR] * #,##0.00_-;-[$R$-pt-BR] * #,##0.00_-;_-[$R$-pt-BR] * ""-""??_-;_-@_-"

    ' Encontra a última linha na coluna A
    lastRow = wsNotasFaltantes.Cells(wsNotasFaltantes.Rows.Count, "A").End(xlUp).Row

    ' Negrito na primeira linha
    wsNotasFaltantes.Rows("1:1").Font.Bold = True

    ' Inserindo bordas normal por toda a tabela
    With wsNotasFaltantes.Range("A1:J" & lastRow).Borders
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With

    ' Aplicar bordas a cada conjunto de colunas
    AplicarBordas wsNotasFaltantes.Range("A1:C1")
    AplicarBordas wsNotasFaltantes.Range("A2:C" & lastRow)
    AplicarBordas wsNotasFaltantes.Range("D1:F1")
    AplicarBordas wsNotasFaltantes.Range("D2:F" & lastRow)
    AplicarBordas wsNotasFaltantes.Range("G1:J1")
    AplicarBordas wsNotasFaltantes.Range("G2:J" & lastRow)

    
    
    ' Aplicar cores aos fundos das células
    wsNotasFaltantes.Range("A1:C1").Interior.Color = RGB(0, 255, 0)     ' Verde
    wsNotasFaltantes.Range("D1:F1").Interior.Color = RGB(18, 154, 238)  ' Azul
    wsNotasFaltantes.Range("G1:J1").Interior.Color = RGB(231, 171, 49)  ' Marrom

    ' Aplicar centralização
    wsNotasFaltantes.Range("H:I").HorizontalAlignment = xlRight
    wsNotasFaltantes.Range("A1:J1").HorizontalAlignment = xlCenter

    ' Aplicar cores alternadas
    For i = 2 To lastRow
        If i Mod 2 = 1 Then
            wsNotasFaltantes.Range("A" & i & ":J" & i).Interior.Color = RGB(255, 255, 255)
        Else
            wsNotasFaltantes.Range("A" & i & ":J" & i).Interior.Color = RGB(220, 220, 220)
        End If
    Next i

    ' Ajustar a largura das colunas para melhor visualização
    wsNotasFaltantes.Columns("A:J").AutoFit
    wsNotasFaltantes.Columns("H:I").ColumnWidth = 14.8
    
    FormatarContEntradas

End Sub


Private Sub FormatarContEntradas()

    ' Desativa atualizações de tela
    Application.ScreenUpdating = False

    Dim wsContabilizacao As Worksheet
    Dim lastRowContabilizacao As Long
    Dim lastRowContabilizacaoM As Long
    Dim i As Long
    
    Set wsContabilizacao = ThisWorkbook.Sheets("Cont-Entradas")
    
   ' Ativar a planilha "Comp-Entradas"
    wsContabilizacao.Activate
    
    
    ' Encontra a última linha na coluna A de "Contabilização"
    lastRowContabilizacao = wsContabilizacao.Cells(wsContabilizacao.Rows.Count, "A").End(xlUp).Row - 2
    lastRowContabilizacaoM = lastRowContabilizacao + 2

    ' Mesclando conjunto de 3 células horizontalmente
    wsContabilizacao.Range("A1:C1").Merge
    wsContabilizacao.Range("D1:E1").Merge
    wsContabilizacao.Range("F1:I1").Merge
    wsContabilizacao.Range("J1:L1").Merge
    
    ' Negrito nas duas primeiras linhas
    With wsContabilizacao.Rows("1:2").Font
        .Bold = True
    End With
    
    ' Inserindo formato de moeda nas células
    wsContabilizacao.Range("J3:L" & lastRowContabilizacaoM).NumberFormat = _
        "_-[$R$-pt-BR] * #,##0.00_-;-[$R$-pt-BR] * #,##0.00_-;_-[$R$-pt-BR] * ""-""??_-;_-@_-"
    
    
    ' Inserindo bordas finas por toda a tabela
    With wsContabilizacao.Range("A1:L" & lastRowContabilizacaoM).Borders
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    
    ' Inserindo bordas grossas nas extremidades verticais e horizontais
    Dim ranges As Variant
    ranges = Array("A3:C", "D3:E", "F3:I", "J3:L")
    
    For Each rng In ranges
        With wsContabilizacao.Range(rng & "3:" & rng & lastRowContabilizacaoM).Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlMedium
        End With
        With wsContabilizacao.Range(rng & "3:" & rng & lastRowContabilizacaoM).Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlMedium
        End With
        With wsContabilizacao.Range(rng & "3:" & rng & lastRowContabilizacaoM).Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlMedium
        End With
        With wsContabilizacao.Range(rng & "3:" & rng & lastRowContabilizacaoM).Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlMedium
        End With
    Next rng

    ' Inserindo bordas grossas nas extremidades verticais e horizontais do cabeçalho
    ranges = Array("A1:C2", "D1:E2", "F1:I2", "J1:L2")
    
    For Each rng In ranges
        With wsContabilizacao.Range(rng).Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlMedium
        End With
        With wsContabilizacao.Range(rng).Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlMedium
        End With
        With wsContabilizacao.Range(rng).Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlMedium
        End With
        With wsContabilizacao.Range(rng).Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlMedium
        End With
    Next rng
    
    
     ' Centraliza o texto nas linhas 1 e 2
    With wsContabilizacao.Rows("1:2")
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
   


    
    ' Inserindo cores do cabeçalho
    With wsContabilizacao.Range("F1:I2").Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = -0.249977111117893
        .PatternTintAndShade = 0
    End With
    
    With wsContabilizacao.Range("D1:E2").Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = 0.399975585192419
        .PatternTintAndShade = 0
    End With
    

    ' Aplica o preenchimento sólido com a cor tema "Accent6" no intervalo A1:C2
    With Range("A1:C2").Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent2
        .TintAndShade = 0.399975585192419
        .PatternTintAndShade = 0
    End With

    ' Aplica o preenchimento sólido com a cor tema "Accent6" e sombreamento no intervalo J1:L2
    With Range("J1:L2").Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent4
        .TintAndShade = 0.399975585192419
        .PatternTintAndShade = 0
    End With



    ' Loop para formatar linhas alternadamente
    For i = 3 To lastRowContabilizacaoM
        ' Se i for ímpar, aplica a primeira cor, senão, aplica a segunda cor
        If i Mod 2 = 1 Then
            wsContabilizacao.Range("A" & i & ":L" & i).Interior.Color = RGB(255, 255, 255)
        Else
            wsContabilizacao.Range("A" & i & ":L" & i).Interior.Color = RGB(220, 220, 220)
        End If
    Next i

    ' Ajustar a largura das colunas para melhor visualização
    wsContabilizacao.Columns("A:I").AutoFit
    wsContabilizacao.Columns("F:I").ColumnWidth = 14.8
    
    ' Ativa atualizações de tela
    Application.ScreenUpdating = True
        
    FormatCompEntradas
    
End Sub



Private Sub FormatCompEntradas()

    Dim wsNotasFaltantes As Worksheet
    Dim lastRow As Long
    Dim i As Long

    ' Desativa atualizações de tela
    Application.ScreenUpdating = False

    ' Defina a planilha de trabalho
    Set wsNotasFaltantes = ThisWorkbook.Sheets("Comp-Entradas")

   ' Ativar a planilha "Comp-Entradas"
    wsNotasFaltantes.Activate


    'Inserindo formato de moeda nas células
    Columns("H:I").Select
    Selection.NumberFormat = _
        "_-[$R$-pt-BR] * #,##0.00_-;-[$R$-pt-BR] * #,##0.00_-;_-[$R$-pt-BR] * ""-""??_-;_-@_-"

    ' Encontra a última linha na coluna A
    lastRow = wsNotasFaltantes.Cells(wsNotasFaltantes.Rows.Count, "A").End(xlUp).Row

    ' Negrito na primeira linha
    wsNotasFaltantes.Rows("1:1").Font.Bold = True

    ' Inserindo bordas normal por toda a tabela
    With wsNotasFaltantes.Range("A1:J" & lastRow).Borders
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With

    ' Aplicar bordas a cada conjunto de colunas
    AplicarBordas wsNotasFaltantes.Range("A1:C1")
    AplicarBordas wsNotasFaltantes.Range("A2:C" & lastRow)
    AplicarBordas wsNotasFaltantes.Range("D1:F1")
    AplicarBordas wsNotasFaltantes.Range("D2:F" & lastRow)
    AplicarBordas wsNotasFaltantes.Range("G1:J1")
    AplicarBordas wsNotasFaltantes.Range("G2:J" & lastRow)

    
    
    ' Aplicar cores aos fundos das células
    wsNotasFaltantes.Range("A1:C1").Interior.Color = RGB(0, 255, 0)     ' Verde
    wsNotasFaltantes.Range("D1:F1").Interior.Color = RGB(18, 154, 238)  ' Azul
    wsNotasFaltantes.Range("G1:J1").Interior.Color = RGB(231, 171, 49)  ' Marrom

    ' Aplicar centralização
    wsNotasFaltantes.Range("H:I").HorizontalAlignment = xlRight
    wsNotasFaltantes.Range("A1:J1").HorizontalAlignment = xlCenter

    ' Aplicar cores alternadas
    For i = 2 To lastRow
        If i Mod 2 = 1 Then
            wsNotasFaltantes.Range("A" & i & ":J" & i).Interior.Color = RGB(255, 255, 255)
        Else
            wsNotasFaltantes.Range("A" & i & ":J" & i).Interior.Color = RGB(220, 220, 220)
        End If
    Next i

    ' Ajustar a largura das colunas para melhor visualização
    wsNotasFaltantes.Columns("A:J").AutoFit
    wsNotasFaltantes.Columns("H:I").ColumnWidth = 14.8

    FormatarContCFe

End Sub






Private Sub FormatarContCFe()

    ' Desativa atualizações de tela
    Application.ScreenUpdating = False

    Dim wsContabilizacao As Worksheet
    Dim lastRowContabilizacao As Long
    Dim lastRowContabilizacaoM As Long
    Dim i As Long
    
    Set wsContabilizacao = ThisWorkbook.Sheets("Cont-CFe")
    
   ' Ativar a planilha "Comp-CFe"
    wsContabilizacao.Activate
    
    
    ' Encontra a última linha na coluna A de "Contabilização"
    lastRowContabilizacao = wsContabilizacao.Cells(wsContabilizacao.Rows.Count, "A").End(xlUp).Row - 2
    lastRowContabilizacaoM = lastRowContabilizacao + 2

    ' Mesclando conjunto de 3 células horizontalmente
    wsContabilizacao.Range("A1:C1").Merge
    wsContabilizacao.Range("D1:E1").Merge
    wsContabilizacao.Range("F1:I1").Merge
    wsContabilizacao.Range("J1:L1").Merge
    
    ' Negrito nas duas primeiras linhas
    With wsContabilizacao.Rows("1:2").Font
        .Bold = True
    End With
    
    ' Inserindo formato de moeda nas células
    wsContabilizacao.Range("J3:L" & lastRowContabilizacaoM).NumberFormat = _
        "_-[$R$-pt-BR] * #,##0.00_-;-[$R$-pt-BR] * #,##0.00_-;_-[$R$-pt-BR] * ""-""??_-;_-@_-"
    
    
    ' Inserindo bordas finas por toda a tabela
    With wsContabilizacao.Range("A1:L" & lastRowContabilizacaoM).Borders
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    
    ' Inserindo bordas grossas nas extremidades verticais e horizontais
    Dim ranges As Variant
    ranges = Array("A3:C", "D3:E", "F3:I", "J3:L")
    
    For Each rng In ranges
        With wsContabilizacao.Range(rng & "3:" & rng & lastRowContabilizacaoM).Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlMedium
        End With
        With wsContabilizacao.Range(rng & "3:" & rng & lastRowContabilizacaoM).Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlMedium
        End With
        With wsContabilizacao.Range(rng & "3:" & rng & lastRowContabilizacaoM).Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlMedium
        End With
        With wsContabilizacao.Range(rng & "3:" & rng & lastRowContabilizacaoM).Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlMedium
        End With
    Next rng

    ' Inserindo bordas grossas nas extremidades verticais e horizontais do cabeçalho
    ranges = Array("A1:C2", "D1:E2", "F1:I2", "J1:L2")
    
    For Each rng In ranges
        With wsContabilizacao.Range(rng).Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlMedium
        End With
        With wsContabilizacao.Range(rng).Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlMedium
        End With
        With wsContabilizacao.Range(rng).Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlMedium
        End With
        With wsContabilizacao.Range(rng).Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlMedium
        End With
    Next rng
    
    
     ' Centraliza o texto nas linhas 1 e 2
    With wsContabilizacao.Rows("1:2")
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
   


    
    ' Inserindo cores do cabeçalho
    With wsContabilizacao.Range("F1:I2").Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = -0.249977111117893
        .PatternTintAndShade = 0
    End With
    
    With wsContabilizacao.Range("D1:E2").Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = 0.399975585192419
        .PatternTintAndShade = 0
    End With
    

    ' Aplica o preenchimento sólido com a cor tema "Accent6" no intervalo A1:C2
    With Range("A1:C2").Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent2
        .TintAndShade = 0.399975585192419
        .PatternTintAndShade = 0
    End With

    ' Aplica o preenchimento sólido com a cor tema "Accent6" e sombreamento no intervalo J1:L2
    With Range("J1:L2").Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent4
        .TintAndShade = 0.399975585192419
        .PatternTintAndShade = 0
    End With



    ' Loop para formatar linhas alternadamente
    For i = 3 To lastRowContabilizacaoM
        ' Se i for ímpar, aplica a primeira cor, senão, aplica a segunda cor
        If i Mod 2 = 1 Then
            wsContabilizacao.Range("A" & i & ":L" & i).Interior.Color = RGB(255, 255, 255)
        Else
            wsContabilizacao.Range("A" & i & ":L" & i).Interior.Color = RGB(220, 220, 220)
        End If
    Next i

    ' Ajustar a largura das colunas para melhor visualização
    wsContabilizacao.Columns("A:I").AutoFit
    wsContabilizacao.Columns("F:I").ColumnWidth = 14.8
    
    ' Ativa atualizações de tela
    Application.ScreenUpdating = True
        
    FormatCompCFe
    
End Sub






Private Sub FormatCompCFe()

    Dim wsNotasFaltantes As Worksheet
    Dim lastRow As Long
    Dim i As Long

    ' Desativa atualizações de tela
    Application.ScreenUpdating = False

    ' Defina a planilha de trabalho
    Set wsNotasFaltantes = ThisWorkbook.Sheets("Comp-CFe")

   ' Ativar a planilha "Comp-CFe"
    wsNotasFaltantes.Activate


    'Inserindo formato de moeda nas células
    Columns("H:I").Select
    Selection.NumberFormat = _
        "_-[$R$-pt-BR] * #,##0.00_-;-[$R$-pt-BR] * #,##0.00_-;_-[$R$-pt-BR] * ""-""??_-;_-@_-"

    ' Encontra a última linha na coluna A
    lastRow = wsNotasFaltantes.Cells(wsNotasFaltantes.Rows.Count, "A").End(xlUp).Row

    ' Negrito na primeira linha
    wsNotasFaltantes.Rows("1:1").Font.Bold = True

    ' Inserindo bordas normal por toda a tabela
    With wsNotasFaltantes.Range("A1:J" & lastRow).Borders
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With

    ' Aplicar bordas a cada conjunto de colunas
    AplicarBordas wsNotasFaltantes.Range("A1:C1")
    AplicarBordas wsNotasFaltantes.Range("A2:C" & lastRow)
    AplicarBordas wsNotasFaltantes.Range("D1:F1")
    AplicarBordas wsNotasFaltantes.Range("D2:F" & lastRow)
    AplicarBordas wsNotasFaltantes.Range("G1:J1")
    AplicarBordas wsNotasFaltantes.Range("G2:J" & lastRow)

    
    
    ' Aplicar cores aos fundos das células
    wsNotasFaltantes.Range("A1:C1").Interior.Color = RGB(0, 255, 0)     ' Verde
    wsNotasFaltantes.Range("D1:F1").Interior.Color = RGB(18, 154, 238)  ' Azul
    wsNotasFaltantes.Range("G1:J1").Interior.Color = RGB(231, 171, 49)  ' Marrom

    ' Aplicar centralização
    wsNotasFaltantes.Range("H:I").HorizontalAlignment = xlRight
    wsNotasFaltantes.Range("A1:J1").HorizontalAlignment = xlCenter

    ' Aplicar cores alternadas
    For i = 2 To lastRow
        If i Mod 2 = 1 Then
            wsNotasFaltantes.Range("A" & i & ":J" & i).Interior.Color = RGB(255, 255, 255)
        Else
            wsNotasFaltantes.Range("A" & i & ":J" & i).Interior.Color = RGB(220, 220, 220)
        End If
    Next i

    ' Ajustar a largura das colunas para melhor visualização
    wsNotasFaltantes.Columns("A:J").AutoFit
    wsNotasFaltantes.Columns("H:I").ColumnWidth = 14.8

    FormatSeqSai
    
End Sub



Private Sub FormatSeqSai()

    Dim wsNotasFaltantes As Worksheet
    Dim lastRow As Long
    Dim i As Long

    ' Desativa atualizações de tela
    Application.ScreenUpdating = False

    ' Defina a planilha de trabalho
    Set wsNotasFaltantes = ThisWorkbook.Sheets("NNLs-Saidas")

   ' Ativar a planilha
    wsNotasFaltantes.Activate



    ' Encontra a última linha na coluna A
    lastRow = wsNotasFaltantes.Cells(wsNotasFaltantes.Rows.Count, "A").End(xlUp).Row

    ' Negrito na primeira linha
    wsNotasFaltantes.Rows("1:1").Font.Bold = True

    ' Inserindo bordas normal por toda a tabela
    With wsNotasFaltantes.Range("A1:D" & lastRow).Borders
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With

    ' Aplicar bordas a cada conjunto de colunas
    AplicarBordas wsNotasFaltantes.Range("A1:D1")
    AplicarBordas wsNotasFaltantes.Range("A2:D" & lastRow)
    
    
    ' Aplicar cores aos fundos das células
    wsNotasFaltantes.Range("A1:D1").Interior.Color = RGB(0, 255, 0)     ' Verde

    ' Aplicar centralização
    wsNotasFaltantes.Range("D:D").HorizontalAlignment = xlCenter

    ' Aplicar cores alternadas
    For i = 2 To lastRow
        If i Mod 2 = 1 Then
            wsNotasFaltantes.Range("A" & i & ":D" & i).Interior.Color = RGB(255, 255, 255)
        Else
            wsNotasFaltantes.Range("A" & i & ":D" & i).Interior.Color = RGB(220, 220, 220)
        End If
    Next i

    ' Ajustar a largura das colunas para melhor visualização
    wsNotasFaltantes.Columns("A:D").AutoFit

    FormatSeqCFs

End Sub

Private Sub FormatSeqCFs()

    Dim wsNotasFaltantes As Worksheet
    Dim lastRow As Long
    Dim i As Long

    ' Desativa atualizações de tela
    Application.ScreenUpdating = False

    ' Defina a planilha de trabalho
    Set wsNotasFaltantes = ThisWorkbook.Sheets("NNLs-CFe")

   ' Ativar a planilha
    wsNotasFaltantes.Activate



    ' Encontra a última linha na coluna A
    lastRow = wsNotasFaltantes.Cells(wsNotasFaltantes.Rows.Count, "A").End(xlUp).Row

    ' Negrito na primeira linha
    wsNotasFaltantes.Rows("1:1").Font.Bold = True

    ' Inserindo bordas normal por toda a tabela
    With wsNotasFaltantes.Range("A1:D" & lastRow).Borders
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With

    ' Aplicar bordas a cada conjunto de colunas
    AplicarBordas wsNotasFaltantes.Range("A1:D1")
    AplicarBordas wsNotasFaltantes.Range("A2:D" & lastRow)
    
    
    ' Aplicar cores aos fundos das células
    wsNotasFaltantes.Range("A1:D1").Interior.Color = RGB(0, 255, 0)     ' Verde

    ' Aplicar centralização
    wsNotasFaltantes.Range("D:D").HorizontalAlignment = xlCenter

    ' Aplicar cores alternadas
    For i = 2 To lastRow
        If i Mod 2 = 1 Then
            wsNotasFaltantes.Range("A" & i & ":D" & i).Interior.Color = RGB(255, 255, 255)
        Else
            wsNotasFaltantes.Range("A" & i & ":D" & i).Interior.Color = RGB(220, 220, 220)
        End If
    Next i

    ' Ajustar a largura das colunas para melhor visualização
    wsNotasFaltantes.Columns("A:D").AutoFit

    ApagarLinhasSeA2Vazio

End Sub



    Sub AplicarBordas(rng As Range)
        With rng.Borders(xlDiagonalDown)
            .LineStyle = xlNone
        End With

        With rng.Borders(xlDiagonalUp)
            .LineStyle = xlNone
        End With

        With rng.Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlMedium
        End With

        With rng.Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlMedium
        End With

        With rng.Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlMedium
        End With

        With rng.Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlMedium
        End With
    End Sub



Private Sub ApagarLinhasSeA2Vazio()
    Dim ws As Worksheet
    Dim planilhas As Variant
    Dim i As Integer

    ' Definir as planilhas onde a verificação será realizada
    planilhas = Array("Comp-Saidas", "Comp-Entradas", "Comp-CFe", "NNLs-Saidas", "NNLs-CFe")
    
    ' Percorrer cada planilha da lista
    For i = LBound(planilhas) To UBound(planilhas)
        Set ws = ThisWorkbook.Sheets(planilhas(i))
        
        ' Verificar se a célula A2 está vazia
        If ws.Range("A2").Value = "" Then
            ' Apagar a linha 2
            ws.Rows(2).Delete
        End If
    Next i
    
    ZZSalvar.ZZSalvarAbas
    
End Sub



