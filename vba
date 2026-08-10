Sub formatar_razao_s_agente()
    Dim resposta As VbMsgBoxResult
    Dim cel As Range
    Dim ws As Worksheet
    Dim lastRow As Long

    Set ws = ActiveSheet

    resposta = MsgBox("Datas filtradas?", vbQuestion + vbYesNo, "Confirmação")

    If resposta <> vbYes Then
        MsgBox "Lembrar de filtrar as datas na coluna B.", vbInformation, "Cancelado"
        Exit Sub
    End If

    ' Insere uma nova coluna à direita da coluna B (coluna C)
    ws.Columns("C:C").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove

    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row

    ' Copia os valores visíveis da coluna B para a coluna C na MESMA linha
    On Error Resume Next ' Caso não tenha células visíveis
    For Each cel In ws.Range("B2:B" & lastRow).SpecialCells(xlCellTypeVisible)
        cel.Offset(0, 1).Value = cel.Value ' Coluna C na mesma linha
    Next cel
    On Error GoTo 0
    
    'Limpar os filtros
    If ws.AutoFilterMode Then ws.AutoFilterMode = False

    'Deletar a célula B2
   ws.Columns("B").Rows("2").Delete Shift:=xlShiftUp
   
    ' Remove filtro anterior, se houver
    If ws.AutoFilterMode Then ws.AutoFilterMode = False
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    ' Aplica filtro da coluna A com "Complemento:"
    ws.Range("A1:I" & lastRow).AutoFilter Field:=1, Criteria1:="Complemento:"

    ' Deleta linhas filtradas (com todas as colunas)
    On Error Resume Next
    With ws.AutoFilter.Range
    .Offset(1, 0).Resize(.Rows.Count - 1).SpecialCells(xlCellTypeVisible).EntireRow.Delete
    End With
    On Error GoTo 0

    ' Remove o filtro
    If ws.AutoFilterMode Then ws.AutoFilterMode = False
    
    'Cabeçalhos
    With ws
    .Range("B1").Value = "Histórico"
    .Range("C1").Value = "Data"
    .Range("D1").Value = "Documento"
    .Range("E1").Value = "Contrapartida"
    .Range("F1").Value = "Lote"
    .Range("G1").Value = "Débito"
    .Range("H1").Value = "Crédito"
    .Range("I1").Value = "Saldo"
    End With
    
If ws.AutoFilterMode Then ws.AutoFilterMode = False
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    ws.Range("A1", ws.Cells(lastRow, ws.Columns.Count).End(xlToLeft)).AutoFilter Field:=1, Criteria1:="Conta:*", Operator:=xlAnd

    Dim rngFiltro As Range
    On Error Resume Next
    Set rngFiltro = ws.Range("A2:A" & lastRow).SpecialCells(xlCellTypeVisible)
    On Error GoTo 0

    If Not rngFiltro Is Nothing Then
        For Each cel In rngFiltro
            cel.EntireRow.Interior.Color = RGB(208, 208, 208)
            ws.Cells(cel.Row, "B").ClearContents
            ws.Cells(cel.Row, "D").ClearContents
        Next cel
    End If

    If ws.AutoFilterMode Then ws.AutoFilterMode = False
    
Rows(1).Interior.Color = RGB(208, 208, 208)

ws.Columns("E:I").AutoFit
ws.Range("A1:I1").AutoFilter
   
End Sub

