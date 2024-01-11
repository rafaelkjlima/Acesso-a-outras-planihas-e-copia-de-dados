Sub Copiarreposicao()
    Dim wbOrigem As Workbook
    Dim wbDestino As Workbook
    Dim planilhaOrigem As Worksheet
    Dim planilhaDestino As Worksheet
    Dim ultimaLinhaOrigem As Long, ultimaLinhaDestino As Long
    Dim i As Long, j As Long
    Dim encontrado As Boolean
    
    ' Abre as pastas de trabalho
    Set wbOrigem = Workbooks.Open("\\servidormicrolins\H\COORDENAÇÃO\TABELA DE TURMAS\TABELA DE TURMA INTERATIVO.xlsm")
            ' Caminho Origem
            
    Set wbDestino = Workbooks.Open("\\servidormicrolins\F\DINAMICA\LISTAGEM REPOSIÇÕES\CONTROLE DE REPOSIÇÃO - VBA V.1.xlsm")
            ' Caminho Destino
            
    ' Planilhas definidas
    Set planilhaOrigem = wbOrigem.Sheets(5)
    Set planilhaDestino = wbDestino.Sheets(1)
    
    ultimaLinhaOrigem = planilhaOrigem.Cells(planilhaOrigem.Rows.Count, "D").End(xlUp).Row
    
    ultimaLinhaDestino = planilhaDestino.Cells(planilhaDestino.Rows.Count, "A").End(xlUp).Row + 1
    
    ' Loop através das células na coluna Reposições da Planilha Origem
    For i = 1 To ultimaLinhaOrigem
        ' Verifique se o valor na célula é igual a zero
        If planilhaOrigem.Cells(i, "D").Value = "Reposição" Then
            ' Verifica se os dados já existem na Planilha Destino
            encontrado = False
            For j = 1 To ultimaLinhaDestino
                If planilhaDestino.Cells(j, "A").Value = planilhaOrigem.Cells(i, "A").Value And _
                   planilhaDestino.Cells(j, "B").Value = planilhaOrigem.Cells(i, "B").Value And _
                   planilhaDestino.Cells(j, "C").Value = planilhaOrigem.Cells(i, "C").Value And _
                   planilhaDestino.Cells(j, "D").Value = planilhaOrigem.Cells(i, "D").Value Then
                    encontrado = True
                    Exit For
                End If
            Next j
            
            ' Se correto, copie os dados para a Planilha Destino
            If Not encontrado Then
                For j = 1 To planilhaOrigem.Cells(i, planilhaOrigem.Columns.Count).End(xlToLeft).Column
                    planilhaOrigem.Cells(i, j).Copy Destination:=planilhaDestino.Cells(ultimaLinhaDestino, j)
                Next j
                 planilhaDestino.Cells(ultimaLinhaDestino, "J").Value = Date
                
                ultimaLinhaDestino = ultimaLinhaDestino + 1
            End If
        End If
    Next i
     wbOrigem.Close SaveChanges:=False
     
End Sub
