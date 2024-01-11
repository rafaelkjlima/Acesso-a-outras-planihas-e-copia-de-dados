Sub Reposições()

    Dim wb1 As Workbook
    Dim wb2 As Workbook
    Dim planilha1 As Worksheet
    Dim planilha2 As Worksheet
    Dim ultimaLinha1 As Long, ultimaLinha2 As Long
    Dim i As Long, j As Long
    Dim encontrado As Boolean
    
    ' Abra as pastas de trabalho
    Set wb1 = Workbooks.Open("Caminho\Para\Workbook1.xlsx") 
		'referente a pasta de origem
    Set wb2 = Workbooks.Open("Caminho\Para\Workbook2.xlsx") 
    		'referente a pasta de destino

    ' Defina as planilhas onde os dados estão em cada pasta de trabalho
    Set planilha1 = wb1.Sheets("Planilha1")
		 'Planilha de origem
    Set planilha2 = wb2.Sheets("Planilha2")
		 'Planilha de destino
    
    ' Encontre a última linha com dados na T2, P2, coluna B em Origem
    ultimaLinha2 = planilha2.Cells(planilha2.Rows.Count, "B").End(xlUp).Row
    
    ' Encontre a última linha com dados na T1, P1, coluna A em Destino
    ultimaLinha1 = planilha1.Cells(planilha1.Rows.Count, "A").End(xlUp).Row
    
    ' Loop através das células na coluna B da T2, P2 em Origem
    For i = 1 To ultimaLinha2
        ' Verifique se o valor na célula é igual 0
        If planilha2.Cells(i, "B").Value = 0 Then

            ' Verifica se os dados já existem na T1, P1 em Destino
            encontrado = False
            For j = 1 To ultimaLinha1
                If planilha1.Cells(j, "A").Value = planilha2.Cells(i, "A").Value And _
                   planilha1.Cells(j, "B").Value = planilha2.Cells(i, "B").Value And _
                   planilha1.Cells(j, "C").Value = planilha2.Cells(i, "C").Value Then
                    encontrado = True
                    Exit For
                End If
            Next j
            
            ' Se não foram encontrados, copie os dados para a T1, P1 em Origem
            If Not encontrado Then
                ultimaLinha1 = ultimaLinha1 + 1 
			' Avança para a próxima linha na T1 em Origem e cola os dados
                For j = 1 To planilha2.Cells(i, planilha2.Columns.Count).End(xlToLeft).Column
                    planilha2.Cells(i, j).Copy Destination:=planilha1.Cells(ultimaLinha1, j)
                Next j
            End If
        End If
    Next i
    
    wb1.Close SaveChanges:=False
    wb2 SaveChanges:=True
End Sub
