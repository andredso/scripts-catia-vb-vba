Sub Exporta_Controle_Estoque()
	Dim WS_1 As Worksheet: Set WS_1 = Worksheets("EstoqueAtual-e-Cadastro")
	Dim strNovoArq1$	
	Dim intNumArq1%: intNumArq1 = 1		
	Dim strCol$: strCol = "A"
	Dim lngRow&: lngRow = 6
	Dim varDadoCell1 As Variant	
	
	strNovoArq1 = "C:\Controle de Estoque - NEO BRASIL.txt" 'Arquivo com materiais novos	
	Open strNovoArq1 For Output As #intNumArq1 'Exporta os dados da planilha para o arquivo criado	
	
	With WS_1
		Do While (.Range(strCol & CStr(lngRow)).Value <> "")			
			varDadoCell1 = .Range("A" & CStr(lngRow)).Value & ";" & .Range("B" & CStr(lngRow)).Value & ";" & .Range("C" & CStr(lngRow)).Value & ";" & .Range("G" & CStr(lngRow)).Value & ";" & .Range("H" & CStr(lngRow)).Value
			lngRow = lngRow + 1			
			if varDadoCell1 <> "" then Print #intNumArq1, varDadoCell1			
			varDadoCell1 = ""			
		Loop
	End With
	Close #intNumArq1 'Salva e fecha o arquivo de texto com os dados	
	Set WS_1 = Nothing
End Sub
'--------------------------------------------------------------------------------------------------------------------------------------------------------------'
Sub Exporta_novos_materiais()
	Dim WS_1 As Worksheet: Set WS_1 = Worksheets("cont_est")
    Dim strNovoArq1$
    Dim intNumArq1%: intNumArq1 = 1
    Dim intReg1%: intReg1 = 408
    Dim strCol$: strCol = "A"
    Dim lngRow&: lngRow = 1
    Dim varDadoCell1 As Variant
    
    strNovoArq1 = "C:\materiais_novos.txt" 'Arquivo com materiais novos
    Open strNovoArq1 For Output As #intNumArq1 'Exporta os dados da planilha para o arquivo criado
    Print #intNumArq1, "INSERT INTO tb_neo_mat (mat_id_pk, mat_gtin13, mat_codneo_sku, mat_desc, mat_dat_cad, mat_dat_alt, pdt_id_fk, mar_id_fk) VALUES"
    
    WS_1.Select
    
    With WS_1
        Do While (.Range(strCol & CStr(lngRow)).Value <> "")
            .Range(strCol & CStr(lngRow)).Select
            If (.Range(strCol & CStr(lngRow + 1)).Value <> "") Then
                If (.Range("H" & CStr(lngRow)).Value = "") Then
                    varDadoCell1 = "(" & CStr(intReg1) & ", '0000000000000', '" & .Range("A" & CStr(lngRow)).Value & "', '" & .Range("B" & CStr(lngRow)).Value & "', NOW(), NOW(), " & .Range("E" & CStr(lngRow)).Value & ", 0),"
                End If
            Else
                varDadoCell1 = varDadoCell1 + ";"
            End If
            lngRow = lngRow + 1
            intReg1 = intReg1 + 1
            If varDadoCell1 <> "" Then Print #intNumArq1, varDadoCell1
            varDadoCell1 = ""
        Loop
    End With
    Close #intNumArq1 'Salva e fecha o arquivo de texto com os dados
    Set WS_1 = Nothing
End Sub
'--------------------------------------------------------------------------------------------------------------------------------------------------------------'
Sub ExportarTexto()
	Dim Ultlinha&
	Dim UltColuna&
	Dim strNovoArq1$
	Dim intNumArq1%
	Dim varDadoCell1 As Variant
	Dim Ultlinha As Long
	Dim UltColuna As Long
	Dim strNovoArq1 As String
	Dim intNumArq1 As Integer
	Dim varDadoCell1 As Variant
	
	intNumArq1 = FreeFile ' Atribui o primeiro número de arquivo disponível (E.g.: #1)
	'Determina a última linha da planilha com dados
	Ultlinha = Cells(Rows.Count, 1).End(xlUp).Row
	
	'Determina a última coluna da planilha com dados
	UltColuna = Cells(1, Columns.Count).End(xlToLeft).Column
	strNovoArq1 = "C:\Teste\Exportado.txt" 'Utilize uma pasta já existente
	
	'Exporta os dados da planilha para o arquivo criado
	Open strNovoArq1 For Output As #intNumArq1
	For i = 1 To Ultlinha
		For j = 1 To UltColuna
			If j = UltColuna Then
				varDadoCell1 = varDadoCell1 & Cells(i, j).Value
			Else
				varDadoCell1 = varDadoCell1 & Cells(i, j).Value & " "
			End If
		Next j
		Print #intNumArq1, varDadoCell1
		varDadoCell1 = ""
	Next i
	Close #intNumArq1 'Salva e fecha o arquivo de texto com os dados
End Sub
'--------------------------------------------------------------------------------------------------------------------------------------------------------------'
Sub Atualiza_Planilha()
    Dim WB_1 As Workbook: Set WB_1 = Workbooks("Controle de Estoque - NEO BRASIL.xlsx")
    Dim WS_1 As Worksheet: Set WS_1 = WB_1.Worksheets("EstoqueAtual-e-Cadastro")
    Dim WB_2 As Workbook: Set WB_2 = Workbooks("Estoque.xls")
    Dim WS_2 As Worksheet: Set WS_2 = WB_2.Worksheets("cont_est")
    Dim lngUL_WS_1&: lngUL_WS_1 = WS_1.Cells(Rows.Count, 1).End(xlUp).Row
    Dim lngUL_WS_2&: lngUL_WS_2 = WS_2.Cells(Rows.Count, 1).End(xlUp).Row
    Dim booAchou As Boolean: booAchou = False
    Dim strNovoArq1$: strNovoArq1 = "C:\materiais_novos.txt" 'Arquivo com materiais novos
    Dim intNumArq1%: intNumArq1 = 1
    Dim intReg1%: intReg1 = 523
    Dim varDadoCell1 As Variant
    
    Open strNovoArq1 For Output As #intNumArq1 'Exporta os dados da planilha para o arquivo criado
    Print #intNumArq1, "INSERT INTO tb_neo_mat (mat_id_pk, mat_gtin13, mat_codneo_sku, mat_desc, mat_dat_cad, mat_dat_alt, pdt_id_fk, mar_id_fk) VALUES"
    
    'Dim strNovoArq2$: strNovoArq2 = "C:\teste.txt" 'Arquivo com materiais novos
    'Dim intNumArq2%: intNumArq2 = 2
    'Dim varDadoCell2 As Variant
    'Open strNovoArq2 For Output As #intNumArq2 'Exporta os dados da planilha para o arquivo criado
    
    Windows("Estoque.xls").Activate
    
    For i = 6 To lngUL_WS_1
        For j = 2 To lngUL_WS_2
            If (CStr(WS_1.Range("A" & CStr(i)).Value) = CStr(WS_2.Range("A" & CStr(j)).Value)) Then
                booAchou = True
                Exit For
            End If
        Next j
        If Not booAchou Then
            lngUL_WS_2 = lngUL_WS_2 + 1
            WS_2.Range("A" & CStr(lngUL_WS_2)).Value = WS_1.Range("A" & CStr(lngRow)).Value
            WS_2.Range("B" & CStr(lngUL_WS_2)).Value = WS_1.Range("B" & CStr(lngRow)).Value
            WS_2.Range("C" & CStr(lngUL_WS_2)).Value = WS_1.Range("C" & CStr(lngRow)).Value
            WS_2.Range("D" & CStr(lngUL_WS_2)).Value = WS_1.Range("G" & CStr(lngRow)).Value
            WS_2.Range("E" & CStr(lngUL_WS_2)).Value = WS_1.Range("H" & CStr(lngRow)).Value
            WS_2.Range("A" & CStr(lngUL_WS_2) & ":C" & CStr(lngUL_WS_2)).Select
            With Selection.Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .Color = 15773696
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
            intReg1 = intReg1 + 1
            varDadoCell1 = "(" & CStr(intReg1) & ", '0000000000000', '" & WS_2.Range("A" & CStr(lngUL_WS_2)).Value & "', '" & WS_2.Range("B" & CStr(lngUL_WS_2)).Value & "', NOW(), NOW(), " & WS_2.Range("E" & CStr(lngUL_WS_2)).Value & ", 0),"
            If varDadoCell1 <> "" Then Print #intNumArq1, varDadoCell1
            varDadoCell1 = ""
        End If
    Next i
    Close #intNumArq1 'Salva e fecha o arquivo de texto com os dados
    
    WS_2.Select
    
    For j = 2 To lngUL_WS_2
        WS_2.Range("A" & CStr(j) & ":I" & CStr(j)).Select
        With Selection.Interior
            .Pattern = xlNone
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
        For i = 6 To lngUL_WS_1
            If (CStr(WS_1.Range("A" & CStr(i)).Value) = CStr(WS_2.Range("A" & CStr(j)).Value)) Then
                WS_2.Range("C" & CStr(j)).Value = WS_1.Range("C" & CStr(i)).Value
                WS_2.Range("C" & CStr(j)).EntireColumn.AutoFit
                WS_2.Range("D" & CStr(j)).Value = WS_1.Range("G" & CStr(i)).Value
                WS_2.Range("D" & CStr(j)).EntireColumn.AutoFit
                    
                If (WS_2.Range("D" & CStr(j)).Value = 0) Then
                    WS_2.Range("J" & CStr(j)).Value = 0
                    WS_2.Range("A" & CStr(j) & ":J" & CStr(j)).Select
                    With Selection.Interior
                        .Pattern = xlSolid
                        .PatternColorIndex = xlAutomatic
                        .Color = 10092543
                        .TintAndShade = 0
                        .PatternTintAndShade = 0
                    End With
                End If
                    
                Select Case WS_2.Range("C" & CStr(j)).Value
                    Case "CH"
                        WS_2.Range("C" & CStr(j)).Value = "UNID"
                    Case "UN"
                        WS_2.Range("C" & CStr(j)).Value = "UNID"
                    Case "M²"
                        WS_2.Range("C" & CStr(j)).Value = "M2"
                    Case "PCT"
                        WS_2.Range("C" & CStr(j)).Value = "PACOTE"
                    Case "RL"
                        WS_2.Range("C" & CStr(j)).Value = "ROLO"
                    Case "KG"
                        WS_2.Range("C" & CStr(j)).Value = "KG"
                End Select
                WS_2.Range("C" & CStr(j)).EntireColumn.AutoFit
                If (WS_2.Range("D" & CStr(j)).Value > 0) And (WS_2.Range("I" & CStr(j)).Value = 0) Then
                    WS_2.Range("D" & CStr(j) & ":I" & CStr(j)).Select
                    With Selection.Interior
                        .Pattern = xlSolid
                        .PatternColorIndex = xlAutomatic
                        .Color = 65535
                        .TintAndShade = 0
                        .PatternTintAndShade = 0
                    End With
                    'varDadoCell2 = WS_2.Range("A" & CStr(j)).Value & " -> " & WS_2.Range("B" & CStr(j)).Value & " -> " & WS_2.Range("C" & CStr(j)).Value & " -> KG =  "
                    'If varDadoCell2 <> "" Then Print #intNumArq2, varDadoCell2
                        'varDadoCell2 = ""
                    'End If
                End If
                Exit For
            End If
        Next i
    Next j
    
    Close #intNumArq2 'Salva e fecha o arquivo de texto com os dados
    
    Set WS_2 = Nothing
    Set WB_2 = Nothing
    Set WS_1 = Nothing
    Set WB_1 = Nothing
End Sub
'--------------------------------------------------------------------------------------------------------------------------------------------------------------'
Sub Exportar_Estoque_Atual()
    Dim WS_1 As Worksheet: Set WS_1 = Worksheets("cont_est")
    Dim strNovoArq1$
    Dim intNumArq1%: intNumArq1 = 1
    Dim intReg1%: intReg1 = 1
    Dim lngRow&: lngRow = 2
    Dim varDadoCell1 As Variant
    
    strNovoArq1 = "C:\materiais_estoque_atual.txt" 'Arquivo com estoque atual de materiais
    Open strNovoArq1 For Output As #intNumArq1 'Exporta os dados da planilha para o arquivo criado
    Print #intNumArq1, "INSERT INTO tb_neo_entext (entext_id_pk, entext_lote, entext_qtd_p, entext_qtd_s, entext_dat_cad, entext_dat_alt, entext_obs, entext_dest, uncom_id_fk_p, uncom_id_fk_s, endest_id_fk, clb_id_fk, mat_id_fk, nfe_id_fk) VALUES"
    
    With WS_1
        Do While (.Range("A" & CStr(lngRow)).Value <> "")
            .Range("A" & CStr(lngRow)).Select
            If (CLng(.Range("D" & CStr(lngRow)).Value) > 0) Then
                Select Case .Range("E" & CStr(lngRow)).Value
                    Case 8
                        varDadoCell1 = "(" & CStr(intReg1) & ", '1111111111', " & Replace(Round(.Range("I" & CStr(lngRow)).Value, 3), ",", ".") & ", " & Replace(Round(.Range("D" & CStr(lngRow)).Value, 3), ",", ".") & ", NOW(), NOW(), 'Fitas / Inventario 26/09/2020', 'P', '" & .Range("H" & CStr(lngRow)).Value & "', '" & .Range("C" & CStr(lngRow)).Value & "', '0104010', '10320', " & .Range("F" & CStr(lngRow)).Value & ", 1),"
                    Case 24
                        varDadoCell1 = "(" & CStr(intReg1) & ", '2222222222', " & Replace(Round(.Range("I" & CStr(lngRow)).Value, 3), ",", ".") & ", " & Replace(Round(.Range("D" & CStr(lngRow)).Value, 3), ",", ".") & ", NOW(), NOW(), 'Fitas e PDV / Inventario 26/09/2020', 'P', '" & .Range("H" & CStr(lngRow)).Value & "', '" & .Range("C" & CStr(lngRow)).Value & "', '0104012', '10320', " & .Range("F" & CStr(lngRow)).Value & ", 2),"
                    Case 25
                        varDadoCell1 = "(" & CStr(intReg1) & ", '3333333333', " & Replace(Round(.Range("I" & CStr(lngRow)).Value, 3), ",", ".") & ", " & Replace(Round(.Range("D" & CStr(lngRow)).Value, 3), ",", ".") & ", NOW(), NOW(), 'PDV / Inventario 26/09/2020', 'P', '" & .Range("H" & CStr(lngRow)).Value & "', '" & .Range("C" & CStr(lngRow)).Value & "', '0104014', '10320', " & .Range("F" & CStr(lngRow)).Value & ", 3),"
                End Select
                intReg1 = intReg1 + 1
                
                If (.Range("A" & CStr(lngRow + 1)).Value = "") Then varDadoCell1 = varDadoCell1 + ";"
            End If
            lngRow = lngRow + 1
            If varDadoCell1 <> "" Then Print #intNumArq1, varDadoCell1
            varDadoCell1 = ""
        Loop
    End With
    Close #intNumArq1 'Salva e fecha o arquivo de texto com os dados
    Set WS_1 = Nothing
End Sub
'--------------------------------------------------------------------------------------------------------------------------------------------------------------'
Sub Atualiza_Estoque_Atual()
    Dim WS_1 As Worksheet: Set WS_1 = Worksheets("cont_est")
    Dim strNovoArq1$
    Dim intNumArq1%: intNumArq1 = 1
    Dim intReg1%: intReg1 = 1
    Dim lngRow&: lngRow = 2
    Dim varDadoCell1 As Variant
    
    strNovoArq1 = "C:\atualiza_estoque_material.txt" 'Arquivo com estoque atual de materiais
    Open strNovoArq1 For Output As #intNumArq1 'Exporta os dados da planilha para o arquivo criado
    
    With WS_1
        Do While (.Range("A" & CStr(lngRow)).Value <> "")
            .Range("A" & CStr(lngRow)).Select
            If (CLng(.Range("D" & CStr(lngRow)).Value) > 0) Then
                Select Case .Range("E" & CStr(lngRow)).Value
                    Case 8
                        varDadoCell1 = "UPDATE tb_neo_entext SET tb_neo_entext.entext_qtd_p=" & Replace(Round(.Range("I" & CStr(lngRow)).Value, 3), ",", ".") & ", tb_neo_entext.entext_qtd_s=" & Replace(Round(.Range("D" & CStr(lngRow)).Value, 3), ",", ".") & " WHERE (tb_neo_entext.entext_lote='1111111111') AND (tb_neo_entext.nfe_id_fk=1) AND (tb_neo_entext.mat_id_fk=" & .Range("F" & CStr(lngRow)).Value & ");"
                    Case 24
                        varDadoCell1 = "UPDATE tb_neo_entext SET tb_neo_entext.entext_qtd_p=" & Replace(Round(.Range("I" & CStr(lngRow)).Value, 3), ",", ".") & ", tb_neo_entext.entext_qtd_s=" & Replace(Round(.Range("D" & CStr(lngRow)).Value, 3), ",", ".") & " WHERE (tb_neo_entext.entext_lote='2222222222') AND (tb_neo_entext.nfe_id_fk=2) AND (tb_neo_entext.mat_id_fk=" & .Range("F" & CStr(lngRow)).Value & ");"
                    Case 25
                        varDadoCell1 = "UPDATE tb_neo_entext SET tb_neo_entext.entext_qtd_p=" & Replace(Round(.Range("I" & CStr(lngRow)).Value, 3), ",", ".") & ", tb_neo_entext.entext_qtd_s=" & Replace(Round(.Range("D" & CStr(lngRow)).Value, 3), ",", ".") & " WHERE (tb_neo_entext.entext_lote='3333333333') AND (tb_neo_entext.nfe_id_fk=3) AND (tb_neo_entext.mat_id_fk=" & .Range("F" & CStr(lngRow)).Value & ");"
                End Select
            End If
            lngRow = lngRow + 1
            If varDadoCell1 <> "" Then Print #intNumArq1, varDadoCell1
            varDadoCell1 = ""
        Loop
    End With
    
    Close #intNumArq1 'Salva e fecha o arquivo de texto com os dados
    
    Set WS_1 = Nothing
End Sub
'--------------------------------------------------------------------------------------------------------------------------------------------------------------'
Sub Atualiza_Codigo_Material()
    Dim WS_1 As Worksheet: Set WS_1 = Worksheets("cont_est")
    Dim WS_2 As Worksheet: Set WS_2 = Worksheets("materiais_novos")
    Dim strCol$: strCol = "A"
    Dim lngUL_WS_1&, lngUL_WS_2&
    
    WS_1.Select
    
    lngUL_WS_1 = WS_1.Cells(Rows.Count, 1).End(xlUp).Row
    lngUL_WS_2 = WS_2.Cells(Rows.Count, 1).End(xlUp).Row
    
    For i = 2 To lngUL_WS_1
        WS_1.Range("F" & CStr(i)).Select
        For j = 1 To lngUL_WS_2
            If (CStr(WS_1.Range("A" & CStr(i)).Value) = CStr(WS_2.Range("F" & CStr(j)).Value)) Then
                WS_1.Range("E" & CStr(i)).Value = WS_2.Range("N" & CStr(j)).Value
                WS_1.Range("E" & CStr(i)).EntireColumn.AutoFit
                WS_1.Range("F" & CStr(i)).Value = WS_2.Range("B" & CStr(j)).Value
                WS_1.Range("F" & CStr(i)).EntireColumn.AutoFit
                Exit For
            End If
        Next j
    Next i
    Set WS_2 = Nothing
    Set WS_1 = Nothing
End Sub
'--------------------------------------------------------------------------------------------------------------------------------------------------------------'
Sub Compara_Planilhas()
    Dim WS_1 As Worksheet: Set WS_1 = Worksheets("materiais_novos")
    Dim WS_2 As Worksheet: Set WS_2 = Worksheets("cont_est")
    Dim lngUL_WS_1&, lngUL_WS_2&
    
    WS_1.Select
    
    lngUL_WS_1 = WS_1.Cells(Rows.Count, 1).End(xlUp).Row
    lngUL_WS_2 = WS_2.Cells(Rows.Count, 1).End(xlUp).Row
    
    For i = 1 To lngUL_WS_1
        WS_1.Range("F" & CStr(i)).Select
        For j = 2 To lngUL_WS_2
             If (CStr(WS_1.Range("F" & CStr(i)).Value) = CStr(WS_2.Range("A" & CStr(j)).Value)) Then
                If (WS_1.Range("N" & CStr(i)).Value = 8) Or (WS_1.Range("N" & CStr(i)).Value = 24) Or (WS_1.Range("N" & CStr(i)).Value = 25) Then
                    With Selection.Interior
                        .Pattern = xlSolid
                        .PatternColorIndex = xlAutomatic
                        .Color = 10092543
                        .TintAndShade = 0
                        .PatternTintAndShade = 0
                    End With
                    Exit For
                End If
            End If
        Next j
    Next i
    Set WS_2 = Nothing
    Set WS_1 = Nothing
End Sub
'--------------------------------------------------------------------------------------------------------------------------------------------------------------'