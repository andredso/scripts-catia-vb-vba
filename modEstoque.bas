Attribute VB_Name = "modEstoque"
Sub Atualiza_Quantidades()
    'Call Atualiza_Materiais
    
    ChDir "C:\Intel\Desenv"
    Workbooks.Open Filename:="C:\Intel\Desenv\CONTROLE SANOL.xlsx", Password:="271411"
    Sheets("EstoqueAtual-e-Cadastro").Visible = True
    Range("A6").Select
    
    Dim WB_1 As Workbook: Set WB_1 = Workbooks("CONTROLE SANOL.xlsx")
    Dim WS_1 As Worksheet: Set WS_1 = WB_1.Worksheets("EstoqueAtual-e-Cadastro")
    Dim WB_2 As Workbook: Set WB_2 = Workbooks("Estoque.xls")
    Dim WS_2 As Worksheet: Set WS_2 = WB_2.Worksheets("cont_est")
    
    Windows("Estoque.xls").Activate
    
    Dim lngUL_WS_1&: lngUL_WS_1 = WS_1.Cells(Rows.Count, 1).End(xlUp).Row
    Dim lngUL_WS_2&: lngUL_WS_2 = WS_2.Cells(Rows.Count, 1).End(xlUp).Row
    
    WS_2.Select
    
    For j = 2 To lngUL_WS_2
        WS_2.Range("A" & CStr(j) & ":J" & CStr(j)).Select
        With Selection.Interior
            .Color = automatic
            .Pattern = xlNone
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
        
        For i = 6 To lngUL_WS_1
            If (CStr(WS_1.Range("A" & CStr(i)).Value) = CStr(WS_2.Range("A" & CStr(j)).Value)) _
            And (WS_2.Range("D" & CStr(j)).Value <> WS_1.Range("G" & CStr(i)).Value) Then
                WS_2.Range("C" & CStr(j)).Value = WS_1.Range("C" & CStr(i)).Value
                WS_2.Range("C" & CStr(j)).EntireColumn.AutoFit
                WS_2.Range("D" & CStr(j)).Value = WS_1.Range("G" & CStr(i)).Value
                WS_2.Range("D" & CStr(j)).EntireColumn.AutoFit
                WS_2.Range("J" & CStr(j)).Value = ""
                
                'WS_2.Range("I" & CStr(j)).Select
                'With Selection.Font
                '    .Color = -1003520
                '    .TintAndShade = 0
                'End With
                'If (WS_2.Range("D" & CStr(j)).Value > 0) And (WS_2.Range("I" & CStr(j)).Value = 0) Then
                '    WS_2.Range("D" & CStr(j) & ":J" & CStr(j)).Select
                '    With Selection.Interior
                '        .Pattern = xlSolid
                '        .PatternColorIndex = xlAutomatic
                '        .Color = 65535
                '        .TintAndShade = 0
                '        .PatternTintAndShade = 0
                '    End With
                'End If
                Exit For
            End If
        Next i
        
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
    Next j
    
    WS_2.Range("A1").Select
    
    WB_2.Save
    WB_1.Close
    
    Set WS_2 = Nothing
    Set WB_2 = Nothing
    Set WS_1 = Nothing
    Set WB_1 = Nothing
    
    Call Concilia_Estoque
End Sub

Sub Exportar_Estoque_Atual()
    Dim WS_1 As Worksheet: Set WS_1 = Worksheets("cont_est")
    Dim strNovoArq1$
    Dim intNumArq1%: intNumArq1 = 1
    Dim intReg1%: intReg1 = 82
    Dim lngRow&: lngRow = 2
    Dim varDadoCell1 As Variant
    
    strNovoArq1 = "C:\materiais_estoque_atual.sql" 'Arquivo com estoque atual de materiais
    Open strNovoArq1 For Output As #intNumArq1 'Exporta os dados da planilha para o arquivo criado
    Print #intNumArq1, "DELETE FROM tb_neo_entext;"
    Print #intNumArq1, "INSERT INTO tb_neo_entext (entext_id_pk, entext_lote, entext_qtd_p, entext_qtd_s, entext_dat_cad, entext_dat_alt, entext_obs, entext_dest, uncom_id_fk_p, uncom_id_fk_s, endest_id_fk, clb_id_fk, mat_id_fk, nfe_id_fk) VALUES"
    
    With WS_1
        Do While (.Range("A" & CStr(lngRow)).Value <> "")
            .Range("A" & CStr(lngRow)).Select
            If (CLng(.Range("D" & CStr(lngRow)).Value) > 0) Then
                Select Case .Range("E" & CStr(lngRow)).Value
                    Case 1
                        varDadoCell1 = "(" & CStr(intReg1) & ", '1111111111', " & Replace(Round(.Range("I" & CStr(lngRow)).Value, 3), ",", ".") & ", " & Replace(Round(.Range("D" & CStr(lngRow)).Value, 3), ",", ".") & ", NOW(), NOW(), 'Fitas / Inventário 17/02/2021', 'P', '" & .Range("H" & CStr(lngRow)).Value & "', '" & .Range("C" & CStr(lngRow)).Value & "', '0104010', '10320', " & .Range("F" & CStr(lngRow)).Value & ", 1),"
                    Case 2
                        varDadoCell1 = "(" & CStr(intReg1) & ", '2222222222', " & Replace(Round(.Range("I" & CStr(lngRow)).Value, 3), ",", ".") & ", " & Replace(Round(.Range("D" & CStr(lngRow)).Value, 3), ",", ".") & ", NOW(), NOW(), 'Fitas e PDV / Inventário 17/02/2021', 'P', '" & .Range("H" & CStr(lngRow)).Value & "', '" & .Range("C" & CStr(lngRow)).Value & "', '0104012', '10320', " & .Range("F" & CStr(lngRow)).Value & ", 2),"
                    Case 3
                        varDadoCell1 = "(" & CStr(intReg1) & ", '3333333333', " & Replace(Round(.Range("I" & CStr(lngRow)).Value, 3), ",", ".") & ", " & Replace(Round(.Range("D" & CStr(lngRow)).Value, 3), ",", ".") & ", NOW(), NOW(), 'PDV / Inventário 17/02/2021', 'P', '" & .Range("H" & CStr(lngRow)).Value & "', '" & .Range("C" & CStr(lngRow)).Value & "', '0104014', '10320', " & .Range("F" & CStr(lngRow)).Value & ", 3),"
                    Case 4
                        varDadoCell1 = "(" & CStr(intReg1) & ", NULL, " & Replace(Round(.Range("I" & CStr(lngRow)).Value, 3), ",", ".") & ", " & Replace(Round(.Range("D" & CStr(lngRow)).Value, 3), ",", ".") & ", NOW(), NOW(), NULL, 'L', '" & .Range("H" & CStr(lngRow)).Value & "', '" & .Range("C" & CStr(lngRow)).Value & "', '0104016', '10320', " & .Range("F" & CStr(lngRow)).Value & ", ''),"
                    Case 5
                        varDadoCell1 = "(" & CStr(intReg1) & ", NULL, " & Replace(Round(.Range("I" & CStr(lngRow)).Value, 3), ",", ".") & ", " & Replace(Round(.Range("D" & CStr(lngRow)).Value, 3), ",", ".") & ", NOW(), NOW(), NULL, 'L', '" & .Range("H" & CStr(lngRow)).Value & "', '" & .Range("C" & CStr(lngRow)).Value & "', '0104016', '10320', " & .Range("F" & CStr(lngRow)).Value & ", ''),"
                    Case 6
                        varDadoCell1 = "(" & CStr(intReg1) & ", NULL, " & Replace(Round(.Range("I" & CStr(lngRow)).Value, 3), ",", ".") & ", " & Replace(Round(.Range("D" & CStr(lngRow)).Value, 3), ",", ".") & ", NOW(), NOW(), NULL, 'L', '" & .Range("H" & CStr(lngRow)).Value & "', '" & .Range("C" & CStr(lngRow)).Value & "', '0105017', '10320', " & .Range("F" & CStr(lngRow)).Value & ", ''),"
                    Case 7
                        varDadoCell1 = "(" & CStr(intReg1) & ", '7777777777', " & Replace(Round(.Range("I" & CStr(lngRow)).Value, 3), ",", ".") & ", " & Replace(Round(.Range("D" & CStr(lngRow)).Value, 3), ",", ".") & ", NOW(), NOW(), 'Personalização / Inventário 17/02/2021', 'P', '" & .Range("H" & CStr(lngRow)).Value & "', '" & .Range("C" & CStr(lngRow)).Value & "', '0105019', '10320', " & .Range("F" & CStr(lngRow)).Value & ", 5),"
                End Select
                intReg1 = intReg1 + 1
                
                If (.Range("A" & CStr(lngRow + 1)).Value = "") Then varDadoCell1 = Mid(varDadoCell1, 1, Len(varDadoCell1) - 1) + ";"
            End If
            lngRow = lngRow + 1
            If (varDadoCell1 <> "") Then Print #intNumArq1, varDadoCell1
            varDadoCell1 = ""
        Loop
        .Range("A1").Select
    End With
    Close #intNumArq1 'Salva e fecha o arquivo de texto com os dados
    Set WS_1 = Nothing
End Sub

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
                'If (WS_1.Range("N" & CStr(i)).Value = 1) Or (WS_1.Range("N" & CStr(i)).Value = 2) Or (WS_1.Range("N" & CStr(i)).Value = 3) Or (WS_1.Range("N" & CStr(i)).Value = 4) Or (WS_1.Range("N" & CStr(i)).Value = 5) Or (WS_1.Range("N" & CStr(i)).Value = 6) Or (WS_1.Range("N" & CStr(i)).Value = 7) Then
                With Selection.Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .Color = 10092543
                    .TintAndShade = 0
                    .PatternTintAndShade = 0
                End With
                Exit For
                'End If
            End If
        Next j
    Next i
    Set WS_2 = Nothing
    Set WS_1 = Nothing
End Sub

Sub Atualiza_Peso_Material()
    Dim WB_1 As Workbook: Set WB_1 = Workbooks("Estoque.xls")
    Dim WS_1 As Worksheet: Set WS_1 = WB_1.Worksheets("cont_est")
    Dim WS_2 As Worksheet: Set WS_2 = WB_1.Worksheets("materiais_novos")
    
    Dim lngUL_WS_1&, lngUL_WS_2&
    
    WS_2.Select
    
    lngUL_WS_1 = WS_1.Cells(Rows.Count, 1).End(xlUp).Row
    lngUL_WS_2 = WS_2.Cells(Rows.Count, 1).End(xlUp).Row
    
    For i = 1 To lngUL_WS_2
        For j = 2 To lngUL_WS_1
            If (CStr(WS_1.Range("A" & CStr(j)).Value) = CStr(WS_2.Range("F" & CStr(i)).Value)) _
            And (CStr(WS_1.Range("F" & CStr(j)).Value) = CStr(WS_2.Range("B" & CStr(i)).Value)) _
            And (WS_2.Range("R" & CStr(i)).Value <> WS_1.Range("G" & CStr(j)).Value) Then
                WS_2.Range("R" & CStr(i)).Select
                Selection.NumberFormat = "@"
                WS_2.Range("R" & CStr(i)).Value = Replace(WS_1.Range("G" & CStr(j)).Value, ",", ".")
                'WS_2.Range("R" & CStr(i)).Value = WS_1.Range("G" & CStr(j)).Value
                WS_2.Range("R" & CStr(i)).EntireColumn.AutoFit
                Exit For
            End If
        Next j
    Next i
        
    WB_1.Save
    
    Set WS_2 = Nothing
    Set WS_1 = Nothing
    Set WB_1 = Nothing
End Sub

Sub Atualiza_Materiais()
    ChDir "C:\Intel\Desenv"
    Workbooks.Open Filename:="C:\Intel\Desenv\CONTROLE SANOL.xlsx", Password:="271411"
    Sheets("EstoqueAtual-e-Cadastro").Visible = True
    Range("A4").Select
    
    Dim WB_1 As Workbook: Set WB_1 = Workbooks("CONTROLE SANOL.xlsx")
    Dim WS_1 As Worksheet: Set WS_1 = WB_1.Worksheets("EstoqueAtual-e-Cadastro")
    
    Dim WB_2 As Workbook: Set WB_2 = Workbooks("Estoque.xls")
    Dim WS_2 As Worksheet: Set WS_2 = WB_2.Worksheets("materiais_novos")
    Dim WS_3 As Worksheet: Set WS_3 = WB_2.Worksheets("cont_est")
    
    Windows("Estoque.xls").Activate
    
    Dim lngUL_WS_1&: lngUL_WS_1 = WS_1.Cells(Rows.Count, 1).End(xlUp).Row
    Dim lngUL_WS_2&: lngUL_WS_2 = WS_2.Cells(Rows.Count, 1).End(xlUp).Row
    Dim lng_Cod_Mat&: lng_Cod_Mat = lngUL_WS_2
    Dim booAchou As Boolean
    
    'Windows("Estoque.xls").Activate
    
    For i = 6 To lngUL_WS_1
        booAchou = False
        For j = 1 To lngUL_WS_2
            If (CStr(WS_1.Range("A" & CStr(i)).Value) = CStr(WS_2.Range("F" & CStr(j)).Value)) Then
                booAchou = True
                Exit For
            End If
        Next j
        If (Not booAchou) Then
            WS_2.Select
            lng_Cod_Mat = lng_Cod_Mat + 1
            WS_2.Range("A" & CStr(lng_Cod_Mat)).Value = "("
            WS_2.Range("B" & CStr(lng_Cod_Mat)).Value = lng_Cod_Mat
            WS_2.Range("C" & CStr(lng_Cod_Mat)).Value = Chr(34) + Chr(44) + Chr(34) '","
            WS_2.Range("D" & CStr(lng_Cod_Mat)).Value = "0000000000000"
            WS_2.Range("E" & CStr(lng_Cod_Mat)).Value = Chr(34) + Chr(44) + Chr(34)
            WS_2.Range("F" & CStr(lng_Cod_Mat)).Value = WS_1.Range("A" & CStr(i)).Value: WS_2.Range("F" & CStr(lng_Cod_Mat)).HorizontalAlignment = xlCenter
            WS_2.Range("G" & CStr(lng_Cod_Mat)).Value = Chr(34) + Chr(44) + Chr(34)
            WS_2.Range("H" & CStr(lng_Cod_Mat)).Value = WS_1.Range("B" & CStr(i)).Value
            WS_2.Range("I" & CStr(lng_Cod_Mat)).Value = Chr(34) + Chr(44) + Chr(34)
            WS_2.Range("J" & CStr(lng_Cod_Mat)).Value = "NOW()"
            WS_2.Range("K" & CStr(lng_Cod_Mat)).Value = ","
            WS_2.Range("L" & CStr(lng_Cod_Mat)).Value = "NOW()"
            WS_2.Range("M" & CStr(lng_Cod_Mat)).Value = ","
            WS_2.Range("N" & CStr(lng_Cod_Mat)).Value = WS_1.Range("H" & CStr(i)).Value: WS_2.Range("N" & CStr(lng_Cod_Mat)).HorizontalAlignment = xlCenter
            WS_2.Range("O" & CStr(lng_Cod_Mat)).Value = ","
            WS_2.Range("P" & CStr(lng_Cod_Mat)).Value = "0"
            WS_2.Range("Q" & CStr(lng_Cod_Mat)).Value = ","
            WS_2.Range("R" & CStr(lng_Cod_Mat)).Value = "0"
            WS_2.Range("S" & CStr(lng_Cod_Mat)).Value = "),"
            'WS_2.Range("A" & CStr(lng_Cod_Mat) & ":Q" & CStr(lng_Cod_Mat)).Select
            'With Selection.Interior
            '    .Pattern = xlSolid
            '    .PatternColorIndex = xlAutomatic
            '    .Color = 15773696
            '    .TintAndShade = 0
            '    .PatternTintAndShade = 0
            'End With
            WS_3.Select
            WS_3.Range("A" & CStr(lng_Cod_Mat + 1)).Value = WS_1.Range("A" & CStr(i)).Value: WS_3.Range("A" & CStr(lng_Cod_Mat + 1)).HorizontalAlignment = xlCenter
            WS_3.Range("B" & CStr(lng_Cod_Mat + 1)).Value = WS_1.Range("B" & CStr(i)).Value
            WS_3.Range("C" & CStr(lng_Cod_Mat + 1)).Value = WS_1.Range("C" & CStr(i)).Value: WS_3.Range("C" & CStr(lng_Cod_Mat + 1)).HorizontalAlignment = xlCenter
            WS_3.Range("D" & CStr(lng_Cod_Mat + 1)).Value = WS_1.Range("G" & CStr(i)).Value
            WS_3.Range("E" & CStr(lng_Cod_Mat + 1)).Value = WS_1.Range("H" & CStr(i)).Value: WS_3.Range("E" & CStr(lng_Cod_Mat + 1)).HorizontalAlignment = xlCenter
            WS_3.Range("F" & CStr(lng_Cod_Mat + 1)).Value = lng_Cod_Mat: WS_3.Range("F" & CStr(lng_Cod_Mat + 1)).HorizontalAlignment = xlCenter
            WS_3.Range("G" & CStr(lng_Cod_Mat + 1)).Value = "0": WS_3.Range("E" & CStr(lng_Cod_Mat + 1)).HorizontalAlignment = xlLeft
            WS_3.Range("H" & CStr(lng_Cod_Mat + 1)).Value = "KG": WS_3.Range("H" & CStr(lng_Cod_Mat + 1)).HorizontalAlignment = xlCenter
            'WS_3.Range("I" & CStr(lng_Cod_Mat + 1)).Formula = "=IF(C" & Trim(CStr(lng_Cod_Mat + 1)) & "<>""KG"";D" & Trim(CStr(lng_Cod_Mat + 1)) & "*G" & Trim(CStr(lng_Cod_Mat + 1)) & ";IF(D" & Trim(CStr(lng_Cod_Mat + 1)) & ">0;D" & Trim(CStr(lng_Cod_Mat + 1)) & ";0))"
            'WS_3.Range("A" & CStr(lng_Cod_Mat) & ":I" & CStr(lng_Cod_Mat)).Select
            'With Selection.Interior
            '    .Pattern = xlSolid
            '    .PatternColorIndex = xlAutomatic
            '    .Color = 15773696
            '    .TintAndShade = 0
            '    .PatternTintAndShade = 0
            'End With
        End If
    Next i
    
    WB_2.Save
    WB_1.Close
    
    Set WS_3 = Nothing
    Set WS_2 = Nothing
    Set WS_1 = Nothing
    Set WB_2 = Nothing
    Set WB_1 = Nothing
End Sub

Sub Concilia_Estoque()
    Dim strMatID$
    Dim lngSaldo&, lngLinha&
    
    ChDir "C:\Intel\Desenv"
    Workbooks.Open Filename:="C:\Intel\Desenv\concilia_estoque.xls"
    
    Dim WB_1 As Workbook: Set WB_1 = Workbooks("concilia_estoque.xls")
    Dim WS_1 As Worksheet: Set WS_1 = WB_1.Worksheets("Plan1")
    
    Dim WB_2 As Workbook: Set WB_2 = Workbooks("Estoque.xls")
    Dim WS_2 As Worksheet: Set WS_2 = WB_2.Worksheets("cont_est")
    
    Windows("concilia_estoque.xls").Activate
    
    Dim lngUL_WS_1&: lngUL_WS_1 = WS_1.Cells(Rows.Count, 1).End(xlUp).Row
    Dim lngUL_WS_2&: lngUL_WS_2 = WS_2.Cells(Rows.Count, 1).End(xlUp).Row
    
    Call Limpa_Concilia_Estoque
    
    WS_1.Select
    lngLinha = 0
    For i = 2 To lngUL_WS_1
        strMatID = CStr(WS_1.Range("F" & CStr(i)).Value)
        lngSaldo = CInt(WS_1.Range("J" & CStr(i)).Value)
        For j = i + 1 To lngUL_WS_1
            If (strMatID = CStr(WS_1.Range("F" & CStr(j)).Value)) Then
                lngSaldo = lngSaldo + CInt(WS_1.Range("J" & CStr(j)).Value)
                If lngLinha = 0 Then
                    lngLinha = i
                End If
                i = j
            Else
                Exit For
            End If
        Next j
        
        For k = 2 To lngUL_WS_2
            If (strMatID = CStr(WS_2.Range("F" & CStr(k)).Value)) Then
                If (lngSaldo <> CInt(WS_2.Range("D" & CStr(k)).Value)) Then
                    WS_1.Range("F" & CStr(i) & ":K" & CStr(i)).Select
                    With Selection.Interior
                        .Color = automatic
                        .Pattern = xlNone
                        .TintAndShade = 0
                        .PatternTintAndShade = 0
                    End With
                    If lngLinha > 0 Then
                        WS_1.Range("F" & CStr(lngLinha) & ":K" & CStr(i)).Select
                    Else
                        WS_1.Range("F" & CStr(i) & ":K" & CStr(i)).Select
                    End If
                    With Selection.Interior
                        .Pattern = xlSolid
                        .PatternColorIndex = xlAutomatic
                        .Color = 10092543
                        .TintAndShade = 0
                        .PatternTintAndShade = 0
                    End With
                End If
                lngLinha = 0
                Exit For
            End If
        Next k
    Next i
    
    'WB_2.Close
    WB_1.Save
    
    Set WS_2 = Nothing
    Set WB_2 = Nothing
    Set WS_1 = Nothing
    Set WB_1 = Nothing
End Sub

Sub Limpa_Concilia_Estoque()
    Columns("F:K").Select
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Cells.Select
    Selection.QueryTable.Delete
    Selection.ClearContents
    Range("A1").Select
    ActiveWorkbook.Save
    With ActiveSheet.QueryTables.Add(Connection:="TEXT;C:\Intel\Desenv\concilia.csv", _
        Destination:=Range("$A$1"))
        .Name = "concilia"
        .FieldNames = True
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .TextFilePromptOnRefresh = False
        .TextFilePlatform = 1252
        .TextFileStartRow = 1
        .TextFileParseType = xlDelimited
        .TextFileTextQualifier = xlTextQualifierDoubleQuote
        .TextFileConsecutiveDelimiter = False
        .TextFileTabDelimiter = False
        .TextFileSemicolonDelimiter = True
        .TextFileCommaDelimiter = False
        .TextFileSpaceDelimiter = False
        .TextFileColumnDataTypes = Array(1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1)
        .TextFileTrailingMinusNumbers = True
        .Refresh BackgroundQuery:=False
    End With
End Sub

