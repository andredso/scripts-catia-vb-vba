Option Explicit

Private Sub cmbCodProd_Change()
    If (cmbCodProd.ListIndex < 0) Then
        Exit Sub
    Else
        lblCodBarDescProd.Caption = ""
        lblCodBarProd.Caption = ""
        lblCompProd.Caption = ""
        lblDescProd.Caption = ""
        lblLargProd.Caption = ""
        txtLargCorte.Text = "": txtLargCorte.Locked = True: txtLargCorte.Enabled = False
        txtComp.Text = "": txtComp.Locked = True: txtComp.Enabled = False
        txtDiamInt.Text = "" ': txtDiamInt.Locked = True: txtDiamInt.Enabled = False
        txtDirEnr.Text = "" ': txtDirEnr.Locked = True: txtDirEnr.Enabled = False
        txtEmen.Text = "" ': txtEmen.Locked = True: txtEmen.Enabled = False
        txtAI01.Text = "": txtAI01.Locked = True: txtAI01.Enabled = False
        txtAI10.Text = "": txtAI10.Locked = True: txtAI10.Enabled = False
        txtAI8001.Text = "": txtAI8001.Locked = True: txtAI8001.Enabled = False
        txtGS1_128.Text = "": txtGS1_128.Locked = True: txtGS1_128.Enabled = False
        txtLeitGS1_128.Text = "": txtLeitGS1_128.Locked = True: txtLeitGS1_128.Enabled = False
        Dim WS_1 As Worksheet: Set WS_1 = Worksheets(cmbFab.Text)
        Dim strNomePais$: strNomePais = ""
        With WS_1
            strNomePais = Verifica_CodPais(Trim(.Range("A" & CStr(cmbCodProd.ListIndex + 2)).Value))
            If (strNomePais = "") Then
                Set WS_1 = Nothing
                Exit Sub
            End If
            lblDescProd.Caption = Trim(.Range("C" & CStr(cmbCodProd.ListIndex + 2)).Value)
            lblLargProd.Caption = Trim(.Range("D" & CStr(cmbCodProd.ListIndex + 2)).Value)
            lblLargCorte.Caption = "milímetros (largura mínima de 5 milímetros e máxima de " & lblLargProd.Caption & " milímetros)"
            lblCompProd.Caption = Trim(.Range("E" & CStr(cmbCodProd.ListIndex + 2)).Value)
            lblCodBarProd.Caption = Trim(.Range("A" & CStr(cmbCodProd.ListIndex + 2)).Value)
            txtAI10.Text = Trim(.Range("H" & CStr(cmbCodProd.ListIndex + 2)).Text)
            Select Case (Len(Trim(CStr(.Range("A" & CStr(cmbCodProd.ListIndex + 2)).Value))))
                Case 8
                    lblCodBarDescProd.Caption = "GTIN-8 ou antigo EAN-8" & "-" & strNomePais
                Case 12
                    lblCodBarDescProd.Caption = "GTIN-12 ou antigo UPC-A (América do Norte)"
                Case 13
                    lblCodBarDescProd.Caption = "GTIN-13 ou antigo EAN-13 (" & strNomePais & ")"
                Case 14
                    lblCodBarDescProd.Caption = "GTIN-14 ou antigo DUN-14 (" & strNomePais & ")"
                Case Is > 14
                    MsgBox "Código de barras com numeração incorreta (acima de 14 dígitos)!", vbExclamation + vbOKOnly
            End Select
            .Range("B" & CStr(cmbCodProd.ListIndex + 2)).Select
            If (LCase(.Range("G" & CStr(cmbCodProd.ListIndex + 2)).Text) = "sim") Then
                txtLargCorte.Locked = False: txtLargCorte.Enabled = True
                txtComp.Locked = False: txtComp.Enabled = True
                txtLargCorte.SetFocus
            End If
        End With
        Set WS_1 = Nothing
        txtLargCorte.Text = lblLargProd.Caption
        txtComp.Text = lblCompProd.Caption
        txtDiamInt.Text = "76" ': txtDiamInt.Locked = False: txtDiamInt.Enabled = True
        txtDirEnr.Text = "9" ': txtDirEnr.Locked = False: txtDirEnr.Enabled = True
        txtEmen.Text = "0" ': txtEmen.Locked = False: txtEmen.Enabled = True
        Monta_CodGS1_128
    End If
End Sub

Private Sub cmbFab_Change()
    If (cmbFab.ListIndex < 0) Then
        Exit Sub
    Else
        LimparResultados
        cmbCodProd.Enabled = True
        Dim i%
        Dim Range_1 As Range
        Dim WS_1 As Worksheet: Set WS_1 = Worksheets(cmbFab.Text)
        With WS_1
            .Select
            i = 2
            Do While (.Range("B" & CStr(i)).Value <> "") 'And .Range("C" & CStr(i)).Interior.Color = 65535
                cmbCodProd.AddItem Trim(CStr(.Range("B" & CStr(i)).Value)) & " - " & Trim(CStr(.Range("D" & CStr(i)).Value)) & "mm X " & Trim(CStr(.Range("E" & CStr(i)).Value)) & "m " & IIf(LCase(.Range("G" & CStr(i)).Text) = "sim", "(Log)", "")
                'cmbCodProd.AddItem Trim(CStr(.Range("B" & CStr(i)).Value)) & " " & Trim(CStr(.Range("D" & CStr(i)).Value)) & "X" & Trim(CStr(.Range("E" & CStr(i)).Value)) & IIf(LCase(.Range("G" & CStr(i)).Text) = "sim", " (Log)", "")
                i = i + 1
            Loop
            If (cmbCodProd.ListCount > 0) Then cmbCodProd.ListIndex = -1
        End With
        Set WS_1 = Nothing
    End If
End Sub

Private Sub cmdSair_Click()
    If (Len(txtAI10.Text) > 0) Then Salva_Numero_Lote
    Unload frmFracionamento 'Unload Me
End Sub

Private Sub txtAI10_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
    If (cmbCodProd.ListIndex = -1) Then Exit Sub
    If Not (Valida_Campos) Then
        txtLargCorte.SetFocus
        Exit Sub
    Else
        Monta_CodGS1_128
        Salva_Numero_Lote
    End If
End Sub

Private Sub txtAI10_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If Not (IsNumeric(Chr(KeyAscii))) Then
        MsgBox "Neste campo apenas são permitidos valores numéricos!", vbExclamation + vbOKOnly
        KeyAscii = 0
    End If
End Sub

Private Sub txtComp_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
    If (cmbCodProd.ListIndex = -1) Then Exit Sub
    txtGS1_128.Text = ""
    txtLeitGS1_128.Text = ""
    txtAI01.Text = ""
    txtAI8001.Text = ""
    If Not (Valida_Campos) Then
        txtComp.SetFocus
        Exit Sub
    Else
        If (Len(txtDiamInt.Text) = 0) And (Len(txtDirEnr.Text) = 0) And (Len(txtEmen.Text) = 0) Then
            txtDiamInt.Text = "76"
            txtDirEnr.Text = "9"
            txtEmen.Text = "0"
        End If
        Monta_CodGS1_128
    End If
End Sub

Private Sub txtComp_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If Not (IsNumeric(Chr(KeyAscii))) Then
        MsgBox "Neste campo apenas são permitidos valores numéricos!", vbExclamation + vbOKOnly
        KeyAscii = 0
    End If
End Sub

Private Sub txtLargCorte_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
    If (cmbCodProd.ListIndex = -1) Then Exit Sub
    txtGS1_128.Text = ""
    txtLeitGS1_128.Text = ""
    txtAI01.Text = ""
    txtAI8001.Text = ""
    If Not (Valida_Campos) Then
        txtLargCorte.SetFocus
        Exit Sub
    Else
        If (Len(txtComp.Text) = 0) And (Len(txtDiamInt.Text) = 0) And (Len(txtDirEnr.Text) = 0) And (Len(txtEmen.Text) = 0) Then
            txtComp.Text = lblCompProd.Caption: txtComp.Locked = False: txtComp.Enabled = True
            txtDiamInt.Text = "76"
            txtDirEnr.Text = "9"
            txtEmen.Text = "0"
        End If
        Monta_CodGS1_128
    End If
End Sub

Private Sub txtLargCorte_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If Not (IsNumeric(Chr(KeyAscii))) Then
        MsgBox "Neste campo apenas são permitidos valores numéricos!", vbExclamation + vbOKOnly
        KeyAscii = 0
    End If
End Sub

Private Sub UserForm_Activate()
    Dim WB_1 As Workbook: Set WB_1 = Workbooks("Fitas_v6.xls")
    Dim WS_1 As Worksheet
    With WB_1
        cmbFab.Clear
        For Each WS_1 In WB_1.Worksheets
            cmbFab.AddItem WS_1.Name
        Next
        cmbFab.TextAlign = fmTextAlignLeft
        cmbFab.ListIndex = -1
        cmbFab.Enabled = True
    End With
    Set WB_1 = Nothing
    LimparResultados
End Sub

Private Sub LimparResultados()
    With frmFracionamento
        .cmbCodProd.Clear
        .cmbCodProd.TextAlign = fmTextAlignLeft
        .cmbCodProd.ListIndex = -1
        .cmbCodProd.Enabled = False
        .lblCodBarDescProd.Caption = ""
        .lblCodBarProd.Caption = ""
        .lblCompProd.Caption = ""
        .lblDescProd.Caption = ""
        .lblLargProd.Caption = ""
        .lblLargCorte.Caption = "milímetros (largura mínima de 5 milímetros)"
        .txtLargCorte.Text = "": .txtLargCorte.Locked = True: .txtLargCorte.Enabled = False: .txtLargCorte.MaxLength = 4
        .txtComp.Text = "": .txtComp.Locked = True: .txtComp.Enabled = False: txtComp.MaxLength = 5
        .txtDiamInt.Text = "": .txtDiamInt.Locked = True: .txtDiamInt.Enabled = False: .txtDiamInt.MaxLength = 3
        .txtDirEnr.Text = "": .txtDirEnr.Locked = True: .txtDirEnr.Enabled = False: .txtDirEnr.MaxLength = 1
        .txtEmen.Text = "": .txtEmen.Locked = True: .txtEmen.Enabled = False: .txtEmen.MaxLength = 1
        .txtAI01.Text = "": .txtAI01.Locked = True: .txtAI01.Enabled = False
        .txtAI10.Text = "": .txtAI10.Locked = True: .txtAI10.Enabled = False: .txtAI10.MaxLength = 16
        .txtAI8001.Text = "": .txtAI8001.Locked = True: .txtAI8001.Enabled = False
        .txtGS1_128.Text = "": .txtGS1_128.Locked = True: .txtGS1_128.Enabled = False
        .txtLeitGS1_128.Text = "": .txtLeitGS1_128.Locked = True: .txtLeitGS1_128.Enabled = False
    End With
End Sub

Private Function Verifica_CodPais(strCodPais$) As String
    Dim x%
    Select Case (Len(strCodPais))
        Case 8
            MsgBox "Código de barras com numeração incorreta (8 dígitos)!", vbCritical + vbOKOnly
            Verifica_CodPais = "": Exit Function
        Case 12
            MsgBox "Código de barras com numeração incorreta (13 dígitos)!", vbCritical + vbOKOnly
            Verifica_CodPais = "": Exit Function
        Case Is < 13
            MsgBox "Código de barras com numeração incorreta (menor que 13 dígitos)!", vbCritical + vbOKOnly
            Verifica_CodPais = "": Exit Function
        Case Is > 14
            MsgBox "Código de barras com numeração incorreta (maior que 14 dígitos)!", vbCritical + vbOKOnly
            Verifica_CodPais = "": Exit Function
        Case 13
            x = 1
        Case 14
            x = 2
    End Select
    If (Val(Mid(strCodPais, x, 3)) > 1) And (Val(Mid(strCodPais, x, 3)) < 20) Then
        Verifica_CodPais = "EUA"
    ElseIf (Val(Mid(strCodPais, x, 3)) > 19) And (Val(Mid(strCodPais, x, 3)) < 30) Then
        Verifica_CodPais = "Distribuição restringida definido pela organização membro GS1"
    ElseIf (Val(Mid(strCodPais, x, 3)) > 29) And (Val(Mid(strCodPais, x, 3)) < 40) Then
        Verifica_CodPais = "EUA"
    ElseIf (Val(Mid(strCodPais, x, 3)) > 39) And (Val(Mid(strCodPais, x, 3)) < 50) Then
        Verifica_CodPais = "Distribuição restringida definido pela organização membro GS1"
    ElseIf (Val(Mid(strCodPais, x, 3)) > 49) And (Val(Mid(strCodPais, x, 3)) < 60) Then
        Verifica_CodPais = "Coupons"
    ElseIf (Val(Mid(strCodPais, x, 3)) > 59) And (Val(Mid(strCodPais, x, 3)) < 140) Then
        Verifica_CodPais = "EUA"
    ElseIf (Val(Mid(strCodPais, x, 3)) = 140) Then
        Verifica_CodPais = "CS Sistemas"
    ElseIf (Val(Mid(strCodPais, x, 3)) > 199) And (Val(Mid(strCodPais, x, 3)) < 300) Then
        Verifica_CodPais = "Distribuição restringida definido pela organização membro GS1"
    ElseIf (Val(Mid(strCodPais, x, 3)) > 299) And (Val(Mid(strCodPais, x, 3)) < 380) Then
        Verifica_CodPais = "França/Mônaco"
    ElseIf (Val(Mid(strCodPais, x, 3)) = 380) Then
        Verifica_CodPais = "Bulgária"
    ElseIf (Val(Mid(strCodPais, x, 3)) = 383) Then
        Verifica_CodPais = "Eslovénia"
    ElseIf (Val(Mid(strCodPais, x, 3)) = 385) Then
        Verifica_CodPais = "Croácia"
    ElseIf (Val(Mid(strCodPais, x, 3)) = 387) Then
        Verifica_CodPais = "Bósnia e Herzegovina"
    ElseIf (Val(Mid(strCodPais, x, 3)) > 399) And (Val(Mid(strCodPais, x, 3)) < 441) Then
        Verifica_CodPais = "Alemanha"
    ElseIf (Val(Mid(strCodPais, x, 3)) > 449) And (Val(Mid(strCodPais, x, 3)) < 460) Then
        Verifica_CodPais = "Japão"
    ElseIf (Val(Mid(strCodPais, x, 3)) > 489) And (Val(Mid(strCodPais, x, 3)) < 500) Then
        Verifica_CodPais = "Japão"
    ElseIf (Val(Mid(strCodPais, x, 3)) > 459) And (Val(Mid(strCodPais, x, 3)) < 470) Then
        Verifica_CodPais = "Rússia"
    ElseIf (Val(Mid(strCodPais, x, 3)) = 470) Then
        Verifica_CodPais = "Quirguistão"
    ElseIf (Val(Mid(strCodPais, x, 3)) = 471) Then
        Verifica_CodPais = "Ilha de Taiwan"
    ElseIf (Val(Mid(strCodPais, x, 3)) = 474) Then
        Verifica_CodPais = "Estônia"
    ElseIf (Val(Mid(strCodPais, x, 3)) = 475) Then
        Verifica_CodPais = "Letônia"
    ElseIf (Val(Mid(strCodPais, x, 3)) = 476) Then
        Verifica_CodPais = "Azerbaijão"
    ElseIf (Val(Mid(strCodPais, x, 3)) = 477) Then
        Verifica_CodPais = "Lituânia"
    ElseIf (Val(Mid(strCodPais, x, 3)) = 478) Then
        Verifica_CodPais = "Usbequistão"
    ElseIf (Val(Mid(strCodPais, x, 3)) = 479) Then
        Verifica_CodPais = "Sri Lanka"
    ElseIf (Val(Mid(strCodPais, x, 3)) = 480) Then
        Verifica_CodPais = "Filipinas"
    ElseIf (Val(Mid(strCodPais, x, 3)) = 481) Then
        Verifica_CodPais = "Bielorrússia"
    ElseIf (Val(Mid(strCodPais, x, 3)) = 482) Then
        Verifica_CodPais = "Ucrânia"
    ElseIf (Val(Mid(strCodPais, x, 3)) = 484) Then
        Verifica_CodPais = "Moldávia"
    ElseIf (Val(Mid(strCodPais, x, 3)) = 485) Then
        Verifica_CodPais = "Armênia"
    ElseIf (Val(Mid(strCodPais, x, 3)) = 486) Then
        Verifica_CodPais = "Geórgia"
    ElseIf (Val(Mid(strCodPais, x, 3)) = 487) Then
        Verifica_CodPais = "Cazaquistão"
    ElseIf (Val(Mid(strCodPais, x, 3)) = 489) Then
        Verifica_CodPais = "Hong Kong"
    ElseIf (Val(Mid(strCodPais, x, 3)) > 449) And (Val(Mid(strCodPais, x, 3)) < 510) Then
        Verifica_CodPais = "Reino Unido"
    ElseIf (Val(Mid(strCodPais, x, 3)) = 520) Then
        Verifica_CodPais = "Grécia"
    ElseIf (Val(Mid(strCodPais, x, 3)) = 528) Then
        Verifica_CodPais = "Líbano"
    ElseIf (Val(Mid(strCodPais, x, 3)) = 529) Then
        Verifica_CodPais = "Chipre"
    ElseIf (Val(Mid(strCodPais, x, 3)) = 530) Then
        Verifica_CodPais = "Albânia"
    ElseIf (Val(Mid(strCodPais, x, 3)) = 531) Then
        Verifica_CodPais = "República da Macedônia"
    ElseIf (Val(Mid(strCodPais, x, 3)) = 535) Then
        Verifica_CodPais = "Malta"
    ElseIf (Val(Mid(strCodPais, x, 3)) = 539) Then
        Verifica_CodPais = "República da Irlanda"
    ElseIf (Val(Mid(strCodPais, x, 3)) > 539) And (Val(Mid(strCodPais, x, 3)) < 550) Then
        Verifica_CodPais = "Bélgica/Luxemburgo"
    ElseIf (Val(Mid(strCodPais, x, 3)) = 560) Then
        Verifica_CodPais = "Portugal"
    ElseIf (Val(Mid(strCodPais, x, 3)) = 569) Then
        Verifica_CodPais = "Islândia"
    ElseIf (Val(Mid(strCodPais, x, 3)) > 569) And (Val(Mid(strCodPais, x, 3)) < 580) Then
        Verifica_CodPais = "Dinamarca/Ilhas Feroé/Groenlândia"
    ElseIf (Val(Mid(strCodPais, x, 3)) = 590) Then
        Verifica_CodPais = "Polónia"
    ElseIf (Val(Mid(strCodPais, x, 3)) = 594) Then
        Verifica_CodPais = "Romênia"
    ElseIf (Val(Mid(strCodPais, x, 3)) = 599) Then
        Verifica_CodPais = "Hungria"
    ElseIf (Val(Mid(strCodPais, x, 3)) > 599) And (Val(Mid(strCodPais, x, 3)) < 602) Then
        Verifica_CodPais = "África do Sul"
    ElseIf (Val(Mid(strCodPais, x, 3)) = 603) Then
        Verifica_CodPais = "Gana"
    ElseIf (Val(Mid(strCodPais, x, 3)) = 608) Then
        Verifica_CodPais = "Bahrein"
    ElseIf (Val(Mid(strCodPais, x, 3)) = 609) Then
        Verifica_CodPais = "lhas Maurício"
    ElseIf (Val(Mid(strCodPais, x, 3)) = 611) Then
        Verifica_CodPais = "Marrocos"
    ElseIf (Val(Mid(strCodPais, x, 3)) = 613) Then
        Verifica_CodPais = "Argélia"
    ElseIf (Val(Mid(strCodPais, x, 3)) = 616) Then
        Verifica_CodPais = "Quênia"
    ElseIf (Val(Mid(strCodPais, x, 3)) = 618) Then
        Verifica_CodPais = "Costa do Marfim"
    ElseIf (Val(Mid(strCodPais, x, 3)) = 619) Then
        Verifica_CodPais = "Tunísia"
    ElseIf (Val(Mid(strCodPais, x, 3)) = 621) Then
        Verifica_CodPais = "Síria"
    ElseIf (Val(Mid(strCodPais, x, 3)) = 622) Then
        Verifica_CodPais = "Egito"
    ElseIf (Val(Mid(strCodPais, x, 3)) = 624) Then
        Verifica_CodPais = "Líbia"
    ElseIf (Val(Mid(strCodPais, x, 3)) = 625) Then
        Verifica_CodPais = "Jordânia"
    ElseIf (Val(Mid(strCodPais, x, 3)) = 626) Then
        Verifica_CodPais = "Irã"
    ElseIf (Val(Mid(strCodPais, x, 3)) = 627) Then
        Verifica_CodPais = "Kuwait"
    ElseIf (Val(Mid(strCodPais, x, 3)) = 628) Then
        Verifica_CodPais = "Arábia Saudita"
    ElseIf (Val(Mid(strCodPais, x, 3)) = 629) Then
        Verifica_CodPais = "Emirados Árabes Unidos"
    ElseIf (Val(Mid(strCodPais, x, 3)) > 639) And (Val(Mid(strCodPais, x, 3)) < 650) Then
        Verifica_CodPais = "Finlândia"
    ElseIf (Val(Mid(strCodPais, x, 3)) > 689) And (Val(Mid(strCodPais, x, 3)) < 700) Then
        Verifica_CodPais = "República Popular da China"
    ElseIf (Val(Mid(strCodPais, x, 3)) > 699) And (Val(Mid(strCodPais, x, 3)) < 710) Then
        Verifica_CodPais = "Noruega"
    ElseIf (Val(Mid(strCodPais, x, 3)) = 729) Then
        Verifica_CodPais = "Israel"
    ElseIf (Val(Mid(strCodPais, x, 3)) > 729) And (Val(Mid(strCodPais, x, 3)) < 740) Then
        Verifica_CodPais = "Suécia"
    ElseIf (Val(Mid(strCodPais, x, 3)) = 740) Then
        Verifica_CodPais = "Guatemala"
    ElseIf (Val(Mid(strCodPais, x, 3)) = 741) Then
        Verifica_CodPais = "El Salvador"
    ElseIf (Val(Mid(strCodPais, x, 3)) = 742) Then
        Verifica_CodPais = "Honduras"
    ElseIf (Val(Mid(strCodPais, x, 3)) = 743) Then
        Verifica_CodPais = "Nicarágua"
    ElseIf (Val(Mid(strCodPais, x, 3)) = 744) Then
        Verifica_CodPais = "Costa Rica"
    ElseIf (Val(Mid(strCodPais, x, 3)) = 745) Then
        Verifica_CodPais = "Panamá"
    ElseIf (Val(Mid(strCodPais, x, 3)) = 746) Then
        Verifica_CodPais = "República Dominicana"
    ElseIf (Val(Mid(strCodPais, x, 3)) = 750) Then
        Verifica_CodPais = "México"
    ElseIf (Val(Mid(strCodPais, x, 3)) > 753) And (Val(Mid(strCodPais, x, 3)) < 756) Then
        Verifica_CodPais = "Canadá"
    ElseIf (Val(Mid(strCodPais, x, 3)) = 759) Then
        Verifica_CodPais = "Venezuela"
    ElseIf (Val(Mid(strCodPais, x, 3)) > 759) And (Val(Mid(strCodPais, x, 3)) < 770) Then
        Verifica_CodPais = "Suíça/Liechtenstein"
    ElseIf (Val(Mid(strCodPais, x, 3)) = 770) Then
        Verifica_CodPais = "Colômbia"
    ElseIf (Val(Mid(strCodPais, x, 3)) = 773) Then
        Verifica_CodPais = "Uruguai"
    ElseIf (Val(Mid(strCodPais, x, 3)) = 775) Then
        Verifica_CodPais = "Peru"
    ElseIf (Val(Mid(strCodPais, x, 3)) = 777) Then
        Verifica_CodPais = "Bolívia"
    ElseIf (Val(Mid(strCodPais, x, 3)) = 779) Then
        Verifica_CodPais = "Argentina"
    ElseIf (Val(Mid(strCodPais, x, 3)) = 780) Then
        Verifica_CodPais = "Chile"
    ElseIf (Val(Mid(strCodPais, x, 3)) = 784) Then
        Verifica_CodPais = "Paraguai"
    ElseIf (Val(Mid(strCodPais, x, 3)) = 786) Then
        Verifica_CodPais = "Equador"
    ElseIf (Val(Mid(strCodPais, x, 3)) > 788) And (Val(Mid(strCodPais, x, 3)) < 791) Then
        Verifica_CodPais = "Brasil"
    ElseIf (Val(Mid(strCodPais, x, 3)) > 799) And (Val(Mid(strCodPais, x, 3)) < 840) Then
        Verifica_CodPais = "Itália/San Marino/Vaticano"
    ElseIf (Val(Mid(strCodPais, x, 3)) > 839) And (Val(Mid(strCodPais, x, 3)) < 850) Then
        Verifica_CodPais = "Espanha/Andorra"
    ElseIf (Val(Mid(strCodPais, x, 3)) = 850) Then
        Verifica_CodPais = "Cuba"
    ElseIf (Val(Mid(strCodPais, x, 3)) = 858) Then
        Verifica_CodPais = "Eslováquia"
    ElseIf (Val(Mid(strCodPais, x, 3)) = 859) Then
        Verifica_CodPais = "República Checa"
    ElseIf (Val(Mid(strCodPais, x, 3)) = 860) Then
        Verifica_CodPais = "Sérvia e Montenegro"
    ElseIf (Val(Mid(strCodPais, x, 3)) = 865) Then
        Verifica_CodPais = "Mongólia"
    ElseIf (Val(Mid(strCodPais, x, 3)) = 867) Then
        Verifica_CodPais = "Coreia do Norte"
    ElseIf (Val(Mid(strCodPais, x, 3)) = 869) Then
        Verifica_CodPais = "Turquia"
    ElseIf (Val(Mid(strCodPais, x, 3)) > 869) And (Val(Mid(strCodPais, x, 3)) < 880) Then
        Verifica_CodPais = "Holanda"
    ElseIf (Val(Mid(strCodPais, x, 3)) = 880) Then
        Verifica_CodPais = "Coreia do Sul"
    ElseIf (Val(Mid(strCodPais, x, 3)) = 884) Then
        Verifica_CodPais = "Cambodja"
    ElseIf (Val(Mid(strCodPais, x, 3)) = 885) Then
        Verifica_CodPais = "Tailândia"
    ElseIf (Val(Mid(strCodPais, x, 3)) = 888) Then
        Verifica_CodPais = "Singapura"
    ElseIf (Val(Mid(strCodPais, x, 3)) = 890) Then
        Verifica_CodPais = "Índia"
    ElseIf (Val(Mid(strCodPais, x, 3)) = 893) Then
        Verifica_CodPais = "Vietnam"
    ElseIf (Val(Mid(strCodPais, x, 3)) = 899) Then
        Verifica_CodPais = "Indonésia"
    ElseIf (Val(Mid(strCodPais, x, 3)) > 899) And (Val(Mid(strCodPais, x, 3)) < 920) Then
        Verifica_CodPais = "Áustria"
    ElseIf (Val(Mid(strCodPais, x, 3)) > 929) And (Val(Mid(strCodPais, x, 3)) < 940) Then
        Verifica_CodPais = "Austrália"
    ElseIf (Val(Mid(strCodPais, x, 3)) > 939) And (Val(Mid(strCodPais, x, 3)) < 950) Then
        Verifica_CodPais = "Nova Zelândia"
    ElseIf (Val(Mid(strCodPais, x, 3)) = 950) Then
        Verifica_CodPais = "GS1 Global Office"
    ElseIf (Val(Mid(strCodPais, x, 3)) = 955) Then
        Verifica_CodPais = "Malásia"
    ElseIf (Val(Mid(strCodPais, x, 3)) = 958) Then
        Verifica_CodPais = "Macau"
    ElseIf (Val(Mid(strCodPais, x, 3)) = 977) Then
        Verifica_CodPais = "Publicações periódicas seriadas (ISSN)"
    ElseIf (Val(Mid(strCodPais, x, 3)) = 978) Then
        Verifica_CodPais = "Bookland (ISBN)"
    ElseIf (Val(Mid(strCodPais, x, 3)) = 979) Then
        Verifica_CodPais = "Formalmente usado para pautas de música"
    ElseIf (Val(Mid(strCodPais, x, 3)) = 980) Then
        Verifica_CodPais = "Refund receipts"
    ElseIf (Val(Mid(strCodPais, x, 3)) > 980) And (Val(Mid(strCodPais, x, 3)) < 983) Then
        Verifica_CodPais = "Coupons e meios de pagamento"
    ElseIf (Val(Mid(strCodPais, x, 3)) > 989) And (Val(Mid(strCodPais, x, 3)) < 1000) Then
        Verifica_CodPais = "Coupons"
    Else
        'MsgBox "Cadastro de código de barras incorreto!", vbExclamation + vbOKOnly
        Verifica_CodPais = "Indeterminado"
    End If
End Function

Private Sub Monta_CodGS1_128()
    Dim strAI01$, strAI10$, strAI8001$
    Dim i%
    strAI01 = ""
    txtAI01.Text = ""
    txtAI8001.Text = ""
    Select Case (Len(Trim(Str(lblCodBarProd.Caption))))
        Case 8
            strAI01 = "000000" & Trim(Str(lblCodBarProd.Caption))
        Case 12
            strAI01 = "00" & Trim(Str(lblCodBarProd.Caption))
        Case 13
            strAI01 = "0" & Trim(Str(lblCodBarProd.Caption))
        Case 14
            strAI01 = Trim(Str(lblCodBarProd.Caption))
    End Select
    txtAI01.Value = strAI01
    strAI8001 = ""
    For i = 1 To (4 - Len(Trim(CStr(Val(txtLargCorte.Text)))))
        strAI8001 = strAI8001 & "0"
    Next i
    strAI8001 = strAI8001 & Trim(CStr(Val(txtLargCorte.Text)))
    For i = 1 To (5 - Len(Trim(CStr(Val(txtComp.Text)))))
        strAI8001 = strAI8001 & "0"
    Next i
    strAI8001 = strAI8001 & Trim(CStr(Val(txtComp.Text)))
    For i = 1 To (3 - Len(Trim(CStr(Val(txtDiamInt.Text)))))
        strAI8001 = strAI8001 & "0"
    Next i
    strAI8001 = strAI8001 & Trim(CStr(Val(txtDiamInt.Text)))
    strAI8001 = strAI8001 & Trim(CStr(Val(txtDirEnr.Text)))
    strAI8001 = strAI8001 & Trim(CStr(Val(txtEmen.Text)))
    txtAI8001.Value = strAI8001
    txtAI8001.Enabled = True
    txtAI01.Enabled = True
    txtAI10.Enabled = True: txtAI10.Locked = False
    If (Len(txtAI10.Text) > 0) Then strAI10 = txtAI10.Text
    txtGS1_128.Text = "(01)" & strAI01 & "(8001)" & strAI8001 & IIf(Len(strAI10) > 0, "(10)" & strAI10, ""): txtGS1_128.Enabled = True
    txtLeitGS1_128.Text = "01" & strAI01 & "8001" & strAI8001 & IIf(Len(strAI10) > 0, "10" & strAI10, ""): txtLeitGS1_128.Enabled = True
End Sub

Private Function Valida_Campos() As Boolean
    If (txtLargCorte.Value = "") Or (Val(txtLargCorte.Value) < 4) Or (Val(txtLargCorte.Value) > Val(lblLargProd.Caption)) Or (Len(Trim(txtLargCorte.Value)) > 4) Then
        Valida_Campos = False
    ElseIf (txtComp.Value = "") Or (Val(txtComp.Value) > Val(lblCompProd.Caption)) Or (Len(Trim(txtComp.Value)) > 5) Or (Val(txtComp.Value) = 0) Then
        Valida_Campos = False
    ElseIf (txtDiamInt.Value = "") Or (Val(txtDiamInt.Value) < 25.4) Then
        Valida_Campos = False
    ElseIf (txtDirEnr.Value = "") Or (Val(txtDirEnr.Value) <> 0 And Val(txtDirEnr.Value) <> 1 And Val(txtDirEnr.Value) <> 9) Then
        Valida_Campos = False
    ElseIf (txtEmen.Value = "") Or (Val(txtDirEnr.Value) < 0) Or (Val(txtDirEnr.Value) > 9) Then
        Valida_Campos = False
    Else
        Valida_Campos = True
    End If
End Function

Private Sub Salva_Numero_Lote()
    Dim WB_1 As Workbook: Set WB_1 = Workbooks("Fitas_v6.xls")
    Dim WS_1 As Worksheet: Set WS_1 = Worksheets(cmbFab.Text)
    With WS_1
        .Unprotect "neo2018"
        .Range("H" & CStr(cmbCodProd.ListIndex + 2)).Value = Trim(txtAI10.Text)
        .Protect "neo2018", DrawingObjects:=True, Contents:=True, Scenarios:=True
    End With
    Set WS_1 = Nothing
    WB_1.Save
    Set WB_1 = Nothing
End Sub

'Private Sub txtDiamInt_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
'    If (cmbCodProd.ListIndex = -1) Then Exit Sub
'    txtGS1_128.Text = ""
'    txtLeitGS1_128.Text = ""
'    txtAI01.Text = ""
'    txtAI8001.Text = ""
'    If Not (Valida_Campos) Then
'        txtDiamInt.SetFocus
'        Exit Sub
'    Else
'        If (Len(txtDirEnr.Text) = 0) And (Len(txtEmen.Text) = 0) Then
'            txtDirEnr.Text = "9"
'            txtEmen.Text = "0"
'        End If
'        Monta_CodGS1_128
'    End If
'End Sub

'Private Sub txtDiamInt_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
'    If Not (IsNumeric(Chr(KeyAscii))) Then
'        MsgBox "Neste campo apenas são permitidos valores numéricos!", vbExclamation + vbOKOnly
'        KeyAscii = 0
'    End If
'End Sub

'Private Sub txtDirEnr_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
'    If (cmbCodProd.ListIndex = -1) Then Exit Sub
'    txtGS1_128.Text = ""
'    txtLeitGS1_128.Text = ""
'    txtAI01.Text = ""
'    txtAI8001.Text = ""
'    If Not (Valida_Campos) Then
'        txtDirEnr.SetFocus
'        Exit Sub
'    Else
'        If (Len(txtEmen.Text) = 0) Then txtEmen.Text = "0"
'        Monta_CodGS1_128
'    End If
'End Sub

'Private Sub txtDirEnr_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
'    If Not (IsNumeric(Chr(KeyAscii))) Then
'        MsgBox "Neste campo apenas são permitidos valores numéricos!", vbExclamation + vbOKOnly
'        KeyAscii = 0
'    End If
'End Sub

'Private Sub txtEmen_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
'    If (cmbCodProd.ListIndex = -1) Then Exit Sub
'    txtGS1_128.Text = ""
'    txtLeitGS1_128.Text = ""
'    txtAI01.Text = ""
'    txtAI8001.Text = ""
'    If Not (Valida_Campos) Then
'        txtEmen.SetFocus
'        Exit Sub
'    Else
'        Monta_CodGS1_128
'    End If
'End Sub

'Private Sub txtEmen_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
'    If Not (IsNumeric(Chr(KeyAscii))) Then
'        MsgBox "Neste campo apenas são permitidos valores numéricos!", vbExclamation + vbOKOnly
'        KeyAscii = 0
'    End If
'End Sub
