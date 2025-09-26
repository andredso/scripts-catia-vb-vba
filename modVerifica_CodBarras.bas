Attribute VB_Name = "modVerifica_CodBarras"
Sub Cria_EAN13()
    Dim strCol$
    Dim lngRow&
    strCol = UCase(InputBox("Informe a Coluna. Exemplo: B", "Consolidar EAN 13", ""))
    lngRow = CLng(InputBox("Informe a Linha. Exemplo: 2", "Consolidar EAN 13", "0")) 'cdate IsDate
    If (strCol = "") Or (IsNumeric(strCol)) Or (IsEmpty(strCol)) Or (IsNull(strCol)) Then
        Exit Sub
    ElseIf Not (IsNumeric(lngRow)) Or (IsEmpty(lngRow)) Or (IsNull(lngRow)) Then
        Exit Sub
    Else
        Dim WS_1 As Worksheet: Set WS_1 = Application.ActiveSheet
        Dim strCodVerificado$
        With WS_1
            Do While (.Range(strCol & CStr(lngRow)).Value <> "")
                .Range(strCol & CStr(lngRow)).Select
                If (Len(Trim(CStr(.Range(strCol & CStr(lngRow)).Text))) = 12) Then
                    .Range(strCol & CStr(lngRow)).NumberFormat = "@"
                    .Range(strCol & CStr(lngRow)).Value = DVEAN13(CStr(.Range(strCol & CStr(lngRow)).Value))
                End If
                lngRow = lngRow + 1
            Loop
            .Range(strCol & CStr(lngRow)).EntireColumn.AutoFit
            .Range(strCol & CStr(2)).Select
        End With
        Set WS_1 = Nothing
    End If
End Sub
Sub Consolidar_EAN13()
    Dim strCol$
    Dim lngRow&
    strCol = UCase(InputBox("Informe a Coluna. Exemplo: B", "Consolidar EAN 13", ""))
    lngRow = CLng(InputBox("Informe a Linha. Exemplo: 2", "Consolidar EAN 13", "0")) 'cdate IsDate
    If (strCol = "") Or (IsNumeric(strCol)) Or (IsEmpty(strCol)) Or (IsNull(strCol)) Then
        Exit Sub
    ElseIf Not (IsNumeric(lngRow)) Or (IsEmpty(lngRow)) Or (IsNull(lngRow)) Then
        Exit Sub
    Else
        Dim WS_1 As Worksheet: Set WS_1 = Application.ActiveSheet
        Dim strCodVerificado$
        With WS_1
            Do While (.Range(strCol & CStr(lngRow)).Value <> "")
                If (Len(Trim(CStr(.Range(strCol & CStr(lngRow)).Text))) <> 13) Then
                    .Range(strCol & CStr(lngRow)).Font.Color = -16776961
                Else
                    strCodVerificado = DVEAN13(Mid(Trim(CStr(.Range(strCol & CStr(lngRow)).Value)), 1, 12))
                    If (Trim(CStr(.Range(strCol & CStr(lngRow)).Value)) <> strCodVerificado) Then
                        .Range(strCol & CStr(lngRow)).NumberFormat = "@"
                        .Range(strCol & CStr(lngRow)).Value = .Range(strCol & CStr(lngRow)).Value & " (" & strCodVerificado & ")"
                        .Range(strCol & CStr(lngRow)).Font.Color = -16776961
                    Else
                        .Range(strCol & CStr(lngRow)).Font.Color = -65536 '.ColorIndex = xlAutomatic
                    End If
                End If
                .Range(strCol & CStr(lngRow)).Font.TintAndShade = 0
                lngRow = lngRow + 1
            Loop
            .Range(strCol & CStr(lngRow)).EntireColumn.AutoFit
        End With
        Set WS_1 = Nothing
    End If
End Sub

Sub Consolidar_DUN14()
    Dim strCol$
    Dim lngRow&
    strCol = UCase(InputBox("Informe a Coluna. Exemplo: B", "Consolidar DUN 14", ""))
    lngRow = CLng(InputBox("Informe a Linha. Exemplo: 2", "Consolidar DUN 14", "0"))
    If (strCol = "") Or (IsNumeric(strCol)) Or (IsEmpty(strCol)) Or (IsNull(strCol)) Then
        Exit Sub
    ElseIf Not (IsNumeric(lngRow)) Or (IsEmpty(lngRow)) Or (IsNull(lngRow)) Then
        Exit Sub
    Else
        Dim WS_1 As Worksheet: Set WS_1 = Application.ActiveSheet
        Dim strCodVerificado$
        With WS_1
            Do While (.Range(strCol & CStr(lngRow)).Value <> "")
                If (Len(Trim(CStr(.Range(strCol & CStr(lngRow)).Text))) <> 14) Then
                    .Range(strCol & CStr(lngRow)).Font.Color = -16776961
                Else
                    strCodVerificado = DVEAN14(Trim(CStr(.Range(strCol & CStr(lngRow)).Value)))
                    If (Trim(CStr(.Range(strCol & CStr(lngRow)).Value)) <> strCodVerificado) Then
                        .Range(strCol & CStr(lngRow)).NumberFormat = "@"
                        .Range(strCol & CStr(lngRow)).Value = .Range(strCol & CStr(lngRow)).Value & " (" & strCodVerificado & ")"
                        .Range(strCol & CStr(lngRow)).Font.Color = -16776961
                    Else
                        .Range(strCol & CStr(lngRow)).Font.Color = -65536 '.ColorIndex = xlAutomatic
                    End If
                End If
                .Range(strCol & CStr(lngRow)).Font.TintAndShade = 0
                lngRow = lngRow + 1
            Loop
            .Range(strCol & CStr(lngRow)).EntireColumn.AutoFit
        End With
        Set WS_1 = Nothing
    End If
End Sub

Function DVEAN13(ByVal strCod$) As String
    Dim intPar%, intImpar%, intSomaPar%, intSomaImpar%, intTotalSoma%, intDv%, intI%
    If (Len(strCod) <> 12) Or (Not IsNumeric(strCod)) Or (IsEmpty(strCod)) Or (IsNull(strCod)) Then
        MsgBox "O Código para cálculo do dígito verificado está errado. Favor corrigir.", vbCritical, "Erro"
        Exit Function
    End If
    intSomaPar = 0
    intSomaImpar = 0
    intTotalSoma = 0
    intDv = 0
    For intI = 2 To 12 Step 2
        intPar = CInt(Mid(strCod, intI, 1))
        intSomaPar = intSomaPar + intPar
    Next intI
    For intI = 1 To 11 Step 2
        intImpar = CInt(Mid(strCod, intI, 1))
        intSomaImpar = intSomaImpar + intImpar
    Next intI
    intSomaPar = intSomaPar * 3
    intTotalSoma = intSomaPar + intSomaImpar
    Do While (intTotalSoma Mod 10 <> 0)
        intDv = intDv + 1
        intTotalSoma = intTotalSoma + 1
    Loop
    DVEAN13 = strCod & CStr(intDv)
End Function

Function DVEAN14(ByVal strCod$) As String
    Dim intNumeros(13) As Integer
    Dim intI%: intI = 0
    Dim intSoma%: intSoma = 0
    Dim intResultado%: intResultado = 0
    Dim intRep%: intRep = Len(strCod) - 1
    Dim intNumeroFinal%: intNumeroFinal = 0
    Dim intMultiplicacao%: intMultiplicacao = 0
    Dim bolValidacao As Boolean
    If intRep = 13 Then
        For intI = 1 To Len(strCod) - 1 Step 1
            intNumeros(intI) = Mid(strCod, intI, 1)
            intMultiplicacao = intMultiplicacao + intNumeros(intI) * IIf(intI Mod 2 = 0, 1, 3)
        Next intI
        intResultado = intMultiplicacao Mod 10
        If intResultado = 0 Then
            intNumeroFinal = 0
        Else
            intNumeroFinal = 10 - intResultado
        End If
        If intNumeroFinal = Right(strCod, 1) Then 'Microsoft.VisualBasic.Right(strCodigo, 1)
            bolValidacao = True
        Else
            bolValidacao = False
        End If
        DVEAN14 = Mid(strCod, 1, 13) & Trim(CStr(intNumeroFinal))
    End If
End Function

'Função para cálculo de dígito verificador na impressão de etiquetas com código de barra padrão EAN13
'Recebe:- Código cujo dígito precisa ser calculado
'Fornece:- Código completo já com o dígito calculado
'Autor:- Mauro Possatto (Fórum Access)
'Data:- 28/10/97
'Alterações:-
'
'Sub verifica_EAN13()
'    Dim strEAN13Informado$, strEAN13Verificado$
'    strEAN13Informado = InputBox("Informe o código EAN13 ou GTIN13 de 13 dígitos numéricos.", "Verificar EAN13(GTIN13)")
'    If (strEAN13Informado = "") Then
'        Exit Sub
'    ElseIf (Len(strEAN13Informado) <> 13) Or Not (IsNumeric(strEAN13Informado)) Or (IsEmpty(strEAN13Informado)) Or (IsNull(strEAN13Informado)) Then
'        MsgBox "Código informado incorretamente!", vbCritical + vbOKOnly, "Erro"
'    Else
'        strEAN13Verificado = DVEAN13(Mid(strEAN13Informado, 1, 12))
'        If (strEAN13Informado = strEAN13Verificado) Then
'            MsgBox "O Código EAN13 (GTIN13) " & strEAN13Informado & " está correto!", vbInformation + vbOKOnly, "Código Correto"
'        Else
'            MsgBox "O Código EAN13 (GTIN13) " & strEAN13Informado & " está incorreto!" & Chr(13) & Chr(10) & _
'            "O correto é " & strEAN13Verificado & ".", vbExclamation + vbOKOnly, "Código Correto"
'        End If
'    End If
'    Dim teste As Boolean
'    teste = dun14("07893333330097")
'    teste = ean13("7893333330097")
'End Sub

'Function SepCaracter(ByVal strCampo$) As String
'    Dim strTemp$: strTemp = " "
'    Dim intJ%
'    For intJ = 1 To Len(strCampo)
'        strTemp = strTemp & Mid(strCampo, intJ, 1) & " "
'    Next intJ
'    SepCaracter = strTemp
'End Function

'--sql-server
'
'CREATE FUNCTION dbo.fn_ValidaEAN
'    (
'    @CodeEAN BIGINT
'    )
'RETURNS INT
'AS
'    BEGIN
'        DECLARE @dig1           INT,
'                @dig2           INT,
'                @dig3           INT,
'                @dig4           INT,
'                @dig5           INT,
'                @dig6           INT,
'                @dig7           INT,
'                @dig8           INT,
'                @dig9           INT,
'                @dig10          INT,
'                @dig11          INT,
'                @dig12          INT,
'                @CodeEANc       VARCHAR(20),
'                @soma_dig_par   INT,
'                @soma_dig_impar INT,
'                @soma_dig       INT,
'                @resultado      INT,
'                @digito_ean     INT
'
'        SELECT  @CodeEANc = CAST(@CodeEAN AS VARCHAR)
'        SELECT  @dig1  = ( SUBSTRING(@CodeEANc,  1, 1) ),
'                @dig2  = ( SUBSTRING(@CodeEANc,  2, 1) ),
'                @dig3  = ( SUBSTRING(@CodeEANc,  3, 1) ),
'                @dig4  = ( SUBSTRING(@CodeEANc,  4, 1) ),
'                @dig5  = ( SUBSTRING(@CodeEANc,  5, 1) ),
'                @dig6  = ( SUBSTRING(@CodeEANc,  6, 1) ),
'                @dig7  = ( SUBSTRING(@CodeEANc,  7, 1) ),
'                @dig8  = ( SUBSTRING(@CodeEANc,  8, 1) ),
'                @dig9  = ( SUBSTRING(@CodeEANc,  9, 1) ),
'                @dig10 = ( SUBSTRING(@CodeEANc, 10, 1) ),
'                @dig11 = ( SUBSTRING(@CodeEANc, 11, 1) ),
'                @dig12 = ( SUBSTRING(@CodeEANc, 12, 1) )
'
'        SELECT  @soma_dig_par = 3 * ( @dig2 + @dig4 + @dig6 + @dig8 + @dig10 + @dig12 ),
'                @soma_dig_impar = @dig1 + @dig3 + @dig5 + @dig7 + @dig9 + @dig11
'        SELECT  @soma_dig = @soma_dig_par + @soma_dig_impar
'        SELECT  @resultado = ( ( 1 + Cast(@soma_dig / 10 AS INT) ) * 10 ) - @soma_dig
'
'        IF ( @resultado = 10 )
'            RETURN 0
'        Else
'            RETURN @resultado
'        RETURN 0
'    End
