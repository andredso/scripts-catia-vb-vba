Attribute VB_Name = "modComprovarCodigos"

Public Function dun14(ByVal strCodigo$) As Boolean
    Dim intNumeros(13) As Integer
    Dim intI%: intI = 0
    Dim intSoma%: intSoma = 0
    Dim intResultado%: intResultado = 0
    Dim intRep%: intRep = Len(strCodigo) - 1
    Dim intNumeroFinal%: intNumeroFinal = 0
    Dim intMultiplicacao%: intMultiplicacao = 0
    Dim bolValidacao As Boolean
    If intRep = 13 Then
        For intI = 1 To Len(strCodigo) - 1 Step 1
            intNumeros(intI) = Mid(strCodigo, intI, 1)
            intMultiplicacao = intMultiplicacao + intNumeros(intI) * IIf(intI Mod 2 = 0, 1, 3)
        Next intI
        intResultado = intMultiplicacao Mod 10
        If intResultado = 0 Then
            intNumeroFinal = 0
        Else
            intNumeroFinal = 10 - intResultado
        End If
        If intNumeroFinal = Right(strCodigo, 1) Then 'Microsoft.VisualBasic.Right(strCodigo, 1)
            bolValidacao = True
        Else
            bolValidacao = False
        End If
    End If
    dun14 = bolValidacao
End Function

Public Function ean13(ByVal strCodigo$) As Boolean
    Dim bolValidacao As Boolean
    Dim intSoma%: intSoma = 0
    Dim intDigito%: intDigito = 0
    Dim intI%
    Dim strEAN$: strEAN = Left(strCodigo, 12) 'Microsoft.VisualBasic.Left(strCodigo, 12)
    For intI = Len(strEAN) To 1 Step -1
        If (intI - 1 = 0) Then Exit For
        intDigito = CInt(Mid(strEAN, intI - 1, 1))
        If ((Len(strEAN) - intI + 1) Mod 2 = 0) Then
            intSoma = intSoma + intDigito * 3
        Else
            intSoma = intSoma + intDigito
        End If
    Next
    Dim intComprovacao%: intComprovacao = (10 - (intSoma Mod 10)) Mod 10
    If (CStr(intComprovacao) = Right(strCodigo, 1)) Then 'Microsoft.VisualBasic.Right(strCodigo, 1) Then
        bolValidacao = True
    Else
        bolValidacao = False
    End If
    ean13 = bolValidacao
End Function

Public Function ean8(ByVal strCodigo$) As Boolean
    Dim intNumeros(7) As Integer
    Dim intI$: intI = 0
    Dim intSoma$: intSoma = 0
    Dim intResultado$: intResultado = 0
    Dim intNumeroFinal%: intNumeroFinal = 0
    Dim intPares%: intPares = 0
    Dim intImpares%: intImpares = 0
    Dim intMultiplicacao%: intMultiplicacao = 0
    Dim bolValidacao As Boolean
    For intI = 1 To Len(strCodigo) - 1 Step 1
        intNumeros(intI) = Mid(strCodigo, intI, 1)
        If (intNumeros(intI) Mod 2 = 0) Then
            intMultiplicacao = intMultiplicacao + intNumeros(intI) * 2
        Else
            intMultiplicacao = intMultiplicacao + intNumeros(intI) * 1
        End If
    Next intI
    intResultado = (intMultiplicacao * 9) Mod 10
    If (intResultado = 0) Then
        intNumeroFinal = 0
    Else
        intNumeroFinal = 10 - intResultado
    End If
    If (intNumeroFinal = Right(strCodigo, 1)) Then 'Microsoft.VisualBasic.Right(strCodigo, 1)
        bolValidacao = True
    Else
        bolValidacao = False
    End If
    ean8 = bolValidacao
End Function

Public Function ean12(ByVal strCodigo$) As Boolean
    Dim bolValidacao As Boolean
    Dim intSoma%: intSoma = 0
    Dim intDigito%: intDigito = 0
    Dim intI%
    Dim strEAN$: strEAN = Left(strCodigo, 11) 'Microsoft.VisualBasic.Left
    For intI = Len(strEAN) To 1 Step -1
        intDigito = CInt(Mid(strEAN, intI - 1, 1)) 'Convert.ToInt32(strEAN.Substring(intI – 1, 1))
        If (Len(strEAN - intI + 1) Mod 2) = 0 Then
            intSoma = intSoma + intDigito * 3
        Else
            intSoma = intSoma + intDigito
        End If
    Next
    Dim intCheckSum%: intCheckSum = (10 - (intSoma Mod 10)) Mod 10
    If CStr(intCheckSum) = Right(strCodigo, 1) Then
        bolValidacao = True
    Else
        bolValidacao = False
    End If
    ean12 = bolValidacao
End Function

