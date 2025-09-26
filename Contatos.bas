Attribute VB_Name = "Módulo2"
Sub Macro1()
    Dim MyFile, MyPath, MyName
    Dim WB_1 As Workbook
    Dim WB_2 As Workbook
    Dim WS_1 As Worksheet
    Dim WS_2 As Worksheet
    Dim intCont1, intCont2, intCont3, intCont4 As Integer
    Dim strNome, strEMail As String
    Dim bolAchou As Boolean
    '---------------------------------------------------------'
    MyPath = "C:\Documents and Settings\orcamento\Desktop\Contatos\*.xls"
    MyName = Dir(MyPath, vbNormal)
    '---------------------------------------------------------'
    Workbooks.Add
    ChDir "C:\Documents and Settings\orcamento\Desktop\Contatos"
    ActiveWorkbook.SaveAs Filename:="C:\Documents and Settings\orcamento\Desktop\Contatos\EMails Grupo Arwek.xls", _
    FileFormat:=xlNormal, Password:="", WriteResPassword:="", _
    ReadOnlyRecommended:=False, CreateBackup:=False
    '---------------------------------------------------------'
    Sheets("Plan2").Select
    Sheets("Plan2").Name = "Industrial"
    ActiveCell.FormulaR1C1 = "Nome"
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "Empresa"
    Range("C1").Select
    ActiveCell.FormulaR1C1 = "E-Mail 01"
    Range("D1").Select
    ActiveCell.FormulaR1C1 = "E-Mail 02"
    Range("E1").Select
    ActiveCell.FormulaR1C1 = "E-Mail 03"
    Range("F1").Select
    ActiveCell.FormulaR1C1 = "Observação"
    Range("G1").Select
    ActiveCell.FormulaR1C1 = "Colaborador"
    Columns("A:G").Select
    Selection.ColumnWidth = 40.43
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Range("A1:G1").Select
    Selection.Font.Bold = True
    Range("A2").Select
    ActiveWorkbook.Save
    '---------------------------------------------------------'
    Sheets("Plan1").Select
    Sheets("Plan1").Name = "Residencial"
    ActiveCell.FormulaR1C1 = "Nome"
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "Empresa"
    Range("C1").Select
    ActiveCell.FormulaR1C1 = "E-Mail 01"
    Range("D1").Select
    ActiveCell.FormulaR1C1 = "E-Mail 02"
    Range("E1").Select
    ActiveCell.FormulaR1C1 = "E-Mail 03"
    Range("F1").Select
    ActiveCell.FormulaR1C1 = "Observação"
    Range("G1").Select
    ActiveCell.FormulaR1C1 = "Colaborador"
    Columns("A:G").Select
    Selection.ColumnWidth = 40.43
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Range("A1:G1").Select
    Selection.Font.Bold = True
    Range("A2").Select
    ActiveWorkbook.Save
    Set WB_2 = Workbooks("EMails Grupo Arwek.xls")
    '---------------------------------------------------------'
    intCont2 = 0: intCont3 = 2: intCont4 = 2
    '---------------------------------------------------------'
    Do While MyName <> ""
        Workbooks.Open Filename:=MyName
        Range("A2").Select
        Set WB_1 = Workbooks(Mid(ActiveWorkbook.Name, 1, InStr(1, ActiveWorkbook.Name, ".") - 1) & ".xls")
        Set WS_1 = WB_1.Worksheets("Contatos")
        '---------------------------------------------------------'
        If (LCase(Mid(WB_1.Name, 1, InStr(1, WB_1.Name, ".") - 1)) = "andre") Or (LCase(Mid(WB_1.Name, 1, InStr(1, WB_1.Name, ".") - 1)) = "omar") Or (LCase(Mid(WB_1.Name, 1, InStr(1, WB_1.Name, ".") - 1)) = "pricila") Or (LCase(Mid(WB_1.Name, 1, InStr(1, WB_1.Name, ".") - 1)) = "simone") Then
            Set WS_2 = WB_2.Worksheets("Industrial")
            intCont2 = intCont3
        Else
            Set WS_2 = WB_2.Worksheets("Residencial")
            intCont2 = intCont4
        End If
        intCont1 = 2
        Do While (Trim(WS_1.Range("A" & CStr(intCont1)).Value) <> "") Or (Trim(WS_1.Range("B" & CStr(intCont1)).Value) <> "") Or (Trim(WS_1.Range("C" & CStr(intCont1)).Value) <> "") Or (Trim(WS_1.Range("D" & CStr(intCont1)).Value) <> "") Or (Trim(WS_1.Range("E" & CStr(intCont1)).Value) <> "") Or (Trim(WS_1.Range("F" & CStr(intCont1)).Value) <> "")
            strNome = "": strEMail = "": bolAchou = False
            If (Filtro_01(Trim(WS_1.Range("A" & CStr(intCont1)).Value)) = False) And _
                (Filtro_01(Trim(WS_1.Range("B" & CStr(intCont1)).Value)) = False) And _
                (Filtro_01(Trim(WS_1.Range("C" & CStr(intCont1)).Value)) = False) And _
                (Filtro_01(Trim(WS_1.Range("D" & CStr(intCont1)).Value)) = False) And _
                (Filtro_01(Trim(WS_1.Range("E" & CStr(intCont1)).Value)) = False) And _
                (Filtro_01(Trim(WS_1.Range("F" & CStr(intCont1)).Value)) = False) And _
                (Filtro_01(Trim(WS_1.Range("G" & CStr(intCont1)).Value)) = False) Then
                If (WS_2.Range("A" & CStr(intCont2 - 1)).Value <> Trim(WS_1.Range("A" & CStr(intCont1)).Value)) Or _
                    (WS_2.Range("B" & CStr(intCont2 - 1)).Value <> Trim(WS_1.Range("B" & CStr(intCont1)).Value)) Or _
                    (WS_2.Range("C" & CStr(intCont2 - 1)).Value <> Trim(WS_1.Range("C" & CStr(intCont1)).Value)) Or _
                    (WS_2.Range("D" & CStr(intCont2 - 1)).Value <> Trim(WS_1.Range("D" & CStr(intCont1)).Value)) Or _
                    (WS_2.Range("E" & CStr(intCont2 - 1)).Value <> Trim(WS_1.Range("E" & CStr(intCont1)).Value)) Or _
                    (WS_2.Range("F" & CStr(intCont2 - 1)).Value <> Trim(WS_1.Range("F" & CStr(intCont1)).Value)) Then
                    If (Trim(WS_1.Range("A" & CStr(intCont1)).Value) <> "") Then
                        Set c = WS_2.Cells.Find(Trim(WS_1.Range("A" & CStr(intCont1)).Value), LookIn:=xlValues)
                        If Not c Is Nothing Then bolAchou = True
                    End If
                    If (Trim(WS_1.Range("B" & CStr(intCont1)).Value) <> "") Then
                        Set c = WS_2.Cells.Find(Trim(WS_1.Range("B" & CStr(intCont1)).Value), LookIn:=xlValues)
                        If Not c Is Nothing Then bolAchou = True
                    End If
                    If (Trim(WS_1.Range("C" & CStr(intCont1)).Value) <> "") Then
                        Set c = WS_2.Cells.Find(Trim(WS_1.Range("C" & CStr(intCont1)).Value), LookIn:=xlValues)
                        If Not c Is Nothing Then bolAchou = True
                    End If
                    If (Trim(WS_1.Range("D" & CStr(intCont1)).Value) <> "") Then
                        Set c = WS_2.Cells.Find(Trim(WS_1.Range("D" & CStr(intCont1)).Value), LookIn:=xlValues)
                        If Not c Is Nothing Then bolAchou = True
                    End If
                    If (Trim(WS_1.Range("E" & CStr(intCont1)).Value) <> "") Then
                        Set c = WS_2.Cells.Find(Trim(WS_1.Range("E" & CStr(intCont1)).Value), LookIn:=xlValues)
                        If Not c Is Nothing Then bolAchou = True
                    End If
                    If (Trim(WS_1.Range("F" & CStr(intCont1)).Value) <> "") Then
                        Set c = WS_2.Cells.Find(Trim(WS_1.Range("F" & CStr(intCont1)).Value), LookIn:=xlValues)
                        If Not c Is Nothing Then bolAchou = True
                    End If
                    Set c = Nothing
                    If (bolAchou = False) Then
                        WS_2.Range("A" & CStr(intCont2)).Value = Trim(WS_1.Range("A" & CStr(intCont1)).Value)
                        WS_2.Range("B" & CStr(intCont2)).Value = Trim(WS_1.Range("B" & CStr(intCont1)).Value)
                        WS_2.Range("C" & CStr(intCont2)).Value = Trim(WS_1.Range("C" & CStr(intCont1)).Value)
                        WS_2.Range("D" & CStr(intCont2)).Value = Trim(WS_1.Range("D" & CStr(intCont1)).Value)
                        WS_2.Range("E" & CStr(intCont2)).Value = Trim(WS_1.Range("E" & CStr(intCont1)).Value)
                        WS_2.Range("F" & CStr(intCont2)).Value = Trim(WS_1.Range("F" & CStr(intCont1)).Value)
                        WS_2.Range("G" & CStr(intCont2)).Value = Mid(WB_1.Name, 1, InStr(1, WB_1.Name, ".") - 1)
                        intCont2 = intCont2 + 1
                        If (LCase(Mid(WB_1.Name, 1, InStr(1, WB_1.Name, ".") - 1)) = "andre") Or (LCase(Mid(WB_1.Name, 1, InStr(1, WB_1.Name, ".") - 1)) = "omar") Or (LCase(Mid(WB_1.Name, 1, InStr(1, WB_1.Name, ".") - 1)) = "pricila") Or (LCase(Mid(WB_1.Name, 1, InStr(1, WB_1.Name, ".") - 1)) = "simone") Then
                            intCont3 = intCont3 + 1
                        Else
                            intCont4 = intCont4 + 1
                        End If
                        'If (LCase(Mid(WB_1.Name, 1, InStr(1, WB_1.Name, ".") - 1)) = "marcio") Then
                        '    intCont4 = intCont4 + 1
                        'End If
                    End If
                End If
            End If
            WS_2.Columns("A:A").EntireColumn.AutoFit
            WS_2.Columns("B:B").EntireColumn.AutoFit
            WS_2.Columns("C:C").EntireColumn.AutoFit
            WS_2.Columns("D:D").EntireColumn.AutoFit
            WS_2.Columns("E:E").EntireColumn.AutoFit
            WS_2.Columns("F:F").EntireColumn.AutoFit
            WS_2.Columns("G:G").EntireColumn.AutoFit
            intCont1 = intCont1 + 1
        Loop
        WB_1.Close: Set WS_1 = Nothing: Set WB_1 = Nothing
        WB_2.Save: Set WS_2 = Nothing
        '---------------------------------------------------------'
        MyName = Dir
        If (MyName = "EMails Grupo Arwek.xls") Then MyName = Dir
    Loop
    WB_2.Save: WB_2.Close: Set WS_2 = Nothing
End Sub

Sub Macro2()
Attribute Macro2.VB_Description = "Macro gravada em  por orcamento"
Attribute Macro2.VB_ProcData.VB_Invoke_Func = " \n14"
    Dim MyFile, MyPath, MyName
    MyPath = "C:\Documents and Settings\orcamento\Desktop\Contatos\*.csv"
    MyName = Dir(MyPath, vbNormal)
    Do While MyName <> ""
        Workbooks.OpenText MyName, xlWindows, 1, xlDelimited, xlTextQualifierSingleQuote, False, False, False, True, False, False, False
        'Workbooks.Open Filename:=MyName
        '---------------------------------------------------------'
        Columns("A:A").EntireColumn.AutoFit
        Columns("B:B").EntireColumn.AutoFit
        Columns("C:C").EntireColumn.AutoFit
        ActiveWindow.ScrollColumn = 2
        Columns("D:M").Select
        Selection.Delete Shift:=xlToLeft
        Range("D1").Select
        Columns("D:D").EntireColumn.AutoFit
        Range("E1").Select
        Columns("E:E").EntireColumn.AutoFit
        Range("F1").Select
        Columns("F:F").EntireColumn.AutoFit
        Range("G1").Select
        Columns("G:G").EntireColumn.AutoFit
        ActiveWindow.ScrollColumn = 3
        ActiveWindow.ScrollColumn = 4
        ActiveWindow.ScrollColumn = 5
        ActiveWindow.ScrollColumn = 6
        ActiveWindow.ScrollColumn = 7
        Columns("H:AF").Select
        Selection.Delete Shift:=xlToLeft
        Range("H1").Select
        Columns("H:H").EntireColumn.AutoFit
        Columns("I:BA").Select
        Selection.Delete Shift:=xlToLeft
        Range("A1:H1").Select
        Range("H1").Activate
        Selection.Font.Bold = True
        Range("A2").Select
        '---------------------------------------------------------'
        Sheets.Add
        Sheets("Plan1").Select
        Sheets("Plan1").Name = "Contatos"
        ActiveCell.FormulaR1C1 = "Nome"
        Range("B1").Select
        ActiveCell.FormulaR1C1 = "Empresa"
        Range("C1").Select
        ActiveCell.FormulaR1C1 = "E-Mail 01"
        Range("D1").Select
        ActiveCell.FormulaR1C1 = "E-Mail 02"
        Range("E1").Select
        ActiveCell.FormulaR1C1 = "E-Mail 03"
        Range("F1").Select
        ActiveCell.FormulaR1C1 = "Observação"
        Columns("A:F").Select
        Selection.ColumnWidth = 40.43
        With Selection
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlBottom
            .WrapText = False
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = xlContext
            .MergeCells = False
        End With
        Range("A1:F1").Select
        Selection.Font.Bold = True
        Range("A2").Select
        Sheets(Mid(ActiveWorkbook.Name, 1, InStr(1, ActiveWorkbook.Name, ".") - 1)).Select
        Range("A2").Select
        ChDir "C:\Documents and Settings\orcamento\Desktop\Contatos"
        ActiveWorkbook.SaveAs Filename:= _
            "C:\Documents and Settings\orcamento\Desktop\Contatos\" & Mid(ActiveWorkbook.Name, 1, InStr(1, ActiveWorkbook.Name, ".") - 1) & ".xls", _
            FileFormat:=xlNormal, Password:="", WriteResPassword:="", _
            ReadOnlyRecommended:=False, CreateBackup:=False
        '---------------------------------------------------------'
        Dim WB_1 As Workbook
        Dim WS_1 As Worksheet
        Dim WS_2 As Worksheet
        Dim intCont1, intCont2 As Integer
        Dim strNome, strEMail As String
        Dim bolAchou As Boolean
        Set WB_1 = Workbooks(Mid(ActiveWorkbook.Name, 1, InStr(1, ActiveWorkbook.Name, ".") - 1) & ".xls")
        WB_1.Sheets("Contatos").Select
        Set WS_1 = Worksheets(Mid(ActiveWorkbook.Name, 1, InStr(1, ActiveWorkbook.Name, ".") - 1))
        Set WS_2 = Worksheets("Contatos")
        intCont1 = 2: intCont2 = 2
        Do While (Trim(WS_1.Range("A" & CStr(intCont1)).Value) <> "") Or (Trim(WS_1.Range("B" & CStr(intCont1)).Value) <> "") Or (Trim(WS_1.Range("C" & CStr(intCont1)).Value) <> "") Or (Trim(WS_1.Range("D" & CStr(intCont1)).Value) <> "") Or (Trim(WS_1.Range("E" & CStr(intCont1)).Value) <> "") Or (Trim(WS_1.Range("F" & CStr(intCont1)).Value) <> "")
            strNome = "": strEMail = "": bolAchou = False
            If (Filtro_01(Trim(WS_1.Range("A" & CStr(intCont1)).Value)) = False) And _
                (Filtro_01(Trim(WS_1.Range("B" & CStr(intCont1)).Value)) = False) And _
                (Filtro_01(Trim(WS_1.Range("C" & CStr(intCont1)).Value)) = False) And _
                (Filtro_01(Trim(WS_1.Range("D" & CStr(intCont1)).Value)) = False) And _
                (Filtro_01(Trim(WS_1.Range("E" & CStr(intCont1)).Value)) = False) And _
                (Filtro_01(Trim(WS_1.Range("F" & CStr(intCont1)).Value)) = False) And _
                (Filtro_01(Trim(WS_1.Range("G" & CStr(intCont1)).Value)) = False) Then
                If (Trim(WS_1.Range("A" & CStr(intCont1)).Value) <> "") Then strNome = Trim(WS_1.Range("A" & CStr(intCont1)).Value)
                If (Trim(WS_1.Range("B" & CStr(intCont1)).Value) <> "") Then strNome = strNome & " " & Trim(WS_1.Range("B" & CStr(intCont1)).Value)
                If (Trim(WS_1.Range("C" & CStr(intCont1)).Value) <> "") Then strNome = strNome & " " & Trim(WS_1.Range("C" & CStr(intCont1)).Value)
                If (Trim(WS_1.Range("D" & CStr(intCont1)).Value) <> "") Then
                    If (InStr(1, Trim(WS_1.Range("D" & CStr(intCont1)).Value), "@", vbTextCompare) > 0) Then
                        strEMail = Mid(Trim(WS_1.Range("D" & CStr(intCont1)).Value), InStr(1, Trim(WS_1.Range("D" & CStr(intCont1)).Value), ":", vbTextCompare) + 2, (Len(Trim(WS_1.Range("D" & CStr(intCont1)).Value)) - 1) - (InStr(1, Trim(WS_1.Range("D" & CStr(intCont1)).Value), ":", vbTextCompare) + 2))
                    Else
                        strEMail = Trim(WS_1.Range("D" & CStr(intCont1)).Value)
                    End If
                End If
                If (WS_2.Range("A" & CStr(intCont2 - 1)).Value <> strNome) Or _
                    (WS_2.Range("B" & CStr(intCont2 - 1)).Value <> Trim(WS_1.Range("H" & CStr(intCont1)).Value)) Or _
                    (WS_2.Range("C" & CStr(intCont2 - 1)).Value <> Trim(WS_1.Range("E" & CStr(intCont1)).Value)) Or _
                    (WS_2.Range("D" & CStr(intCont2 - 1)).Value <> Trim(WS_1.Range("F" & CStr(intCont1)).Value)) Or _
                    (WS_2.Range("E" & CStr(intCont2 - 1)).Value <> Trim(WS_1.Range("G" & CStr(intCont1)).Value)) Or _
                    (WS_2.Range("F" & CStr(intCont2 - 1)).Value <> strEMail) Then
                    If (strNome <> "") Then
                        Set c = WS_2.Cells.Find(strNome, LookIn:=xlValues)
                        If Not c Is Nothing Then bolAchou = True
                    End If
                    If (Trim(WS_1.Range("H" & CStr(intCont1)).Value) <> "") Then
                        Set c = WS_2.Cells.Find(Trim(WS_1.Range("H" & CStr(intCont1)).Value), LookIn:=xlValues)
                        If Not c Is Nothing Then bolAchou = True
                    End If
                    If (Trim(WS_1.Range("E" & CStr(intCont1)).Value) <> "") Then
                        Set c = WS_2.Cells.Find(Trim(WS_1.Range("E" & CStr(intCont1)).Value), LookIn:=xlValues)
                        If Not c Is Nothing Then bolAchou = True
                    End If
                    If (Trim(WS_1.Range("F" & CStr(intCont1)).Value) <> "") Then
                        Set c = WS_2.Cells.Find(Trim(WS_1.Range("F" & CStr(intCont1)).Value), LookIn:=xlValues)
                        If Not c Is Nothing Then bolAchou = True
                    End If
                    If (Trim(WS_1.Range("G" & CStr(intCont1)).Value) <> "") Then
                        Set c = WS_2.Cells.Find(Trim(WS_1.Range("G" & CStr(intCont1)).Value), LookIn:=xlValues)
                        If Not c Is Nothing Then bolAchou = True
                    End If
                    If (strEMail <> "") Then
                        Set c = WS_2.Cells.Find(strEMail, LookIn:=xlValues)
                        If Not c Is Nothing Then bolAchou = True
                    End If
                    Set c = Nothing
                    If (InStr(1, strNome, "@", vbTextCompare) = 0) And _
                    (InStr(1, Trim(WS_1.Range("H" & CStr(intCont1)).Value), "@", vbTextCompare) = 0) And _
                    (InStr(1, Trim(WS_1.Range("E" & CStr(intCont1)).Value), "@", vbTextCompare) = 0) And _
                    (InStr(1, Trim(WS_1.Range("F" & CStr(intCont1)).Value), "@", vbTextCompare) = 0) And _
                    (InStr(1, Trim(WS_1.Range("G" & CStr(intCont1)).Value), "@", vbTextCompare) = 0) And _
                    (InStr(1, strEMail, "@", vbTextCompare) = 0) Then bolAchou = True
                    If (bolAchou = False) Then
                            WS_2.Range("A" & CStr(intCont2)).Value = strNome
                            WS_2.Range("B" & CStr(intCont2)).Value = Trim(WS_1.Range("H" & CStr(intCont1)).Value)
                            WS_2.Range("C" & CStr(intCont2)).Value = Trim(WS_1.Range("E" & CStr(intCont1)).Value)
                            WS_2.Range("D" & CStr(intCont2)).Value = Trim(WS_1.Range("F" & CStr(intCont1)).Value)
                            WS_2.Range("E" & CStr(intCont2)).Value = Trim(WS_1.Range("G" & CStr(intCont1)).Value)
                            WS_2.Range("F" & CStr(intCont2)).Value = strEMail
                            intCont2 = intCont2 + 1
                    End If
                Else
                    'strNome = ""
                End If
            End If
            intCont1 = intCont1 + 1
        Loop
        Set WS_1 = Nothing
        Set WS_2 = Nothing
        Set WB_1 = Nothing
        Columns("A:A").EntireColumn.AutoFit
        Columns("B:B").EntireColumn.AutoFit
        Columns("C:C").EntireColumn.AutoFit
        Columns("D:D").EntireColumn.AutoFit
        Columns("E:E").EntireColumn.AutoFit
        Columns("F:F").EntireColumn.AutoFit
        Columns("G:G").EntireColumn.AutoFit
        Sheets(Mid(ActiveWorkbook.Name, 1, InStr(1, ActiveWorkbook.Name, ".") - 1)).Select
        ActiveWindow.SelectedSheets.Delete
        ActiveWorkbook.Save
        ActiveWindow.Close
        '---------------------------------------------------------'
    MyName = Dir
    Loop
End Sub

Private Function Filtro_01(strEMail As String) As Boolean
    If (InStr(1, strEMail, "grupoarwek", vbTextCompare) <> 0) _
    Or (InStr(1, strEMail, "arwek", vbTextCompare) <> 0) _
    Or (InStr(1, strEMail, "atmosambiental", vbTextCompare) <> 0) _
    Or (InStr(1, strEMail, "atmos", vbTextCompare) <> 0) _
    Or (InStr(1, strEMail, "aletron", vbTextCompare) <> 0) _
    Or (InStr(1, strEMail, "wgui", vbTextCompare) <> 0) _
    Or (InStr(1, strEMail, "herosfiltros", vbTextCompare) <> 0) _
    Or (InStr(1, strEMail, "comparco", vbTextCompare) <> 0) _
    Or (InStr(1, strEMail, "gabler", vbTextCompare) <> 0) _
    Or (InStr(1, strEMail, "sepher", vbTextCompare) <> 0) _
    Or (InStr(1, strEMail, "termodin", vbTextCompare) <> 0) Then
        Filtro_01 = True
    Else
        Filtro_01 = False
    End If
End Function
