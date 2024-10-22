Attribute VB_Name = "Main"
Sub Add_column()
    ' Inserts a new column at position O and copies content from column L to column O.
    
    Columns("O:O").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("L:L").Select
    Selection.Copy
    Columns("O:O").Select
    Selection.PasteSpecial Paste:=xlPasteAllUsingSourceTheme, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
End Sub

Sub Tabs_to_Anal()
    ' Loops through multiple sheets and updates data in "Analisi" sheets with values from "WIP".
    
    Dim r As Long
    Dim rows As Long
    Dim arr() As Variant
    Dim tabName As String
    Dim per_current As String
    per_current = InputBox("What's current period? Type in format yymm")
    
    arr = Array("GAMP", "GAPI", "DAEN", "LABA") ' Array of sheet names
    
    For i = LBound(arr) To UBound(arr)
        ' Clear contents in "Analisi" sheets for current iteration
        tabName = "ODA " & arr(i)
        Sheets("Analisi " & arr(i)).Select
        Range("A6:K6").Select
        Range(Selection, Selection.End(xlDown)).Select
        Selection.ClearContents
    
        ' Copy data from "WIP" to "Analisi" sheet
        Windows("WIP.xlsx").Activate
        Sheets(arr(i)).Select
        Range("A3:J3").Select
        Range(Selection, Selection.End(xlDown)).Select
        Selection.Copy
        
        ' Paste data into "Analisi" sheet
        Windows("UV3221_WIP_" & per_current & "_IA.xlsm").Activate
        Sheets("Analisi " & arr(i)).Select
        Range("A6").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        
        rows = Selection.Rows.Count + 5
    
        ' Applying formulas
        Range("L6").FormulaR1C1 = "=RC[-2]+RC[-1]"
        Range("M6").FormulaR1C1 = "=SUMIF('" & tabName & "'!C12,RC[-10],'" & tabName & "'!C10)"
        Range("N6").FormulaR1C1 = "=IF(RC[-1]=""0"",""0"",IF(RC[-1]<RC[-2],RC[-1],RC[-2]))"
        
        Range("L6:N6").AutoFill Destination:=Range("L6:N" & rows), Type:=xlFillDefault
    Next i
    
    ' Special handling for negative values in GAMP and DAEN sheets
    For r = 6 To 200 Step 1
        If Sheets("Analisi GAMP").Range("N" & r) < 0 Then
            Sheets("Analisi GAMP").Range("N" & r).Value = 0
            Sheets("Analisi GAMP").Range("N" & r).Interior.Color = rgbGreen
        End If
    Next r
    
    For r = 6 To 600 Step 1
        If Sheets("Analisi DAEN").Range("N" & r) < 0 Then
            Sheets("Analisi DAEN").Range("N" & r).Value = 0
            Sheets("Analisi DAEN").Range("N" & r).Interior.Color = rgbGreen
        End If
    Next r
    
    MsgBox "Done"
End Sub

Sub Files_to_Sheets()
    ' Updates the "ODA" sheets with data from the source files and performs formula-based manipulations.
    
    Dim n_lines As Long
    Dim per_current As String
    per_current = InputBox("What's current period? Type in format yymm")
    Dim arr() As Variant
    arr = Array("GAMP", "GAPI", "DAEN")
    
    For i = LBound(arr) To UBound(arr)
        ' Clear content from "ODA" sheets
        Windows("UV3221_WIP_" & per_current & "_IA.xlsm").Activate
        Sheets("ODA " & arr(i)).Select
        Range("A2:K2").Select
        Range(Selection, Selection.End(xlDown)).Select
        Selection.ClearContents
        
        ' Copy data from "ODA" source file
        Workbooks("ODA.xlsm").Worksheets(arr(i)).Activate
        Range("K3:A3").Select
        Range(Selection, Selection.End(xlDown)).Select
        Selection.Copy
        
        ' Paste data into "ODA" sheet
        Windows("UV3221_WIP_" & per_current & "_IA.xlsm").Activate
        Sheets("ODA " & arr(i)).Activate
        Range("A2").Select
        ActiveSheet.Paste
            
        n_lines = Selection.Rows.Count
        
        ' Apply formulas
        Range("L2").FormulaR1C1 = "=IF(OR(LEFT(RC[-6],1)=""A"", LEFT(RC[-6],1)=""C""), LEFT(RC[-6],6), RC[-6])"
        Range("M2").FormulaR1C1 = "=IF(ISERROR(LEFT(RC[-11],2)+1),"""",""TERZI"")"
        
        Range("L2:M2").Select
        Selection.AutoFill Destination:=Range("L2:M" & n_lines + 1), Type:=xlFillDefault
    
        ' Deleting "TERZI" records (third-party)
        ActiveSheet.Range("$A$1:$M$30009").AutoFilter Field:=13, Criteria1:="TERZI"
        Range("L2:M2").Select
        Range(Selection, Selection.End(xlDown)).Select
        Selection.ClearContents
        ActiveSheet.Range("$A$1:$M$30009").AutoFilter Field:=13
    
        ' Deleting "Italia" records
        ActiveSheet.Range("$A$1:$M$" & n_lines + 1).AutoFilter Field:=5, Criteria1:="Italia", Operator:=xlOr, Criteria2:="="
        Range("L2:M" & n_lines + 1).Select
        Selection.ClearContents
        ActiveSheet.ShowAllData
    
        ' Replace "/" with empty strings
        Columns("F:F").Select
        Selection.Replace What:="/", Replacement:="", LookAt:=xlPart, _
            SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
            ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Next i
    
    ' Special case for LABA: The selection of cells must handle empty cells
    Windows("UV3221_WIP_" & per_current & "_IA.xlsm").Activate
    Sheets("ODA LABA").Select
    ActiveCell.CurrentRegion.Select
    Selection.ClearContents
    
    ' Copy data from "LABA" sheet
    Workbooks("ODA.xlsm").Worksheets("LABA").Activate
    Dim rng As Range
    Set rng = Range("A3").CurrentRegion
    Set rng = rng.Offset(1, 0)
    Set rng = rng.Resize(rng.Rows.Count - 1)
    rng.Copy
    
    ' Paste data into "ODA LABA" sheet
    Windows("UV3221_WIP_" & per_current & "_IA.xlsm").Activate
    Sheets("ODA LABA").Select
    Range("A1").Select
    ActiveSheet.Paste
    
    n_lines = Selection.Rows.Count
    
    ' Apply formulas for LABA
    Range("L2").FormulaR1C1 = "=IF(LEFT(RC[-6],1)=""A"",LEFT(RC[-6],6),RC[-6])"
    Range("M2").FormulaR1C1 = "=IF(ISERROR(LEFT(RC[-11],2)+1),"""",""TERZI"")"
    
    Range("L2:M2").Select
    Selection.AutoFill Destination:=Range("L2:M" & n_lines + 1), Type:=xlFillDefault
    
    ' Deleting "TERZI" records in LABA
    ActiveSheet.Range("$A$1:$M$30009").AutoFilter Field:=13, Criteria1:="TERZI"
    Range("L2:M2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    ActiveSheet.Range("$A$1:$M$30009").AutoFilter Field:=13
    
    ' Deleting "Italia" records in LABA
    ActiveSheet.Range("$A$1:$M$" & n_lines + 1).AutoFilter Field:=5, Criteria1:="Italia", Operator:=xlOr, Criteria2:="="
        Range("L2:M" & n_lines + 1).Select
    Selection.ClearContents
    ActiveSheet.ShowAllData
    
    ' Replace "/" with empty strings
    Columns("F:F").Select
    Selection.Replace What:="/", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    
    MsgBox "Done"
End Sub
