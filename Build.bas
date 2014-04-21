Attribute VB_Name = "Build"
Sub Converted()
    
    Sheets("Adjusted Raw").Select
    Sheets("Adjusted Raw").Copy After:=Sheets(Sheets.count)
    Sheets("Adjusted Raw (2)").Select
    Sheets("Adjusted Raw (2)").Name = "Converted"
    
End Sub

Sub Post()
    Dim columncount As Integer
    Dim rowcount As Integer
    Dim sigdec As Integer
    Dim x As Integer
    Dim y As Integer
    
    columncount = Sheets("Stats").Cells(2, 2).value + 3
    rowcount = Sheets("Stats").Cells(1, 2).value + 1
    sigdec = Sheets("Instructions!").Cells(20, 21).value
    
    'Filling in the post
    Sheets("Post").Select
    
    ActiveSheet.Cells(1, 4).Formula = "=text('Converted'!d1," & Chr(34) & "YYYY-MM" & Chr(34) & ")"
    ActiveSheet.Cells(1, 4).Copy

    For x = 4 To columncount
        Cells(1, x).Select
        ActiveSheet.Paste
    Next
    Rows("1:1").Select
    Selection.Copy
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    For y = 1 To rowcount
        For x = 1 To 3
            ActiveSheet.Cells(y, x).value = Sheets("Adjusted Raw").Cells(y, x).value
        Next
    Next
       
    ActiveSheet.Cells(2, 4).Formula = "=if('Adjusted Raw'!D2=" & Chr(34) & Chr(34) & ",if('Converted'!D2=" & Chr(34) & Chr(34) & "," & Chr(34) & Chr(34) & ",postround('Converted'!D2," & sigdec & ")),if('Adjusted Raw'!D2=" & Chr(34) & "ND" & Chr(34) & ", " & Chr(34) & "ND" & Chr(34) & ",postround('Adjusted Raw'!D2," & sigdec & ")))"
    ActiveSheet.Cells(2, 4).Copy
    ActiveSheet.Range(ActiveSheet.Cells(2, 4), ActiveSheet.Cells(rowcount, columncount)).Select
    ActiveSheet.Paste
    
End Sub

Sub Grid()
    Dim columncount As Integer
    Dim rowcount As Integer
    Dim x As Integer
    Dim y As Integer
    
    columncount = Sheets("Stats").Cells(2, 2).value + 3
    rowcount = Sheets("Stats").Cells(1, 2).value + 1
    
     Sheets("Grid").Select
    
    
    For x = 1 To columncount
        ActiveSheet.Cells(1, x).value = Sheets("Adjusted Raw").Cells(1, x).value
    Next
    
    ActiveSheet.Cells(1, 4).Formula = "=text('Converted'!d1," & Chr(34) & "YYYY-MM" & Chr(34) & ")"
    ActiveSheet.Cells(1, 4).Copy

    For x = 4 To columncount
        Cells(1, x).Select
        ActiveSheet.Paste
    Next
    Rows("1:1").Select
    Selection.Copy
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    For y = 1 To rowcount
        For x = 1 To 3
            ActiveSheet.Cells(y, x).value = Sheets("Adjusted Raw").Cells(y, x).value
        Next
    Next
   
    ActiveSheet.Cells(2, 4).Formula = "=if('Converted'!D2=" & Chr(34) & Chr(34) & "," & Chr(34) & Chr(34) & ",log('Converted'!D2))"
    ActiveSheet.Cells(2, 4).Copy
    ActiveSheet.Range(ActiveSheet.Cells(2, 4), ActiveSheet.Cells(rowcount, columncount)).Select
    ActiveSheet.Paste
    
    Sheets("Converted").Select
    ActiveSheet.Range(ActiveSheet.Cells(1, 1), ActiveSheet.Cells(rowcount, columncount)).Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("Post").Select
    ActiveSheet.Range(ActiveSheet.Cells(1, 1), ActiveSheet.Cells(rowcount, columncount)).Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False

End Sub

Sub Stats()
Attribute Stats.VB_ProcData.VB_Invoke_Func = "U\n14"
'
' Updates the stats page with any updated custom post labling, etc...
'

'
    Dim columncount As Integer
    Dim rowcount As Integer
    Dim ActualDataPoints As Integer
    Dim TotalDataPoints As Integer
    Dim Interpolated As Integer
    Dim TotalND As Integer
    Dim AssumedND As Integer
    Dim ElevatedND As Integer
   
    Sheets("Stats").Select
    ActiveSheet.Cells(1, 1).Formula = "=CountA('Adjusted Raw'!A:A)"
    rowcount = ActiveSheet.Cells(1, 1).value
    ActiveSheet.Cells(2, 1).Formula = "=CountA('Adjusted Raw'!1:1)"
    columncount = ActiveSheet.Cells(2, 1).value
    ActiveSheet.Cells(1, 1).value = "Number of Wells:"
    ActiveSheet.Cells(2, 1).value = "Number of Events:"
    ActiveSheet.Cells(2, 2).value = columncount - 3
    ActiveSheet.Cells(1, 2).value = rowcount - 1
   
       'Filling In Data on the Stats Page
    Sheets("Stats").Select
    
    'named rows and columns
    ActiveSheet.Cells(4, 2).value = "Count"
    ActiveSheet.Cells(4, 3).value = "Percent"
    ActiveSheet.Cells(5, 1).value = "Data Points from Raw Data:"
    ActiveSheet.Cells(6, 1).value = "Non-detects from Raw Data:"
    ActiveSheet.Cells(7, 1).value = "Total Number of Data Points:"
    ActiveSheet.Cells(8, 1).value = "Total Interpolated Values:"
    ActiveSheet.Cells(9, 1).value = "Total Assumed Non-detect:"
    ActiveSheet.Cells(10, 1).value = "Total Elevated Non-detect:"
    
    'Gathering Data for Stats
    Sheets("Adjusted Raw").Select
    ActualDataPoints = WorksheetFunction.count(ActiveSheet.Range(ActiveSheet.Cells(2, 4), ActiveSheet.Cells(rowcount, columncount))) + WorksheetFunction.CountIf(ActiveSheet.Range(ActiveSheet.Cells(2, 4), ActiveSheet.Cells(rowcount, columncount)), "ND")
    TotalND = PSAFunctions.CountColor(ActiveSheet.Range(ActiveSheet.Cells(2, 4), ActiveSheet.Cells(rowcount, columncount)), Sheets("Instructions!").Range("R18"))

    Sheets("Converted").Select
    TotalDataPoints = WorksheetFunction.count(ActiveSheet.Range(ActiveSheet.Cells(2, 4), ActiveSheet.Cells(rowcount, columncount))) + WorksheetFunction.CountIf(ActiveSheet.Range(ActiveSheet.Cells(2, 4), ActiveSheet.Cells(rowcount, columncount)), "ND")
    Interpolated = PSAFunctions.CountColor(ActiveSheet.Range(ActiveSheet.Cells(2, 4), ActiveSheet.Cells(rowcount, columncount)), Sheets("Instructions!").Range("R15"))
    AssumedND = PSAFunctions.CountColor(ActiveSheet.Range(ActiveSheet.Cells(2, 4), ActiveSheet.Cells(rowcount, columncount)), Sheets("Instructions!").Range("R16")) + PSAFunctions.CountColor(ActiveSheet.Range(ActiveSheet.Cells(2, 4), ActiveSheet.Cells(rowcount, columncount)), Sheets("Instructions!").Range("R14")) + PSAFunctions.CountColor(ActiveSheet.Range(ActiveSheet.Cells(2, 4), ActiveSheet.Cells(rowcount, columncount)), Sheets("Instructions!").Range("R13"))
    ElevatedND = PSAFunctions.CountColor(ActiveSheet.Range(ActiveSheet.Cells(2, 4), ActiveSheet.Cells(rowcount, columncount)), Sheets("Instructions!").Range("R17"))

    
    'Formulas for calculating stats
    Sheets("Stats").Select
    ActiveSheet.Cells(5, 2).value = ActualDataPoints
    ActiveSheet.Cells(6, 2).value = TotalND
    ActiveSheet.Cells(7, 2).value = TotalDataPoints
    ActiveSheet.Cells(8, 2).value = Interpolated
    ActiveSheet.Cells(9, 2).value = AssumedND
    ActiveSheet.Cells(10, 2).value = ElevatedND
    ActiveSheet.Cells(5, 3).Formula = "=(b5/$b$7)"
    ActiveSheet.Cells(5, 3).Copy
    ActiveSheet.Range("C5:C10").Select
    ActiveSheet.Paste
    Selection.Style = "Percent"
    ActiveSheet.Columns("A:A").EntireColumn.AutoFit
End Sub

Sub Append()
    Dim columncount As Integer
    Dim rowcount As Integer
    Dim sigdec As Integer
    Dim x As Integer
    Dim y As Integer

    columncount = Sheets("Stats").Cells(2, 2).value + 3
    rowcount = Sheets("Stats").Cells(1, 2).value + 1
    sigdec = Sheets("Instructions!").Cells(20, 21).value
    
    Sheets("Post").Select
    
    For y = 2 To rowcount
        For x = 4 To columncount
            If Cells(y, x).Interior.Color = 9592886 Then
                If Sheets("Adjusted Raw").Cells(y, x) = "ND" Then
                    Sheets("Post").Cells(y, x).Formula = "ND (" & Stripped(Sheets("converted").Cells(y, x), sigdec) & ")"
                Else
                    Sheets("Post").Cells(y, x).Formula = Postround(Sheets("Adjusted Raw").Cells(y, x), sigdec) & " (" & Stripped(Sheets("converted").Cells(y, x), sigdec) & ")"
                End If
            End If
        Next
    Next
    
End Sub
