Attribute VB_Name = "PostGridMacro"
Sub AllPost()
Attribute AllPost.VB_Description = "G"
Attribute AllPost.VB_ProcData.VB_Invoke_Func = "G\n14"
'
' Post Macro
'
' Keyboard Shortcut: Ctrl+Shift+G
'
    'counted variables
    Dim columncount As Integer
    Dim rowcount As Integer
    
    If Sheets(Sheets.count).Name <> "Stats" Then
        If Sheets(Sheets.count).Name = "Adjusted Raw" Then
            Call Build.Converted
        End If
        Sheets.Add After:=Sheets(Sheets.count)
        Sheets(Sheets.count).Name = "Post"
        Sheets.Add After:=Sheets(Sheets.count)
        Sheets(Sheets.count).Name = "Grid"
        Sheets.Add After:=Sheets(Sheets.count)
        Sheets(Sheets.count).Name = "Stats"
    End If
    
    Sheets("Stats").Select
    ActiveSheet.Cells(1, 1).Formula = "=CountA('Adjusted Raw'!A:A)"
    rowcount = ActiveSheet.Cells(1, 1).value
    ActiveSheet.Cells(2, 1).Formula = "=CountA('Adjusted Raw'!1:1)"
    columncount = ActiveSheet.Cells(2, 1).value
    ActiveSheet.Cells(1, 1).value = "Number of Wells:"
    ActiveSheet.Cells(2, 1).value = "Number of Events:"
    ActiveSheet.Cells(2, 2).value = columncount - 3
    ActiveSheet.Cells(1, 2).value = rowcount - 1
    
    Call Build.Post
    Call Build.Grid
    Call Build.Stats
    
    If Sheets("Instructions!").Cells(21, 21) = "Yes" Then
        Call Build.Append
    End If
    
End Sub
