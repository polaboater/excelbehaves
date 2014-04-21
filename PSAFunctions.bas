Attribute VB_Name = "PSAFunctions"
Function interpolate(y_final, y_initial, x_final, x_initial, x_new)
'//Fn written by Andi Barendt
    'Basic interpolation formula.
    
    interpolate = ((y_final - y_initial) / (x_final - x_initial)) * (x_new - x_initial) + y_initial
End Function

Function CountColor(Rng As Range, RngColor As Range) As Integer
'//Fn found at http://www.mrexcel.com/td0106.html
    'The below function counts the number of cells of a certain color.
    'In the first term, select the range of cells to search.
    'In the second term, select a cell containing the color to search for.
    
    Dim cll As Range
    Dim Clr As Long
    Clr = RngColor.Range("A1").Interior.Color
    For Each cll In Rng
        If cll.Interior.Color = Clr Then
            CountColor = CountColor + 1
        End If
    Next cll
End Function

Function ColorPercent(Rng As Range, RngColor As Range) As Variant
'//Fn inspired by code found at http://www.mrexcel.com/td0106.html, adapted by Andi Barendt
    'The below function counts the total number of cells and cells of a given color
    'The funtion returns a decimal value percentage of cells of the given color
    'In the first term, select the range of cells to search.
    'In the second term, select a cell containing the color to search for.
    
    Dim cll As Range
    Dim Clr As Variant
    Dim count As Variant
    Dim ColorCount As Variant
           
    Clr = RngColor.Range("A1").Interior.Color
    count = WorksheetFunction.count(Rng)
    ColorCount = 0
    For Each cll In Rng
        If cll.Interior.Color = Clr Then
            ColorCount = ColorCount + 1
        End If
    Next cll
    
    ColorPercent = ColorCount / count
    
End Function

Function Postround(cll As Variant, dec As Integer) As String
'Function developed and written by Andi Barendt
'Lets call this v.2.2
'What it does:
    'Captures and rounds values from a cell to a number of significant DECIMALS (note: not sig figs)
    'Leverages the function Stripped() to do the heavy lifting

'What it does not do:
    'calculate sigfigs for values greater than zero
    
'Variable declaration
    Dim value As Variant        'the value of the cell we are rounding
   
        
'This is where the magic happens in the code
'First, the color of the cell is checked by the code and then shunted into the proper if statement.
'Then, the value is adjusted by the proper factor and shunted into stripped for math.
    
    'Light blue -- standard NS
    'Checks the color in the first if
    If cll.Interior.Color = 14857357 Then
        value = cll
        Postround = "NS (" & Stripped(value, dec) & ")"
    
    'Yellow -- ND, 1/2 detection limit
    ElseIf cll.Interior.Color = 65535 Then
        value = cll * 2
        Postround = "<" & Stripped(value, dec)
    
    'Pink -- ND, detection limit
    ElseIf cll.Interior.Color = 12040422 Then
        value = cll
        Postround = "<" & Stripped(value, dec)
    
    'Orange -- ND, 1/10 detection limit
    ElseIf cll.Interior.Color = 4626167 Then
        value = cll * 10
        Postround = "<" & Stripped(value, dec)

    'Purple -- NS, ND, detection limit
    ElseIf cll.Interior.Color = 13082801 Then
        value = cll
        Postround = "NS (<" & Stripped(value, dec) & ")"
    
    'Light green -- NS, ND, 1/2 detection limit
    ElseIf cll.Interior.Color = 5296274 Then
        value = cll * 2
        Postround = "NS (<" & Stripped(value, dec) & ")"
    
    'Dark Green - NS, ND, 1/10 DL
    ElseIf cll.Interior.Color = 5287936 Then
        value = cll * 10
        Postround = "NS (<" & Stripped(value, dec) & ")"
    
    'Middle grey in selection bar -- Use for first round stripping off
    ElseIf cll.Interior.Color = 12566463 Then
        value = cll
        Postround = "<" & Stripped(value, dec)
    
    'Everything else
    Else
        value = cll
        Postround = Stripped(value, dec)
    End If

    
End Function

Function Stripped(cll As Variant, dec As Integer) As String
    'Variable declaration
    Dim log As Variant          'becomes the log of the value being affected
    Dim exit_loop As Integer    'tells the FOR loops if they need to break or continue
    Dim n As Integer            'eventually is the number of decimals needed to display at least 1 digit
    Dim count As Integer        'general counting variable, counts up to decimals when needed
    Dim value As Variant        'the value of the cell we are rounding
    Dim n_start As Integer      'an adjustment factor for n when the zero condition is met
    Dim n_dec As Integer        'an adjustment factor for n when the zero condition is met
    Dim n_2 As Integer          'an adjustment factor for n when the zero condition is met
    Dim value_n As Variant      'an adjustment of value for determining n_2
    Dim inttest As Integer      'used to test if a value for sigfigs
    Dim int_log As Integer      'used for a logic test
    Dim decimals As Integer
    Dim negdec As Integer       'the negative of the number of decimals, used for determining if an adjusted n is needed
    
'Assigns initial values to non-changing variables
    If dec = 0 Then
        decimals = 1
    Else
        decimals = dec
    End If
    negdec = decimals * -1
    exit_loop = 0
    value = cll
    log = WorksheetFunction.log(value)
    
'This section will capture one sigfig of any value with log < -n
    'The two if statements determine if an adjustment to n is needed. If no, then n = decimals
    If cll < 1 And cll > 0 Then
        If (log - negdec) > 0 Then
            For inttest = 0 To 324
                If (log - negdec) = inttest Then
                    int_log = 1
                    exit_loop = 1
                        If exit_loop <> 0 Then
                            Exit For
                        End If
                Else
                    exit_loop = 0
                End If
            Next
            
            If int_log = 1 Then
                value_n = value * 10 ^ decimals
                n_start = WorksheetFunction.RoundUp(WorksheetFunction.log(value_n), 0)
                n_dec = n_start + negdec
                n_2 = n_dec + 1
                n = decimals - n_2
            Else
                value_n = value * 10 ^ decimals
                n_start = WorksheetFunction.RoundUp(WorksheetFunction.log(value_n), 0)
                n_dec = n_start + negdec
                'n_2 = n_dec + 1
                n = decimals - n_dec
            End If
            
        Else
            value_n = value * 10 ^ decimals
            n_start = WorksheetFunction.RoundUp(WorksheetFunction.log(value_n), 0)
            n_dec = n_start + negdec
            n_2 = n_dec + 1
            n = decimals - n_2
        End If
    Else
        n = decimals
    End If
    
    value = cll
    log = WorksheetFunction.log(value)
        For count = 0 To decimals
            If log < count Then
                Stripped = WorksheetFunction.Fixed(value, n, False)
                exit_loop = 1
            If exit_loop <> 0 Then
                Exit For
            End If
            ElseIf log >= count And n = 0 Then
                Stripped = WorksheetFunction.Fixed(value, 0, False)
                exit_loop = 1
                If exit_loop <> 0 Then
                    Exit For
                End If
            Else
                n = n - 1
                exit_loop = 0
            End If
        Next

End Function

Function SigFigs(cll As Variant, decimals As Integer) As Double
'Yes, it finally exists, every engineer's favorite thing: sig figs. _
Note: this does not returen a string of your number at x number of sig figs (think postround), _
instead it rounds the number to the number of sig figs as a double so that you can use it in _
calculations. Any making things pretty has to be done by you, the engineer. _
    Remember, the computer is stupid. _
    Example output: _
    sigfigs(3.389,3) = 3.39 _
    sigfigs(3.389,2) = 3.4
    
'Variable declaration
    Dim log As Variant          'becomes the log of the value being affected
    Dim exit_loop As Integer    'tells the FOR loops if they need to break or continue
    Dim n As Integer            'eventually is the number of decimals needed to display at least 1 digit
    Dim count As Integer        'general counting variable, counts up to decimals when needed
    Dim value As Variant        'the value of the cell we are rounding
    Dim n_start As Integer      'an adjustment factor for n when the zero condition is met
    Dim n_dec As Integer        'an adjustment factor for n when the zero condition is met
    Dim n_2 As Integer          'an adjustment factor for n when the zero condition is met
    Dim value_n As Variant      'an adjustment of value for determining n_2
    Dim inttest As Integer      'used to test if a value for sigfigs
    Dim int_log As Integer      'used for a logic test
    Dim negdec As Integer       'the negative of the number of decimals, used for determining if an adjusted n is needed
    Dim dec As Integer          'an always non-zero number of significant digits (minimum is one)
    Dim neg_check As Integer    'checks and adapts for negatives
    
'Assigns initial values to non-changing variables
    'ensures that there is always at least one digit
    If decimals = 0 Then
        dec = 1
    Else
        dec = decimals
    End If
    negdec = dec * -1
    exit_loop = 0
    
    'checks for negatives
    If cll < 0 Then
        neg_check = -1
        value = cll * neg_check
    Else
        neg_check = 1
        value = cll
    End If
    
    log = WorksheetFunction.log(value)
    
'This section will capture one sigfig of any value with log <= 0
    'The two if statements determine if an adjustment to n is needed. If no, then n = decimals
    If log <= 0 Then
        If (log - negdec) > 0 Then
        ' this statement handles the finicky zone between log = -n and log = 0, its a pain.
            'checks to see if there is an integer or zero log
            For inttest = 0 To 324
                If (log - negdec) = inttest Then
                    int_log = 1
                    exit_loop = 1
                        If exit_loop <> 0 Then
                            Exit For
                        End If
                Else
                    exit_loop = 0
                End If
            Next
            
            'this affects integers only
            If int_log = 1 Then
                value_n = value * 10 ^ dec
                n_start = WorksheetFunction.RoundUp(WorksheetFunction.log(value_n), 0)
                n_dec = n_start + negdec
                n_2 = n_dec + 1
                n = dec - n_2
            'this handles non-integers
            Else
                value_n = value * 10 ^ dec
                n_start = WorksheetFunction.RoundUp(WorksheetFunction.log(value_n), 0)
                n_dec = n_start + negdec
                'n_2 = n_dec + 1
                n = dec - n_dec
            End If
        'this handles values with log < -n
        Else
            value_n = value * 10 ^ dec
            n_start = WorksheetFunction.RoundUp(WorksheetFunction.log(value_n), 0)
            n_dec = n_start + negdec
            n_2 = n_dec + 1
            n = dec - n_2
        End If
        
        SigFigs = WorksheetFunction.Round(neg_check * value, n)
    Else
        'values less than 10
        If log < 1 Then
            SigFigs = WorksheetFunction.Round(neg_check * value, dec - 1)
        'values greater than 10
        Else
            n = 0
            Do Until WorksheetFunction.log(value) < 1
                value = value / 10
                n = n + 1
            Loop
            SigFigs = WorksheetFunction.Round(neg_check * value, dec - 1) * 10 ^ n
        End If
    End If
    
End Function

