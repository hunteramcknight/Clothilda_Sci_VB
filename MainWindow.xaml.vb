Class MainWindow
    Dim num0 As Double
    Dim num1, num2, sign, fn As String
    Dim refresher, newNum As Boolean
    Public Sub Reset_All()
        o_Display.Text = "0"
        sign = ""
        num0 = 0
        num1 = ""
        num2 = ""
        refresher = True
        newNum = False
        fn = ""

    End Sub

    Private Sub Cmd_Clear_Click() Handles cmd_ClearAll.Click

        Reset_All()

    End Sub

    Private Sub Fx_sin_Click() Handles f_sin.Click
        fn = "sin"
        MathFunction(num0, fn)
    End Sub
    Private Sub Fx_cos_Click() Handles f_cos.Click
        fn = "cos"
        MathFunction(num0, fn)
    End Sub
    Private Sub Fx_tan_Click() Handles f_tan.Click
        fn = "tan"
        MathFunction(num0, fn)
    End Sub
    Private Sub Fx_csc_Click() Handles f_csc.Click
        fn = "csc"
        MathFunction(num0, fn)
    End Sub
    Private Sub Fx_sec_Click() Handles f_sec.Click
        fn = "sec"
        MathFunction(num0, fn)
    End Sub
    Private Sub Fx_cot_Click() Handles f_cot.Click
        fn = "cot"
        MathFunction(num0, fn)
    End Sub
    Private Sub Fx_sinh_Click() Handles f_sinh.Click
        fn = "sinh"
        MathFunction(num0, fn)
    End Sub
    Private Sub Fx_cosh_Click() Handles f_cosh.Click
        fn = "cosh"
        MathFunction(num0, fn)
    End Sub
    Private Sub Fx_tanh_Click() Handles f_tanh.Click
        fn = "tanh"
        MathFunction(num0, fn)
    End Sub
    Private Sub Fx_csch_Click() Handles f_csch.Click
        fn = "csch"
        MathFunction(num0, fn)
    End Sub
    Private Sub Fx_sech_Click() Handles f_sech.Click
        fn = "sech"
        MathFunction(num0, fn)
    End Sub
    Private Sub Fx_coth_Click() Handles f_coth.Click
        fn = "coth"
        MathFunction(num0, fn)
    End Sub
    Private Sub Fx_ln_Click() Handles f_ln.Click
        fn = "ln"
        MathFunction(num0, fn)
    End Sub
    Private Sub Fx_log_Click() Handles f_log.Click
        fn = "log"
        MathFunction(num0, fn)
    End Sub
    Private Sub Fx_sqrt_Click() Handles f_sqrt.Click
        fn = "sqrt"
        MathFunction(num0, fn)
    End Sub
    Private Sub o_0_Click(sender As Object, e As RoutedEventArgs) Handles o_0.Click
        If sign <> "" And newNum = False Then
            o_Display.Text = ""
            newNum = True
        Else
        End If

        refresher = True

        If Val(o_Display.Text) <> 0 Or o_Display.Text = "" Then 'Eliminate leading zero. This is done for all 'numeric' keys except the decimal.
            o_Display.Text = o_Display.Text & 0
        Else
        End If

    End Sub

    Private Sub o_1_Click(sender As Object, e As RoutedEventArgs) Handles o_1.Click
        'Numeric key methods check to see If the string sign <> "" And boolean newNum = False, as would both be set by an operation key method.
        ' This is required in order to set the display to null in preparation for the newyly pressed numeric key (since it is the next term in the arithmetic sequence).
        'The null display is not seen by the user since the newly pressed key is instantly displayed as the first digit.

        If sign <> "" And newNum = False Then
            o_Display.Text = ""
            newNum = True
        Else
        End If

        refresher = True

        If Val(o_Display.Text) = 0 Then 'Eliminate leading zero. This is done for all 'numeric' keys except the decimal.
            o_Display.Text = 1
        Else 'When a numeric key is not the first digit in its arithmetic term, clicking it appends it to the far right ("ones place") of the current output.
            o_Display.Text = o_Display.Text & 1
        End If

    End Sub

    Private Sub o_2_Click(sender As Object, e As RoutedEventArgs) Handles o_2.Click
        If sign <> "" And newNum = False Then
            o_Display.Text = ""
            newNum = True
        Else
        End If

        refresher = True

        If Val(o_Display.Text) = 0 Then
            o_Display.Text = 2
        Else
            o_Display.Text = o_Display.Text & 2
        End If

    End Sub
    Private Sub o_3_Click(sender As Object, e As RoutedEventArgs) Handles o_3.Click
        If sign <> "" And newNum = False Then
            o_Display.Text = ""
            newNum = True
        Else
        End If

        refresher = True

        If Val(o_Display.Text) = 0 Then
            o_Display.Text = 3
        Else
            o_Display.Text = o_Display.Text & 3
        End If

    End Sub

    Private Sub o_4_Click(sender As Object, e As RoutedEventArgs) Handles o_4.Click
        If sign <> "" And newNum = False Then
            o_Display.Text = ""
            newNum = True
        Else
        End If

        refresher = True

        If Val(o_Display.Text) = 0 Then
            o_Display.Text = 4
        Else
            o_Display.Text = o_Display.Text & 4
        End If

    End Sub

    Private Sub o_5_Click(sender As Object, e As RoutedEventArgs) Handles o_5.Click
        If sign <> "" And newNum = False Then
            o_Display.Text = ""
            newNum = True
        Else
        End If

        refresher = True

        If Val(o_Display.Text) = 0 Then
            o_Display.Text = 5
        Else
            o_Display.Text = o_Display.Text & 5
        End If

    End Sub
    Private Sub o_6_Click(sender As Object, e As RoutedEventArgs) Handles o_6.Click
        If sign <> "" And newNum = False Then
            o_Display.Text = ""
            newNum = True
        Else
        End If

        refresher = True

        If Val(o_Display.Text) = 0 Then
            o_Display.Text = 6
        Else
            o_Display.Text = o_Display.Text & 6
        End If

    End Sub

    Private Sub o_7_Click(sender As Object, e As RoutedEventArgs) Handles o_7.Click
        If sign <> "" And newNum = False Then
            o_Display.Text = ""
            newNum = True
        Else
        End If

        refresher = True

        If Val(o_Display.Text) = 0 Then
            o_Display.Text = 7
        Else
            o_Display.Text = o_Display.Text & 7
        End If

    End Sub

    Private Sub o_8_Click(sender As Object, e As RoutedEventArgs) Handles o_8.Click
        If sign <> "" And newNum = False Then
            o_Display.Text = ""
            newNum = True
        Else
        End If

        refresher = True

        If Val(o_Display.Text) = 0 Then
            o_Display.Text = 8
        Else
            o_Display.Text = o_Display.Text & 8
        End If

    End Sub
    Private Sub o_9_Click(sender As Object, e As RoutedEventArgs) Handles o_9.Click
        If sign <> "" And newNum = False Then
            o_Display.Text = ""
            newNum = True
        Else
        End If

        refresher = True

        If Val(o_Display.Text) = 0 Then
            o_Display.Text = 9
        Else
            o_Display.Text = o_Display.Text & 9
        End If

    End Sub



    Private Sub o_0_Copy_Click(sender As Object, e As RoutedEventArgs)

    End Sub

    Private Sub o_Decimal_Click() Handles o_Decimal.Click

        'This conditional prevents multiple decimals within a numeric string.
        If (InStr(o_Display.Text, ".") = 0) Then

            If sign <> "" And newNum = False Then
                newNum = True
            Else
            End If

            o_Display.Text = o_Display.Text & "."
            '    refresher = True
        Else
        End If

    End Sub

    Private Sub o_E_Click() Handles o_E.Click
        If sign <> "" And newNum = False Then
            'o_Display.Text = ""
            newNum = True
        Else
        End If

        refresher = True
        If Val(o_Display.Text) = 0 Then
            o_Display.Text = Math.Exp(1)
        Else
            o_Display.Text = o_Display.Text * Math.Exp(1)
        End If

    End Sub

    Private Sub o_Pi_Click() Handles o_Pi.Click
        If sign <> "" And newNum = False Then
            'o_Display.Text = ""
            newNum = True
        Else
        End If

        refresher = True
        If Val(o_Display.Text) = 0 Then
            o_Display.Text = Math.PI
        Else
            o_Display.Text = o_Display.Text * Math.PI
        End If

    End Sub

    Private Sub Opp_add_Click() Handles opp_add.Click
        'Operation methods look to see if the boolean variable "refresher" is true, as passed only by the numeric key methods.
        'and then set the refresher back to false so that other operation keys are immediately temporarily disabled until 
        'reenabled by a numeric key method. Numeric keys are not contingent on the value of the refresher in order to function.



        If num1 <> "" And refresher = True Then
            sign = "add"
            Evaluate()
            num1 = Val(o_Display.Text)
        Else
        End If

        If o_Display.Text <> "" Then
            num1 = Val(o_Display.Text)
            sign = "add"
            newNum = False
            refresher = False
        Else
        End If


    End Sub
    Private Sub Opp_sub_Click() Handles opp_sub.Click

        If num1 > "" And refresher = True Then
            sign = "sub"
            Evaluate()
            num1 = Val(o_Display.Text)
        Else
        End If

        num1 = Val(o_Display.Text)
            sign = "sub"
            newNum = False
        refresher = False

        refresher = False

    End Sub
    Private Sub Opp_mul_Click() Handles opp_mul.Click

        If num1 > "" And refresher = True Then
            sign = "mul"
            Evaluate()
            num1 = Val(o_Display.Text)
        Else
        End If

        num1 = Val(o_Display.Text)
            sign = "mul"
            newNum = False
        refresher = False

    End Sub
    Private Sub Opp_div_Click() Handles opp_div.Click

        If num1 > "" And refresher = True Then
            sign = "div"
            Evaluate()
        Else
        End If

        num1 = Val(o_Display.Text)
            sign = "div"
            newNum = False
            refresher = True


    End Sub

    'Private Sub Opt_Rad_click(sender As Object, e As RoutedEventArgs) Handles opt_Rad.Click
    '    o_Display.Text = (Val(o_Display.Text) / 180) * Math.PI
    'End Sub

    'Private Sub Opt_Deg_click(sender As Object, e As RoutedEventArgs) Handles opt_Deg.Click
    '    o_Display.Text = (Val(o_Display.Text) / Math.PI) * 180
    'End Sub

    Private Sub Opp_pow_Click() Handles opp_pow.Click

        If num1 > "" And refresher = True Then
            sign = "pow"
            Evaluate()
        Else
        End If

        If o_Display.Text <> "" Then
            num1 = Val(o_Display.Text)
            sign = "pow"
            newNum = False
            refresher = True
        Else
        End If

    End Sub

    Private Sub Opp_equal_Click() Handles opp_equal.Click


        Evaluate()

    End Sub

    Public Sub Evaluate()
        If num1 <> "" Then


            num2 = Val(o_Display.Text)



            Select Case sign
                Case "add"
                    o_Display.Text = Val(num1) + Val(num2)
                Case "sub"
                    o_Display.Text = Val(num1) - Val(num2)
                Case "mul"
                    o_Display.Text = Val(num1) * Val(num2)
                Case "div"
                    o_Display.Text = Val(num1) / Val(num2)
                Case "pow"
                    o_Display.Text = Val(num1) ^ Val(num2)
                Case ""
            End Select


            refresher = False
            newNum = True

        Else
        End If

    End Sub

    Public Function MathFunction(ByRef num0 As Double, ByVal fn As String)

        If o_Display.Text = "" Then num0 = 0 Else num0 = Val(o_Display.Text)
        'num0 = Val(o_Display.Text)


        Select Case fn
            Case "sqrt"
                MathFunction = Math.Sqrt(num0)
            Case "ln"
                MathFunction = Math.Log(num0)
            Case "log"
                MathFunction = Math.Log10(num0)
            Case "sin"
                MathFunction = Math.Sin(num0)
            Case "cos"
                MathFunction = Math.Cos(num0)
            Case "tan"
                MathFunction = Math.Tan(num0)
            Case "csc"
                MathFunction = 1 / Math.Sin(num0)
            Case "sec"
                MathFunction = 1 / Math.Cos(num0)
            Case "cot"
                MathFunction = 1 / Math.Tan(num0)
            Case "asin"
                MathFunction = Math.Asin(num0)
            Case "acos"
                MathFunction = Math.Acos(num0)
            Case "atan"
                MathFunction = Math.Atan(num0)
            Case "sinh"
                MathFunction = (1 - Math.Exp(-2 * num0)) / (2 * Math.Exp(-1 * num0))
            Case "cosh"
                MathFunction = (1 + Math.Exp(-2 * num0)) / (2 * Math.Exp(-1 * num0))
            Case "tanh"
                MathFunction = ((1 - Math.Exp(-2 * num0)) / (2 * Math.Exp(-1 * num0))) / ((1 + Math.Exp(-2 * num0)) / (2 * Math.Exp(-1 * num0)))
            Case "csch"
                MathFunction = 1 / ((1 - Math.Exp(-2 * num0)) / (2 * Math.Exp(-1 * num0)))
            Case "sech"
                MathFunction = 1 / ((1 + Math.Exp(-2 * num0)) / (2 * Math.Exp(-1 * num0)))
            Case "coth"
                MathFunction = 1 / (((1 - Math.Exp(-2 * num0)) / (2 * Math.Exp(-1 * num0))) / ((1 + Math.Exp(-2 * num0)) / (2 * Math.Exp(-1 * num0))))

            Case ""
                'MathFunction = 1

        End Select



        'Return MathFunction
#Disable Warning BC42104 ' Variable is used before it has been assigned a value
        o_Display.Text = Val(MathFunction)
#Enable Warning BC42104 ' Variable is used before it has been assigned a value

#Disable Warning BC42105 ' Function doesn't return a value on all code paths
    End Function
#Enable Warning BC42105 ' Function doesn't return a value on all code paths

End Class