' Resting Metabolic Rate (RMR) Calculator
' Jason Friend-Fellows
' August 29, 1974, August 29, 1974


Option Explicit On
Option Strict On

Public Class Mainfrm
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles btnCalculate.Click
        'Declare variables for the calculations
        Dim weight, weightlbs, height, heightcm, age As Decimal
        weightlbs = CDec(txtWeight.Text)
        heightcm = CDec(txtHeight.Text)
        age = CDec(txtAge.Text)
        weight = CDec(weightlbs * 0.45359237)
        height = CDec(heightcm * 2.54)
        If cboGender.Text = "Male" Then
            lblRMR.Text = CType(CType((weight * 10) + (height * 6.25) - (age * 5) + 5, Integer), String)
        Else
            lblRMR.Text = CType(CType((weight * 10) + (height * 6.25) - (age * 5) - 161, Integer), String)
        End If
    End Sub

    Private Sub Button1_Click_1(sender As Object, e As EventArgs) Handles btnClear.Click
        txtWeight.Clear()
        txtHeight.Clear()
        txtAge.Clear()
        lblRMR.Text = ""
        cboGender.Text = "Select Gender"
    End Sub


End Class
