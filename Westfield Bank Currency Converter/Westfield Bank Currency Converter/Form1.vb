Public Class Form1
    'When Update Conversion rates button clciked, displayed form is hidden and appropriate form is displayed
    Private Sub BtnUpdateConversionRates_Click(sender As Object, e As EventArgs) Handles BtnUpdateConversionRates.Click
        Me.Visible = False
        Form2.Visible = True
    End Sub
    Private Sub BtnHelp_Click(sender As Object, e As EventArgs) Handles BtnHelp.Click
        Me.Visible = False
        Form3.Visible = True
    End Sub
    'Variable decalration
    Dim ConvertFromValue As Decimal
    Dim ConvertToValue As Decimal
    Dim CommissionRate As Decimal
    Dim InitialValue As Decimal
    Dim CommissionValue As Decimal
    Dim ConvertedAmount As Decimal
    Dim Multiplier As Decimal
    Dim FinalAmount As Decimal
    Private Sub BtnConvert_Click(sender As Object, e As EventArgs) Handles BtnConvert.Click
        Call LoadRates() 'Calls the LoadRates module, in order to draw values from the PercentageRates array items

        CommissionRate = PercentageRatesArray(8) 'Assigns the CommissionRates variable the value stored in location 8 of the array

        'Validation to check that the same To and From are not selected
        If RadiobtnConvertFromGBP.Checked And RadioBtnConvertToGBP.Checked Then
            MsgBox("Please select differing currencies") 'Checks each possible case, presents error message if any of the 8 scenarios are true
        ElseIf RadiobtnConvertFromEUR.Checked And RadioBtnConvertToEUR.Checked Then
            MsgBox("Please select differing currencies", vbInformation)
        ElseIf RadiobtnConvertFromUSD.Checked And RadioBtnConvertToUSD.Checked Then
            MsgBox("Please select differing currencies", vbInformation)
        ElseIf RadiobtnConvertFromAUD.Checked And RadioBtnConvertToAUD.Checked Then
            MsgBox("Please select differing currencies", vbInformation)
        ElseIf RadiobtnConvertFromNZD.Checked And RadioBtnConvertToNZD.Checked Then
            MsgBox("Please select differing currencies", vbInformation)
        ElseIf RadiobtnConvertFromCAD.Checked And RadioBtnConvertToCAD.Checked Then
            MsgBox("Please select differing currencies", vbInformation)
        ElseIf RadiobtnConvertFromCHF.Checked And RadioBtnConvertToCHF.Checked Then
            MsgBox("Please select differing currencies", vbInformation)
        ElseIf RadiobtnConvertFromJPY.Checked And RadioBtnConvertToJPY.Checked Then
            MsgBox("Please select differing currencies", vbInformation)
        ElseIf TxtBxInitialAmount.Text = "" Then 'Also checks to make sure a value has been entered to convert
            MsgBox("Please enter a value in initial amount", vbInformation)
        Else 'Runs the rest of the program if validation is passed
            'In this if statement the program is checking which radiobutton has been checked
            If RadiobtnConvertFromGBP.Checked Then
                'When a checked button is found, it inserts a specific item from the Array into a variable.
                ConvertFromValue = PercentageRatesArray(0)
                'It also formats a label that shows a currency symbol, depending on which button is checked.
                LblCommissionAmountSymbol.Text = ("£")
            ElseIf RadiobtnConvertFromEUR.Checked Then
                ConvertFromValue = PercentageRatesArray(1)
                LblCommissionAmountSymbol.Text = ("€")
            ElseIf RadiobtnConvertFromUSD.Checked Then
                ConvertFromValue = PercentageRatesArray(2)
                LblCommissionAmountSymbol.Text = ("$")
            ElseIf RadiobtnConvertFromAUD.Checked Then
                ConvertFromValue = PercentageRatesArray(3)
                LblCommissionAmountSymbol.Text = ("AU$")
            ElseIf RadiobtnConvertFromNZD.Checked Then
                ConvertFromValue = PercentageRatesArray(4)
                LblCommissionAmountSymbol.Text = ("NZ$")
            ElseIf RadiobtnConvertFromCAD.Checked Then
                ConvertFromValue = PercentageRatesArray(5)
                LblCommissionAmountSymbol.Text = ("CA$")
            ElseIf RadiobtnConvertFromCHF.Checked Then
                ConvertFromValue = PercentageRatesArray(6)
                LblCommissionAmountSymbol.Text = ("Fr.")
            ElseIf RadiobtnConvertFromJPY.Checked Then
                ConvertFromValue = PercentageRatesArray(7)
                LblCommissionAmountSymbol.Text = ("¥")
            End If
            'The same process is repeated for the convert to currency. Via the If statement the program finds out which radiobutton has been selected.
            If RadioBtnConvertToGBP.Checked Then
                'Once the selected radiobutton has been found, the appropriate value from the array is stored in a variable.
                ConvertToValue = PercentageRatesArray(0)
                'This also updates a descriptive label to the appropriate one, depdning on the selected radiobutton.
                LblFinalAmountSymbol.Text = ("£")
            ElseIf RadioBtnConvertToEUR.Checked Then
                ConvertToValue = PercentageRatesArray(1)
                LblFinalAmountSymbol.Text = ("€")
            ElseIf RadioBtnConvertToUSD.Checked Then
                ConvertToValue = PercentageRatesArray(2)
                LblFinalAmountSymbol.Text = ("$")
            ElseIf RadioBtnConvertToAUD.Checked Then
                ConvertToValue = PercentageRatesArray(3)
                LblFinalAmountSymbol.Text = ("AU$")
            ElseIf RadioBtnConvertToNZD.Checked Then
                ConvertToValue = PercentageRatesArray(4)
                LblFinalAmountSymbol.Text = ("NZ$")
            ElseIf RadioBtnConvertToCAD.Checked Then
                ConvertToValue = PercentageRatesArray(5)
                LblFinalAmountSymbol.Text = ("CA$")
            ElseIf RadioBtnConvertToCHF.Checked Then
                ConvertToValue = PercentageRatesArray(6)
                LblFinalAmountSymbol.Text = ("Fr.")
            ElseIf RadioBtnConvertToJPY.Checked Then
                ConvertToValue = PercentageRatesArray(7)
                LblFinalAmountSymbol.Text = ("¥")
            End If
            'This line inserts the value stored in the Initial amount textbox into a variable.
            InitialValue = TxtBxInitialAmount.Text
            'The multipler is then calculated by divideing the two values assigned to variables by the previous if statements.
            Multiplier = ConvertFromValue / ConvertToValue
            'The amount taken off the inital amount by commission is then calculated by dividing the initial value by the stored commission rate.
            CommissionValue = InitialValue * CommissionRate
            'A converted amount is then calculated by taking the amount of commission from the initial value.
            ConvertedAmount = InitialValue - CommissionValue
            'The commission rate is then displayed in a textbox, formatted into a percentage.
            TxtBxCommissionRate.Text = FormatPercent(CommissionRate)
            'The amount of money taken by the commission is then displayed in another textbox, formatted to 2 decimal places
            TxtBxCommissionAmount.Text = Math.Round(CommissionValue, 2)
            'The final amount is then calulated by dividing the converted amount by the multiplier
            FinalAmount = ConvertedAmount / Multiplier
            'This is then displayed in another textbox, and formatted to have 2 decimal places.
            TxtBxFinalAmount.Text = Math.Round(FinalAmount, 2)
        End If
    End Sub
End Class
