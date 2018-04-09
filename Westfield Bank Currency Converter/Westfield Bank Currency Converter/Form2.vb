Public Class Form2
    Private nonNumberEntered As Boolean = False

    Private Sub TxtBxGBP_Keydown(Sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles TxtBxGBP.KeyDown 'Private sub has been made, this allows only numbers into txtSideA.'
        nonNumberEntered = False ' Initialize the flag to false
        If e.KeyCode < Keys.D0 OrElse e.KeyCode > Keys.D9 Then ' Determine whether the keystroke is a number from the top of the keyboard
            If e.KeyCode < Keys.NumPad0 OrElse e.KeyCode > Keys.NumPad9 Then 'Determine whether the keystroke is a number from the keypad
                If e.KeyCode <> Keys.Back Then ' Determine wheether the keystroke is a backspace
                    If e.KeyCode <> Keys.OemPeriod Then ' A non-numerical keystroke was pressed
                        nonNumberEntered = True ' Set the flag to true and evaluate in KeyPress event.
                    End If
                End If
            End If
        End If
    End Sub

    Private Sub TxtBxGBP_KeyPress(Sender As Object, e As System.Windows.Forms.KeyPressEventArgs) Handles TxtBxGBP.KeyPress
        If nonNumberEntered = True Then 'Check for set in the KeyDown event.
            e.Handled = True ' Stop the character from being entered into control since it is non-numerical 
        End If
    End Sub
    Private Sub TxtBxEUR_Keydown(Sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles TxtBxEUR.KeyDown 'Private sub has been made, this allows only numbers into txtSideA.'
        nonNumberEntered = False ' Initialize the flag to false
        If e.KeyCode < Keys.D0 OrElse e.KeyCode > Keys.D9 Then ' Determine whether the keystroke is a number from the top of the keyboard
            If e.KeyCode < Keys.NumPad0 OrElse e.KeyCode > Keys.NumPad9 Then 'Determine whether the keystroke is a number from the keypad
                If e.KeyCode <> Keys.Back Then ' Determine wheether the keystroke is a backspace
                    If e.KeyCode <> Keys.OemPeriod Then ' A non-numerical keystroke was pressed
                        nonNumberEntered = True ' Set the flag to true and evaluate in KeyPress event.
                    End If
                End If
            End If
        End If
    End Sub

    Private Sub TxtBxEUR_KeyPress(Sender As Object, e As System.Windows.Forms.KeyPressEventArgs) Handles TxtBxEUR.KeyPress
        If nonNumberEntered = True Then 'Check for set in the KeyDown event.
            e.Handled = True ' Stop the character from being entered into control since it is non-numerical 
        End If
    End Sub
    Private Sub TxtBxUSD_Keydown(Sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles TxtBxUSD.KeyDown 'Private sub has been made, this allows only numbers into txtSideA.'
        nonNumberEntered = False ' Initialize the flag to false
        If e.KeyCode < Keys.D0 OrElse e.KeyCode > Keys.D9 Then ' Determine whether the keystroke is a number from the top of the keyboard
            If e.KeyCode < Keys.NumPad0 OrElse e.KeyCode > Keys.NumPad9 Then 'Determine whether the keystroke is a number from the keypad
                If e.KeyCode <> Keys.Back Then ' Determine wheether the keystroke is a backspace
                    If e.KeyCode <> Keys.OemPeriod Then ' A non-numerical keystroke was pressed
                        nonNumberEntered = True ' Set the flag to true and evaluate in KeyPress event.
                    End If
                End If
            End If
        End If
    End Sub

    Private Sub TxtBxUSD_KeyPress(Sender As Object, e As System.Windows.Forms.KeyPressEventArgs) Handles TxtBxUSD.KeyPress
        If nonNumberEntered = True Then 'Check for set in the KeyDown event.
            e.Handled = True ' Stop the character from being entered into control since it is non-numerical 
        End If
    End Sub
    Private Sub TxtBxAUD_Keydown(Sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles TxtBxAUD.KeyDown 'Private sub has been made, this allows only numbers into txtSideA.'
        nonNumberEntered = False ' Initialize the flag to false
        If e.KeyCode < Keys.D0 OrElse e.KeyCode > Keys.D9 Then ' Determine whether the keystroke is a number from the top of the keyboard
            If e.KeyCode < Keys.NumPad0 OrElse e.KeyCode > Keys.NumPad9 Then 'Determine whether the keystroke is a number from the keypad
                If e.KeyCode <> Keys.Back Then ' Determine wheether the keystroke is a backspace
                    If e.KeyCode <> Keys.OemPeriod Then ' A non-numerical keystroke was pressed
                        nonNumberEntered = True ' Set the flag to true and evaluate in KeyPress event.
                    End If
                End If
            End If
        End If
    End Sub

    Private Sub TxtBxAUD_KeyPress(Sender As Object, e As System.Windows.Forms.KeyPressEventArgs) Handles TxtBxAUD.KeyPress
        If nonNumberEntered = True Then 'Check for set in the KeyDown event.
            e.Handled = True ' Stop the character from being entered into control since it is non-numerical 
        End If
    End Sub
    Private Sub TxtBxNZD_Keydown(Sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles TxtBxNZD.KeyDown 'Private sub has been made, this allows only numbers into txtSideA.'
        nonNumberEntered = False ' Initialize the flag to false
        If e.KeyCode < Keys.D0 OrElse e.KeyCode > Keys.D9 Then ' Determine whether the keystroke is a number from the top of the keyboard
            If e.KeyCode < Keys.NumPad0 OrElse e.KeyCode > Keys.NumPad9 Then 'Determine whether the keystroke is a number from the keypad
                If e.KeyCode <> Keys.Back Then ' Determine wheether the keystroke is a backspace
                    If e.KeyCode <> Keys.OemPeriod Then ' A non-numerical keystroke was pressed
                        nonNumberEntered = True ' Set the flag to true and evaluate in KeyPress event.
                    End If
                End If
            End If
        End If
    End Sub

    Private Sub TxtBxNZD_KeyPress(Sender As Object, e As System.Windows.Forms.KeyPressEventArgs) Handles TxtBxNZD.KeyPress
        If nonNumberEntered = True Then 'Check for set in the KeyDown event.
            e.Handled = True ' Stop the character from being entered into control since it is non-numerical 
        End If
    End Sub
    Private Sub TxtBxCAD_Keydown(Sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles TxtBxCAD.KeyDown 'Private sub has been made, this allows only numbers into txtSideA.'
        nonNumberEntered = False ' Initialize the flag to false
        If e.KeyCode < Keys.D0 OrElse e.KeyCode > Keys.D9 Then ' Determine whether the keystroke is a number from the top of the keyboard
            If e.KeyCode < Keys.NumPad0 OrElse e.KeyCode > Keys.NumPad9 Then 'Determine whether the keystroke is a number from the keypad
                If e.KeyCode <> Keys.Back Then ' Determine wheether the keystroke is a backspace
                    If e.KeyCode <> Keys.OemPeriod Then ' A non-numerical keystroke was pressed
                        nonNumberEntered = True ' Set the flag to true and evaluate in KeyPress event.
                    End If
                End If
            End If
        End If
    End Sub

    Private Sub TxtBxCAD_KeyPress(Sender As Object, e As System.Windows.Forms.KeyPressEventArgs) Handles TxtBxCAD.KeyPress
        If nonNumberEntered = True Then 'Check for set in the KeyDown event.
            e.Handled = True ' Stop the character from being entered into control since it is non-numerical 
        End If
    End Sub
    Private Sub TxtBxCHF_Keydown(Sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles TxtBxCHF.KeyDown 'Private sub has been made, this allows only numbers into txtSideA.'
        nonNumberEntered = False ' Initialize the flag to false
        If e.KeyCode < Keys.D0 OrElse e.KeyCode > Keys.D9 Then ' Determine whether the keystroke is a number from the top of the keyboard
            If e.KeyCode < Keys.NumPad0 OrElse e.KeyCode > Keys.NumPad9 Then 'Determine whether the keystroke is a number from the keypad
                If e.KeyCode <> Keys.Back Then ' Determine wheether the keystroke is a backspace
                    If e.KeyCode <> Keys.OemPeriod Then ' A non-numerical keystroke was pressed
                        nonNumberEntered = True ' Set the flag to true and evaluate in KeyPress event.
                    End If
                End If
            End If
        End If
    End Sub

    Private Sub TxtBxCHF_KeyPress(Sender As Object, e As System.Windows.Forms.KeyPressEventArgs) Handles TxtBxCHF.KeyPress
        If nonNumberEntered = True Then 'Check for set in the KeyDown event.
            e.Handled = True ' Stop the character from being entered into control since it is non-numerical 
        End If

    End Sub
    Private Sub TxtBxJPY_Keydown(Sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles TxtBxJPY.KeyDown 'Private sub has been made, this allows only numbers into txtSideA.'
        nonNumberEntered = False ' Initialize the flag to false
        If e.KeyCode < Keys.D0 OrElse e.KeyCode > Keys.D9 Then ' Determine whether the keystroke is a number from the top of the keyboard
            If e.KeyCode < Keys.NumPad0 OrElse e.KeyCode > Keys.NumPad9 Then 'Determine whether the keystroke is a number from the keypad
                If e.KeyCode <> Keys.Back Then ' Determine wheether the keystroke is a backspace
                    If e.KeyCode <> Keys.OemPeriod Then ' A non-numerical keystroke was pressed
                        nonNumberEntered = True ' Set the flag to true and evaluate in KeyPress event.
                    End If
                End If
            End If
        End If
    End Sub

    Private Sub TxtBxJPY_KeyPress(Sender As Object, e As System.Windows.Forms.KeyPressEventArgs) Handles TxtBxJPY.KeyPress
        If nonNumberEntered = True Then 'Check for set in the KeyDown event.
            e.Handled = True ' Stop the character from being entered into control since it is non-numerical 
        End If

    End Sub
    Private Sub TxtBxCommissionRates_Keydown(Sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles TxtBxCommissionRates.KeyDown 'Private sub has been made, this allows only numbers into txtSideA.'
        nonNumberEntered = False ' Initialize the flag to false
        If e.KeyCode < Keys.D0 OrElse e.KeyCode > Keys.D9 Then ' Determine whether the keystroke is a number from the top of the keyboard
            If e.KeyCode < Keys.NumPad0 OrElse e.KeyCode > Keys.NumPad9 Then 'Determine whether the keystroke is a number from the keypad
                If e.KeyCode <> Keys.Back Then ' Determine wheether the keystroke is a backspace
                    If e.KeyCode <> Keys.OemPeriod Then ' A non-numerical keystroke was pressed
                        nonNumberEntered = True ' Set the flag to true and evaluate in KeyPress event.
                    End If
                End If
            End If
        End If
    End Sub

    Private Sub TxtBxCommissionRates_KeyPress(Sender As Object, e As System.Windows.Forms.KeyPressEventArgs) Handles TxtBxCommissionRates.KeyPress
        If nonNumberEntered = True Then 'Check for set in the KeyDown event.
            e.Handled = True ' Stop the character from being entered into control since it is non-numerical 
        End If

    End Sub

    Private Sub Form2_Load(sender As Object, e As EventArgs) Handles Me.Load
        Call LoadRates() 'Fetches the converstion rates from the module
        'These are then set to display in textboxes
        TxtBxCommissionRates.Text = PercentageRatesArray(8)
        TxtBxGBP.Text = PercentageRatesArray(0)
        TxtBxEUR.Text = PercentageRatesArray(1)
        TxtBxUSD.Text = PercentageRatesArray(2)
        TxtBxAUD.Text = PercentageRatesArray(3)
        TxtBxNZD.Text = PercentageRatesArray(4)
        TxtBxCAD.Text = PercentageRatesArray(5)
        TxtBxCHF.Text = PercentageRatesArray(6)
        TxtBxJPY.Text = PercentageRatesArray(7)
        'The textboxes are set to not be editable by default
        TxtBxCommissionRates.Enabled = False
        TxtBxGBP.Enabled = False
        TxtBxEUR.Enabled = False
        TxtBxUSD.Enabled = False
        TxtBxAUD.Enabled = False
        TxtBxNZD.Enabled = False
        TxtBxCAD.Enabled = False
        TxtBxCHF.Enabled = False
        TxtBxJPY.Enabled = False
    End Sub
    'Hides this form and navigates back to main form.
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles BtnBack.Click
        Me.Visible = False
        Form1.Visible = True
    End Sub
    'When edit button is clicked the textboxes are enabled so that text can be edited.
    Private Sub BtnEdit_Click(sender As Object, e As EventArgs) Handles BtnEdit.Click
        TxtBxCommissionRates.Enabled = True
        TxtBxGBP.Enabled = True
        TxtBxEUR.Enabled = True
        TxtBxUSD.Enabled = True
        TxtBxAUD.Enabled = True
        TxtBxNZD.Enabled = True
        TxtBxCAD.Enabled = True
        TxtBxCHF.Enabled = True
        TxtBxJPY.Enabled = True
    End Sub

    Private Sub BtnSave_Click(sender As Object, e As EventArgs) Handles BtnSave.Click
        'When the save button is clicked the values stored in the textboxes is set as the array items.
        PercentageRatesArray(0) = TxtBxGBP.Text
        PercentageRatesArray(1) = TxtBxEUR.Text
        PercentageRatesArray(2) = TxtBxUSD.Text
        PercentageRatesArray(3) = TxtBxAUD.Text
        PercentageRatesArray(4) = TxtBxNZD.Text
        PercentageRatesArray(5) = TxtBxCAD.Text
        PercentageRatesArray(6) = TxtBxCHF.Text
        PercentageRatesArray(7) = TxtBxJPY.Text
        PercentageRatesArray(8) = TxtBxCommissionRates.Text
        'These values are then written to the textfile
        Dim Rates As New System.IO.StreamWriter("C:\Users\jake-\Documents\L3 Dip Applied Computing\Unit 614 VB Programs\Westfield Bank Currency Converter\PercentageRates.txt")
        'This for statement iterates 8 times, writing a value on each selected line.
        For i As Integer = 0 To 8
            Rates.WriteLine(PercentageRatesArray(i))
        Next
        'The object file is then closed
        Rates.Close()
        'A message box is displayed to inform the user the details have been saved.
        MsgBox("Rates have been saved", vbInformation)
        'The textboxes are then disabled again so that they cannot be edited.
        TxtBxCommissionRates.Enabled = False
        TxtBxGBP.Enabled = False
        TxtBxEUR.Enabled = False
        TxtBxUSD.Enabled = False
        TxtBxAUD.Enabled = False
        TxtBxNZD.Enabled = False
        TxtBxCAD.Enabled = False
        TxtBxCHF.Enabled = False
        TxtBxJPY.Enabled = False
    End Sub
    'When the reload butto is clicked, the current rates saved in the array are inserted into textboxes.
    Private Sub BtnReload_Click(sender As Object, e As EventArgs) Handles BtnReload.Click
        TxtBxCommissionRates.Text = PercentageRatesArray(8)
        TxtBxGBP.Text = PercentageRatesArray(0)
        TxtBxEUR.Text = PercentageRatesArray(1)
        TxtBxUSD.Text = PercentageRatesArray(2)
        TxtBxAUD.Text = PercentageRatesArray(3)
        TxtBxNZD.Text = PercentageRatesArray(4)
        TxtBxCAD.Text = PercentageRatesArray(5)
        TxtBxCHF.Text = PercentageRatesArray(6)
        TxtBxJPY.Text = PercentageRatesArray(7)
    End Sub
End Class