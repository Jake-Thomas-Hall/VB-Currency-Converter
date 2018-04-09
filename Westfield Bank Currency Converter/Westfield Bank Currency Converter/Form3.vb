Public Class Form3
    'Hides current form and shows main form.
    Private Sub BtnBack_Click(sender As Object, e As EventArgs) Handles BtnBack.Click
        Me.Visible = False
        Form1.Visible = True
    End Sub
End Class