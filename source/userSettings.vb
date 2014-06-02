Public Class userSettings

    Private Sub frmSettings_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        txtMessage1.Text = My.Settings.MessageText1
        txtMessage2.Text = My.Settings.MessageText2
    End Sub

    Private Sub txtMessage1_Validated(sender As Object, e As EventArgs) Handles txtMessage1.Validated
        My.Settings.MessageText1 = txtMessage1.Text
    End Sub

    Private Sub txtMessage2_Validated(sender As Object, e As EventArgs) Handles txtMessage2.Validated
        My.Settings.MessageText2 = txtMessage2.Text
    End Sub

    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        My.Settings.Save()
        Me.Close()
    End Sub

    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        Me.Close()
    End Sub

End Class