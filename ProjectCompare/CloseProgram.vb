'Title:         Close Program
'Date:          2-27-15
'Author:        Terry Holmes

'Description:   This form will close the program

Option Strict On

Public Class CloseProgram

    Private Sub btnYes_Click(sender As Object, e As EventArgs) Handles btnYes.Click

        'This will close the program
        Logon.Close()
        Me.Close()

    End Sub

    Private Sub btnNo_Click(sender As Object, e As EventArgs) Handles btnNo.Click

        'This will return to the calling form
        Me.Close()

    End Sub
End Class