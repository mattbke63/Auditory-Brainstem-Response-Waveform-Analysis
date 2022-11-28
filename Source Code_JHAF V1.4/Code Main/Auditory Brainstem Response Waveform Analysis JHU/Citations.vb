Public Class Citations
    Private Sub citations_Load(sender As Object, e As EventArgs) Handles Me.Load
        loaddocumentnumber()

    End Sub

    Private Sub loaddocumentnumber()
        If My.Computer.FileSystem.FileExists(CurDir() & "\Citation.txt") = True Then
            Dim fPath = CurDir() & "\Citation.txt"
            Dim afile As New IO.StreamReader(fPath, True)
            TextBox1.Text = afile.ReadLine
            'Label41.Text = afile.ReadLine
        Else
            TextBox1.Text = "Citation information not currently avalible"
        End If
    End Sub


End Class