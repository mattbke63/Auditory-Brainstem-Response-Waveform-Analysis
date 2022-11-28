Imports System.IO
Public Class TDT_CSV
    Dim Currentfile As String
    Private Sub tdt_CSV_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'MsgBox("Good!")
        loadfilename()
        translatefile()
        File.Create(CurDir() & "\Completed.f")
        Close()
    End Sub
    Private Sub loadfilename()
        Dim fPath = CurDir() & "\currentfile.f"
        Dim afile As New IO.StreamReader(fPath, True)
        Currentfile = afile.ReadLine
    End Sub
    Private Sub translatefile()
        File.Delete(CurDir() & "\currentcsv.csv")
        File.Copy(Currentfile, CurDir() & "\currentcsv.csv")
    End Sub
End Class
