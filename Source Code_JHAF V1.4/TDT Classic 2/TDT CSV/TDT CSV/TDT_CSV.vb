Imports System.IO
Public Class TDT_CSV
    Dim Currentfile As String
    Dim source As String
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
        Dim filedata As String() = IO.File.ReadAllLines(Currentfile)
        Dim currentline As String
        Dim currentwrite As String
        Dim counter As Integer
        Dim totalcommas As Integer
        Dim topline As String
        Dim totalrows As Integer
        Dim rowcounter As Integer
        File.Delete(CurDir() & "\currentcsv.csv")

        determinesourcetype()

        If source = "CLICKS" Then
            File.Copy(CurDir() & "\mCSVconvertClicks.csv", CurDir() & "\currentcsv.csv")
        Else
            File.Copy(CurDir() & "\mCSVconvertTones.csv", CurDir() & "\currentcsv.csv")
        End If

        totalrows = filedata.Count
        'MsgBox(totalrows)
        rowcounter = 1
        While rowcounter <> totalrows
            currentline = filedata(rowcounter)
            topline = filedata(0)

            If Microsoft.VisualBasic.Right(currentline, 1) <> "," Then
                currentline = currentline & ","
            End If

            If Microsoft.VisualBasic.Right(topline, 1) <> "," Then
                topline = topline & ","
            End If



            totalcommas = Len(currentline) - Len(Replace(currentline, ",", ""))
            'MsgBox(totalcommas)


            If source = "CLICKS" Then
                counter = 0
                While counter <> 11
                    currentwrite = currentwrite & currentline.Split(","c)(counter) & ","
                    counter = counter + 1
                End While

                counter = counter + (totalcommas - (Len(topline) - Len(Replace(topline, ",", ""))))

                While counter <> totalcommas - 1
                    currentwrite = currentwrite & currentline.Split(","c)(counter) & ","
                    counter = counter + 1
                End While
            Else

                counter = 0
                While counter <> 12
                    currentwrite = currentwrite & currentline.Split(","c)(counter) & ","
                    counter = counter + 1
                End While

                counter = counter + (totalcommas - (Len(topline) - Len(Replace(topline, ",", ""))))

                currentwrite = currentwrite & currentline.Split(","c)(counter + 1) & ","
                currentwrite = currentwrite & currentline.Split(","c)(counter) & ","

                counter = counter + 2


                While counter <> totalcommas - 1
                    currentwrite = currentwrite & currentline.Split(","c)(counter) & ","
                    counter = counter + 1
                End While




            End If

            Dim outFile1 As IO.StreamWriter = System.IO.File.AppendText(CurDir() & "\currentcsv.csv")
            outFile1.WriteLine(currentwrite)
            outFile1.Close()

            currentwrite = ""
            rowcounter = rowcounter + 1
        End While
        '


    End Sub


    Private Sub determinesourcetype()
        Dim sourcesub As String() = IO.File.ReadAllLines(Currentfile)
        Dim currentline As String

        currentline = sourcesub(0)
        If InStr(currentline, "Frequency(Hz)") <> 0 Then
            source = "TONES"
        Else
            source = "CLICKS"
        End If

        'MsgBox(currentline.Split(","c)(0)
    End Sub

End Class
