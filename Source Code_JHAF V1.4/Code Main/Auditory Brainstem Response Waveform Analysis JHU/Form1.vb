Imports System.IO

Public Class Form1
    Dim PrgFrom As String
    Dim PrgBack As String
    Dim PrgTo As String
    Dim FileList As New List(Of String)
    Dim FileList2 As New List(Of String)
    Dim PrgCnt As Integer
    Dim CurrentFile As String
    Dim CurrentFile2 As String
    Dim ErrorText As String
    Dim Count As Integer
    Dim thisdate As String
    Dim fileline As Integer
    Dim totalnumber As Integer

    Dim filenamemgb As String
    Dim stimval As String
    Dim dataoutput As String
    Dim concatname As String
    Dim finalwrite As String
    Dim stimulus As String
    Dim linenum As Integer
    Dim freq As String
    Dim finalwrite2 As String
    Dim dataoutputuncalc As String
    Dim DecibleLevelnum As String
    Dim blnFlag As Boolean
    Dim Refreshflag As Boolean
    Dim rewindrefreshflag As Boolean
    Dim rewindloopflag As Boolean
    Dim conversionnum As Double
    Dim settingsname As String
    Dim CSVFILENAME As String
    Dim rewindvalue As Integer
    Dim currentpart As Integer
    Dim stopexecution As String
    Dim voltagemultipler As Integer
    Dim testflag As Boolean
    Dim newrefreshflag As Boolean
    Dim RewindCalcfilename As String
    Dim rewinduncalcfilename As String
    Dim versionnumber As String
    Dim binvaluefirst As Integer
    Dim msgboxflag As Boolean
    Dim previousplotgraphvalue As Integer
    Dim PointerPosOnChart_xAxis_pan As Integer
    Dim PointerPosOnChart_yAxis_pan As Integer






    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click


        If InStr(PrgFrom, "\\") > 0 Or My.Computer.FileSystem.DirectoryExists(PrgFrom) = False Then
            MsgBox("Please Check Folder Name!")
        Else

            'On Error GoTo LoopError

            ErrorText = "Saving Settings"
            savesettings()

            Citations.Close()


            PrgCnt = 0
            FileList.Clear()


            If ComboBox2.Text = "Lauer Lab m File" Then
                lauralmfile()
                ComboBox2.Text = "BioSig RZ (TDT)"
            End If
            'MsgBox("test")

            ErrorText = "enumerating file list"
            FileOpen(1, TextBox1.Text & "\Programs.txt", OpenMode.Output)
            For Each File In IO.Directory.GetFiles(PrgFrom, "*.csv", IO.SearchOption.AllDirectories)

                PrgCnt += 1
                FileList.Add(File)
                WriteLine(1, File)

            Next
            FileClose(1)
            thisdate = DateString & "_" & TimeString
            thisdate = Replace(thisdate, ":", "-")


            ErrorText = "backing up old results and moving new results file in"
            If File.Exists(TextBox1.Text & "\RESULTS.CSV") Then
                File.Move(TextBox1.Text & "\RESULTS.CSV", TextBox1.Text & "\Old Results\RESULTS_" & thisdate & ".CSV")
                File.Delete(TextBox1.Text & "\RESULTS.CSV")
            End If
            File.Copy(CurDir() & "\RESULTS_template.CSV", TextBox1.Text & "\RESULTS.CSV")

            ErrorText = "backing up old results and moving new results uncalc file in"
            If File.Exists(TextBox1.Text & "\RESULTS_uncalc.CSV") Then
                File.Move(TextBox1.Text & "\RESULTS_uncalc.CSV", TextBox1.Text & "\Old Results\RESULTS_uncalc_" & thisdate & ".CSV")
                File.Delete(TextBox1.Text & "\RESULTS_uncalc.CSV")
            End If
            File.Copy(CurDir() & "\RESULTS_template_uncalc.CSV", TextBox1.Text & "\RESULTS_uncalc.CSV")

            Dim outFile121 As IO.StreamWriter = System.IO.File.AppendText(TextBox1.Text & "\RESULTS_uncalc.CSV")
            outFile121.WriteLine("")
            outFile121.Close()

            'rewind value
            rewindvalue = 1
            previousplotgraphvalue = rewindvalue
            currentpart = 1

            'stop execution variable
            stopexecution = False

            'turn on chart
            Chart1.Visible = True
            Label27.Visible = True
            Button5.Visible = True
            Button14.Visible = True
            Button20.Visible = True

            'Turn off settings and execute button
            Button1.Visible = False
            ' Button17.Visible = False
            Label28.Visible = False
            'Label29.Visible = False
            Label30.Visible = False
            Label37.Visible = False
            Label31.Visible = False
            RadioButton1.Visible = False
            RadioButton2.Visible = False
            NumericUpDown11.Visible = False
            CheckBox1.Visible = False
            CheckBox2.Visible = False
            CheckBox3.Visible = False
            CheckBox4.Visible = False
            CheckBox5.Visible = False
            CheckBox6.Visible = False
            CheckBox7.Visible = False
            CheckBox8.Visible = False
            CheckBox9.Visible = False
            CheckBox10.Visible = False
            TextBox2.Visible = False
            Button4.Visible = False
            Button7.Visible = False
            Button8.Visible = False
            ComboBox1.Visible = False
            Label32.Visible = False
            ComboBox2.Visible = False
            Label40.Visible = False

            'GroupBox1.Visible = False
            CheckBox11.Visible = False
            Button10.Visible = False
            Button11.Visible = False
            Button12.Visible = False
            Button18.Visible = False
            GroupBox2.Visible = False
            Label33.Visible = False
            Button15.Visible = False

            'Rewind Files Names
            RewindCalcfilename = TextBox1.Text & "\results.csv"
            rewinduncalcfilename = TextBox1.Text & "\results_uncalc.csv"



            fileline = 1
            For Each File In FileList

                Threading.Thread.Sleep(2)

                'Part count text
                Label35.Text = fileline & "/" & PrgCnt
                ErrorText = "currentfile name export"
                CurrentFile = File
                Dim outFile100 As IO.StreamWriter = System.IO.File.CreateText(CurDir() & "\Currentfile.f")
                'MsgBox(CurrentFile)
                outFile100.WriteLine(CurrentFile)
                outFile100.Close()

                System.Diagnostics.Process.Start(CurDir() & "\" & ComboBox2.Text & ".exe")
                CurrentFile2 = CurDir() & "\currentcsv.csv"

                'MsgBox(CurDir() & "\completed.f")
                While My.Computer.FileSystem.FileExists(CurDir() & "\completed.f") = False
                    Threading.Thread.Sleep(2)

                End While
                Threading.Thread.Sleep(2)


                ErrorText = "Completed Flag"
                My.Computer.FileSystem.DeleteFile(CurDir() & "\completed.f")

                linenum = 1
                ErrorText = "file name extraction error"
                Filenameextraction()

                ErrorText = "error parsing file name"
                SecondtryStimulusextract()

                'End If

loopcsvdecible:
                'button for previous 
                If currentpart = 1 Then
                    Button9.Visible = False
                End If
                If currentpart <> 1 Then
                    Button9.Visible = True
                End If


                ErrorText = "Decible extraction error"
                DecibleLevel()
                If DecibleLevelnum = "" Then
                    GoTo nofeaturesfound
                End If

                ErrorText = "Skipping unchecked Decibles"
                If CheckBox1.Checked = False And DecibleLevelnum <= 10 Then
                    GoTo skipdecible
                End If
                If CheckBox2.Checked = False And 11 <= DecibleLevelnum And DecibleLevelnum <= 20 Then
                    GoTo skipdecible
                End If
                If CheckBox3.Checked = False And 21 <= DecibleLevelnum And DecibleLevelnum <= 30 Then
                    GoTo skipdecible
                End If
                If CheckBox4.Checked = False And 31 <= DecibleLevelnum And DecibleLevelnum <= 40 Then
                    GoTo skipdecible
                End If
                If CheckBox5.Checked = False And 41 <= DecibleLevelnum And DecibleLevelnum <= 50 Then
                    GoTo skipdecible
                End If
                If CheckBox6.Checked = False And 51 <= DecibleLevelnum And DecibleLevelnum <= 60 Then
                    GoTo skipdecible
                End If
                If CheckBox7.Checked = False And 61 <= DecibleLevelnum And DecibleLevelnum <= 70 Then
                    GoTo skipdecible
                End If
                If CheckBox8.Checked = False And 71 <= DecibleLevelnum And DecibleLevelnum <= 80 Then
                    GoTo skipdecible
                End If
                If CheckBox9.Checked = False And 81 <= DecibleLevelnum And DecibleLevelnum <= 90 Then
                    GoTo skipdecible
                End If
                If CheckBox10.Checked = False And DecibleLevelnum >= 91 Then
                    GoTo skipdecible
                End If




                ErrorText = "Total Number of data calculation error"
                totalnumbercalc()
                conversionnum = NumericUpDown11.Value / (totalnumber)

                ErrorText = "SUBID extraction error"
                subID()

                ErrorText = "frequency extraction error"
                frequencynum()

                ErrorText = "Voltage Multiplier error"
                voltagemultiply()

                ErrorText = "datapoint extraction error"
                datapoints()
                'rewindvalue = rewindvalue + 1
                currentpart = currentpart + 1

                If stopexecution = True Then
                    GoTo EndLoop
                End If

                ErrorText = "writing out CSV results"
                finalwrite = filenamemgb & "," & stimval & "," & ",," & "," & freq & "," & DecibleLevelnum & "," & dataoutput & "," & versionnumber & ","
                Dim outFile1 As IO.StreamWriter = System.IO.File.AppendText(TextBox1.Text & "\RESULTS.CSV")
                outFile1.WriteLine(finalwrite)
                outFile1.Close()




                finalwrite2 = filenamemgb & "," & stimval & "," & ",," & "," & freq & "," & DecibleLevelnum & dataoutputuncalc & ","
                Dim outFile12 As IO.StreamWriter = System.IO.File.AppendText(TextBox1.Text & "\RESULTS_uncalc.CSV")
                outFile12.WriteLine(finalwrite2)
                outFile12.Close()

                'MsgBox(concatname)


skipdecible:

                linenum = linenum + 1

                GoTo loopcsvdecible

nofeaturesfound:
                ErrorText = "Error moving Active CSV check to see if raw data file is open"
                If CheckBox11.Checked = True Then
                    movecsv()
                End If
                fileline = fileline + 1
            Next
        End If

        GoTo EndLoop

freqerror:
        MsgBox("Error while parsing file name'" & filenamemgb & "'please ensure frequency is in the right column, dont change the format dummy!")
            GoTo EndLoop

parseerror:

        MsgBox("Error while parsing file name '" & filenamemgb & "' please ensure clicks or tones is spelled correctly and reboot program, if thats not the issue contact your amazing brother to fix his code")
        GoTo EndLoop
LoopError:
        ' MsgBox(My.Computer.FileSystem.GetFileInfo(CurrentFile).Length & "     " & My.Computer.FileSystem.ReadAllBytes(CurrentFile2).Length)
        MsgBox("Error while " & ErrorText & ":" & Chr(13) & CurrentFile2 & Chr(13) & Err.Description)
EndLoop:
        File.Delete(TextBox1.Text & "\Programs.txt")
        MsgBox("Program Complete")

        'turn on settings and execute button
        Button1.Visible = True
        'Button17.Visible = True
        Label28.Visible = True
        'Label29.Visible = True
        Label30.Visible = True
        Label37.Visible = True

        Label38.Visible = False
        Label39.Visible = False

        Label31.Visible = True
        RadioButton1.Visible = True
        RadioButton2.Visible = True
        NumericUpDown11.Visible = True
        CheckBox1.Visible = True
        CheckBox2.Visible = True
        CheckBox3.Visible = True
        CheckBox4.Visible = True
        CheckBox5.Visible = True
        CheckBox6.Visible = True
        CheckBox7.Visible = True
        CheckBox8.Visible = True
        CheckBox9.Visible = True
        CheckBox10.Visible = True
        TextBox2.Visible = True
        Button4.Visible = True
        Button7.Visible = True
        Button8.Visible = True
        ComboBox1.Visible = True
        Label32.Visible = True
        ComboBox2.Visible = True
        Label40.Visible = True
        'GroupBox1.Visible = True
        CheckBox11.Visible = True
        Button10.Visible = True
        Button11.Visible = True
        Button12.Visible = True
        Button18.Visible = True
        GroupBox2.Visible = True
        Label33.Visible = True
        Button15.Visible = True
        Label35.Text = "0/0"

        'hide chart
        Chart1.Visible = False
        Label27.Visible = False
        Button5.Visible = False
        Button9.Visible = False
        Button14.Visible = False
        Button19.Visible = False
        Button20.Visible = False

    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles Me.Load
        If My.Computer.FileSystem.FileExists(CurDir() & "\currentfolder.txt") = False Then
            Dim fPath = CurDir() & "\currentfolder.txt"
            Dim afile As New IO.StreamWriter(fPath, True)
            afile.WriteLine("D:\test")
            MsgBox("")
            afile.Close()
        End If
        Dim linesstimulus_ As String() = IO.File.ReadAllLines(CurDir() & "\currentfolder.txt")
        Dim test As String
        test = linesstimulus_(0)

        versionnumber = "Version 1.4"
        Label36.Text = versionnumber


        'this is the data for what is read in the from
        TextBox1.Text = test
        Label2.Text = TextBox1.Text & "\Place CSV Here\"
        Label3.Text = TextBox1.Text & "\Old Results\"

        Label5.Text = "CSV Files for evaluation:"
        Label6.Text = "Old Generated Results:"

        'Button4.Visible = False

        settingsname = "Config"
        loadsettings()

        comboboxpopulate()
        Chart1.Visible = False
        Label27.Visible = False
        Button5.Visible = False
        Button9.Visible = False
        Button14.Visible = False

        Label38.Visible = False
        Label39.Visible = False
        Button19.Visible = False
        Button20.Visible = False

        ErrorText = "Setting Algorithum"
        algorithmumset()

        Label13.Text = "0"
        Label14.Text = "0"
        Label15.Text = "0"
        Label16.Text = "0"
        Label17.Text = "0"

        Label22.Text = "0"
        Label23.Text = "0"
        Label24.Text = "0"
        Label25.Text = "0"
        Label26.Text = "0"



    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Close()

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

        FolderBrowserDialog1.RootFolder = Environment.SpecialFolder.MyComputer

        If FolderBrowserDialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
            Try
                TextBox1.Text = FolderBrowserDialog1.SelectedPath
                Dim outFile121 As IO.StreamWriter = System.IO.File.CreateText(CurDir() & "\currentfolder.txt")
                outFile121.WriteLine(TextBox1.Text)
                outFile121.Close()
            Catch Ex As Exception
                MessageBox.Show("Cannot open folder. Original error: " & Ex.Message)
            End Try
        End If

        Label2.Text = TextBox1.Text & "\Place CSV Here\"
        Label3.Text = TextBox1.Text & "\Old Results\"

        PrgFrom = Label2.Text
        PrgBack = Label3.Text


    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

        Label2.Text = TextBox1.Text & "\Place CSV Here\"
        Label3.Text = TextBox1.Text & "\Old Results\"

        PrgFrom = Label2.Text
        PrgBack = Label3.Text


    End Sub


    Private Sub Filenameextraction()

        Dim dieliminator As Integer
        Dim textcounter As Integer
        Dim tempname As String
        Dim tempname1 As String

        dieliminator = 0
        textcounter = 4

        While dieliminator = 0
            tempname = Microsoft.VisualBasic.Right(CurrentFile, textcounter)
            tempname1 = Microsoft.VisualBasic.Left(tempname, 1)

            If tempname1 = "\" Then
                dieliminator = 1
            End If

            textcounter = textcounter + 1
        End While

        textcounter = textcounter - 2


        filenamemgb = Microsoft.VisualBasic.Right(CurrentFile, textcounter)
        filenamemgb = Microsoft.VisualBasic.Left(filenamemgb, textcounter - 4)

    End Sub

    Private Sub subID()

        Dim linesstimulus As String() = IO.File.ReadAllLines(CurrentFile2)
        Dim line2 As String
        Dim commas As Integer
        Dim linecounter As Integer
        Dim tempnamestim As String
        Dim tempname1stim As String
        Dim comma7 As Integer
        Dim triggerpoint As Integer

        line2 = linesstimulus(1)

        linecounter = 1
        commas = 0
        comma7 = 0
        triggerpoint = 0

        While commas < 8
            tempnamestim = Microsoft.VisualBasic.Left(line2, linecounter)
            tempname1stim = Microsoft.VisualBasic.Right(tempnamestim, 1)

            If tempname1stim = "," Then
                commas = commas + 1
            End If

            If commas = 7 And triggerpoint = 0 Then
                comma7 = linecounter
                triggerpoint = 1
            End If


            linecounter = linecounter + 1
        End While

        linecounter = linecounter - 2
        comma7 = comma7

        stimval = Microsoft.VisualBasic.Left(line2, linecounter)
        'MsgBox(linecounter - comma7)
        stimval = Microsoft.VisualBasic.Right(stimval, linecounter - comma7)
        'MsgBox(stimval)

    End Sub

    Private Sub datapoints()
        Dim linesstimulus_ As String() = IO.File.ReadAllLines(CurrentFile2)
        Dim line2_ As String
        Dim commas_ As Integer
        Dim linecounter_ As Integer
        Dim tempnamestim_ As String
        Dim tempname1stim_ As String
        Dim comma7_ As Integer
        Dim stimval_ As String
        Dim triggerpoint_ As Integer
        Dim lastpoint As Integer
        Dim numbers(totalnumber) As Decimal
        Dim matrixnum As Integer
        Dim converttoint As Decimal
        Dim counter As Integer
        Dim maxnumbers(4) As Decimal
        Dim maxlocation(4) As Integer
        Dim currentlocation As Integer
        Dim lastnumb As Integer
        Dim minnumbers(4) As Decimal
        Dim minlocation(4) As Integer
        Dim trigger As Integer
        Dim tempmax As Decimal
        Dim tempmin As Integer
        Dim tempmaxcounter As Integer
        Dim tempmincounter As Integer
        Dim mgb As Integer
        Dim leftsideeq As Decimal
        Dim rightsideeq As Decimal
        Dim freqoverride As Integer
        Dim searchoverride As Integer
        Dim Sylveon As Integer
        Dim godsomanyvars As Integer
        Dim somanycounters As Integer
        Dim trigger2 As Integer
        Dim newchartheader As String


        blnFlag = False
        Refreshflag = False
        rewindrefreshflag = False
        rewindloopflag = False

        'these are starting parameters for starting to go through the filter below.
        somanycounters = 0
        freqoverride = 0
        searchoverride = 200
        godsomanyvars = 0

startloopover:
        'these values override search values if they exceed total number of values
        If searchoverride > totalnumber - 5 Then
            searchoverride = totalnumber - 5
        End If


        'this is the line of code the determines what row in the CSV it searches through
        line2_ = linesstimulus_(linenum)

        'This checks to make sure the data set ends in a comma EXCEL REMOVES LAST COMMA IF SAVED IN EXCEL
        If Microsoft.VisualBasic.Right(line2_, 1) <> "," Then
            line2_ = line2_ & ","
            'MsgBox("")
        End If

        'These are counters
        linecounter_ = 1
        commas_ = 0
        comma7_ = 0
        triggerpoint_ = 0
        matrixnum = 0
        trigger = 0


        'this loop filters through the first lines in the CSV till it hits desired column
        While commas_ <= 47
            tempnamestim_ = Microsoft.VisualBasic.Left(line2_, linecounter_)
            tempname1stim_ = Microsoft.VisualBasic.Right(tempnamestim_, 1)

            If tempname1stim_ = "," Then
                commas_ = commas_ + 1
            End If

            linecounter_ = linecounter_ + 1
        End While
        lastpoint = linecounter_

        'this loop puts all data for analysis into an array
        While commas_ <= totalnumber + 48

            tempnamestim_ = Microsoft.VisualBasic.Left(line2_, linecounter_)
            tempname1stim_ = Microsoft.VisualBasic.Right(tempnamestim_, 1)


            If tempname1stim_ = "," Then
                stimval_ = Microsoft.VisualBasic.Left(line2_, linecounter_ - 1)
                stimval_ = Microsoft.VisualBasic.Right(stimval_, linecounter_ - lastpoint)
                converttoint = CDec(stimval_)
                'MsgBox(voltagemultipler)
                numbers(matrixnum) = converttoint * voltagemultipler
                matrixnum = matrixnum + 1
                lastpoint = linecounter_ + 1
                commas_ = commas_ + 1
            End If


            linecounter_ = linecounter_ + 1
        End While


        'these are counters
        counter = 10
        lastnumb = 0
        currentlocation = 0
        triggerpoint_ = 0
        tempmin = 0
        tempmax = -100000
        tempmaxcounter = 0
        trigger = 0


        'this look goes through all instances of volcalizations
        While counter < searchoverride - 5
            If freqoverride = 0 Then
                If freq = "Click" Then
                    mgb = 650
                End If
            End If

            'this breaks away when 5 vocalizations are reached
            If currentlocation = 5 Then
                GoTo fullbreak

            End If


            'these are the values that are used for the filter to deterimine the slow
            leftsideeq = (numbers(counter - 5) - numbers(counter - 4)) + (numbers(counter - 4) - numbers(counter - 3)) + (numbers(counter - 3) - numbers(counter - 2)) + (numbers(counter - 2) - numbers(counter - 1)) + (numbers(counter - 1) - numbers(counter))
            rightsideeq = (numbers(counter) - numbers(counter + 1)) + (numbers(counter + 1) - numbers(counter + 2)) + (numbers(counter + 2) - numbers(counter + 3)) + (numbers(counter + 3) - numbers(counter + 4)) + (numbers(counter + 4) - numbers(counter + 5))


            'this loop filters by looking at previous points and future points to determine if the average slope is higher
            If leftsideeq < rightsideeq And leftsideeq <= 0 Then

                'this value determines if the trigger has been met if it is truly a max value
                If tempmax < numbers(counter) Then
                    tempmax = numbers(counter)
                    tempmaxcounter = counter
                    triggerpoint_ = 1
                    trigger = 0
                    'MsgBox("triggermax: " & counter)
                End If


                lastnumb = counter

            End If

            'this checks to see if maximum point has been seen
            If triggerpoint_ = 1 Then
                'this checks to see if the slope for minimum is met
                If leftsideeq > rightsideeq And rightsideeq <= 0 Then
                    'this checks for futrure minumum points
                    If numbers(counter) < numbers(counter + 1) And numbers(counter) < numbers(counter + 2) And numbers(counter) < numbers(counter + 3) And numbers(counter) < numbers(counter + 4) And numbers(counter) < numbers(counter + 5) Then
                        trigger = 1
                        'MsgBox("low point trigger: " & counter)
                    End If
                End If
            End If


            'these are the threshholds for different frequencies
            If trigger = 1 And triggerpoint_ = 1 Then
                If freq <> "Click" And freqoverride = 0 Then
                    'This is the formula used to determine cut off for frequencies to stop false triggers
                    'this was determined by a 250 cut off at 42000 and a cut off of 400 at 8000 and is a linear slope across all frequencies
                    If freq = "Unknown Stimulus" Then
                        mgb = ((-3 / 680) * 16000 + 7400 / 17) * (DecibleLevelnum / 90)
                    Else
                        mgb = ((-3 / 680) * CInt(freq) + 7400 / 17) * (DecibleLevelnum / 90)
                    End If

                End If







                If tempmax - numbers(counter) > mgb Then

                    'this stores all max vocalizations in a matrix
                    maxnumbers(currentlocation) = tempmax
                    maxlocation(currentlocation) = tempmaxcounter

                    'this stores all min localizations in a matrix
                    minnumbers(currentlocation) = numbers(counter)
                    minlocation(currentlocation) = counter


                    currentlocation = currentlocation + 1

                    'MsgBox(filenamemgb)
                    'MsgBox("taken data point high: " & tempmaxcounter)

                End If
                trigger = 0
                triggerpoint_ = 0
                tempmax = -100000
            End If



            counter = counter + 1
        End While



        'this sets the original threshold value to look at later
        If godsomanyvars = 0 Then
            Sylveon = mgb
            godsomanyvars = godsomanyvars + 1

        End If

        'if the original thresholds is decreased by 1/3 of the normal value this will set threshold back to orignal and call it done
        If Sylveon * (1 / 2) > mgb Then
            Sylveon = mgb
            somanycounters = 1
            'test mgb no longer have script go back to original settings and us tuned down settings

            GoTo fullbreak

        End If
        If somanycounters = 1 Then
            GoTo fullbreak
        End If


        If maxlocation(4) = 0 Then
            freqoverride = 1
            mgb = mgb - (Sylveon * 0.05)
            searchoverride = searchoverride + 10
            Array.Clear(maxlocation, 0, maxlocation.Length)
            Array.Clear(minlocation, 0, minlocation.Length)
            Array.Clear(maxnumbers, 0, maxnumbers.Length)
            Array.Clear(minnumbers, 0, minnumbers.Length)
            GoTo startloopover
        End If


fullbreak:



        'this is for calcualtions for uncalc'd data
        counter = 0
        dataoutputuncalc = ""
        While counter <= totalnumber

            dataoutputuncalc = dataoutputuncalc & "," & CStr(numbers(counter))
            'MsgBox(dataoutputuncalc)
            counter = counter + 1
        End While

        'skips the manual method of changing data 
        If RadioButton2.Checked = True Then
            GoTo skipcharting
        End If

repopgraph:
        'this is the data to calculate the chart
        Dim test As Integer
        test = 0
        'MsgBox(totalnumber)
        totalnumber = UBound(numbers)
        While test < totalnumber
            Me.Chart1.Series("Amplitude (nV)").Points.AddXY(test, numbers(test))
            test = test + 1
        End While



        newchartheader = Replace(DecibleLevelnum, "+", "")
        If Len(newchartheader) <> Len(DecibleLevelnum) Then
            newchartheader = newchartheader & " dB atten"
        End If


        Me.Chart1.Titles.Add("Sub ID: " & stimval & "         Decibel Level:" & newchartheader & "         Frequency Number: " & freq)
        Me.Chart1.ChartAreas("ChartArea1").AxisX.Title = "Bin Number"
        Me.Chart1.ChartAreas("ChartArea1").AxisY.Title = "Amplitude (nV)"


        Me.Chart1.Series("MAX").Points.AddXY(maxlocation(0), maxnumbers(0))
        Me.Chart1.Series("MAX").Points.AddXY(maxlocation(1), maxnumbers(1))
        Me.Chart1.Series("MAX").Points.AddXY(maxlocation(2), maxnumbers(2))
        Me.Chart1.Series("MAX").Points.AddXY(maxlocation(3), maxnumbers(3))
        Me.Chart1.Series("MAX").Points.AddXY(maxlocation(4), maxnumbers(4))

        Me.Chart1.Series("MIN").Points.AddXY(minlocation(0), minnumbers(0))
        Me.Chart1.Series("MIN").Points.AddXY(minlocation(1), minnumbers(1))
        Me.Chart1.Series("MIN").Points.AddXY(minlocation(2), minnumbers(2))
        Me.Chart1.Series("MIN").Points.AddXY(minlocation(3), minnumbers(3))
        Me.Chart1.Series("MIN").Points.AddXY(minlocation(4), minnumbers(4))
        'Me.chart1.Series("MIN").Points.AddXY(totalnumber, numbers(totalnumber - 1))




        Label13.Text = maxnumbers(0)
        Label14.Text = maxnumbers(1)
        Label15.Text = maxnumbers(2)
        Label16.Text = maxnumbers(3)
        Label17.Text = maxnumbers(4)

        NumericUpDown1.Value = maxlocation(0)
        NumericUpDown2.Value = maxlocation(1)
        NumericUpDown3.Value = maxlocation(2)
        NumericUpDown4.Value = maxlocation(3)
        NumericUpDown5.Value = maxlocation(4)

        Label26.Text = minnumbers(0)
        Label25.Text = minnumbers(1)
        Label24.Text = minnumbers(2)
        Label23.Text = minnumbers(3)
        Label22.Text = minnumbers(4)

        NumericUpDown10.Value = minlocation(0)
        NumericUpDown9.Value = minlocation(1)
        NumericUpDown8.Value = minlocation(2)
        NumericUpDown7.Value = minlocation(3)
        NumericUpDown6.Value = minlocation(4)


        trigger2 = 0



        Do Until blnFlag = True


            Me.Show()
            Application.DoEvents()
            binerrorcheck()


            'wait loop while in rewind condition to return values back to normal

            If rewindloopflag = True Then

                Label13.Text = maxnumbers(0)
                Label14.Text = maxnumbers(1)
                Label15.Text = maxnumbers(2)
                Label16.Text = maxnumbers(3)
                Label17.Text = maxnumbers(4)

                NumericUpDown1.Value = maxlocation(0)
                NumericUpDown2.Value = maxlocation(1)
                NumericUpDown3.Value = maxlocation(2)
                NumericUpDown4.Value = maxlocation(3)
                NumericUpDown5.Value = maxlocation(4)

                Label26.Text = minnumbers(0)
                Label25.Text = minnumbers(1)
                Label24.Text = minnumbers(2)
                Label23.Text = minnumbers(3)
                Label22.Text = minnumbers(4)

                NumericUpDown10.Value = minlocation(0)
                NumericUpDown9.Value = minlocation(1)
                NumericUpDown8.Value = minlocation(2)
                NumericUpDown7.Value = minlocation(3)
                NumericUpDown6.Value = minlocation(4)
                rewindloopflag = False

                trigger2 = trigger2 + 1
            End If


            If Refreshflag = True Then
                Refreshflag = False




                maxnumbers(0) = numbers(NumericUpDown1.Value)
                maxnumbers(1) = numbers(NumericUpDown2.Value)
                maxnumbers(2) = numbers(NumericUpDown3.Value)
                maxnumbers(3) = numbers(NumericUpDown4.Value)
                maxnumbers(4) = numbers(NumericUpDown5.Value)

                maxlocation(0) = NumericUpDown1.Value
                maxlocation(1) = NumericUpDown2.Value
                maxlocation(2) = NumericUpDown3.Value
                maxlocation(3) = NumericUpDown4.Value
                maxlocation(4) = NumericUpDown5.Value

                minnumbers(0) = numbers(NumericUpDown10.Value)
                minnumbers(1) = numbers(NumericUpDown9.Value)
                minnumbers(2) = numbers(NumericUpDown8.Value)
                minnumbers(3) = numbers(NumericUpDown7.Value)
                minnumbers(4) = numbers(NumericUpDown6.Value)

                minlocation(0) = NumericUpDown10.Value
                minlocation(1) = NumericUpDown9.Value
                minlocation(2) = NumericUpDown8.Value
                minlocation(3) = NumericUpDown7.Value
                minlocation(4) = NumericUpDown6.Value


                clearchart()

                GoTo repopgraph
            End If
            If stopexecution = True Then
                GoTo othername2
            End If
        Loop



        clearchart()


skipcharting:


        If maxlocation(4) = 0 And maxlocation(3) = 0 Then
            dataoutput = CStr(maxnumbers(0) - minnumbers(0)) & "," & CStr(maxlocation(0) * conversionnum) & "," & CStr(maxnumbers(1) - minnumbers(1)) & "," & CStr(maxlocation(1) * conversionnum) & "," & CStr(maxnumbers(2) - minnumbers(2)) & "," & CStr(maxlocation(2) * conversionnum) & "," & "0" & "," & "0" & "," & "0" & "," & "0"

            GoTo othername2
        End If

        If maxlocation(4) = 0 Then
            dataoutput = CStr(maxnumbers(0) - minnumbers(0)) & "," & CStr(maxlocation(0) * conversionnum) & "," & CStr(maxnumbers(1) - minnumbers(1)) & "," & CStr(maxlocation(1) * conversionnum) & "," & CStr(maxnumbers(2) - minnumbers(2)) & "," & CStr(maxlocation(2) * conversionnum) & "," & CStr(maxnumbers(3) - minnumbers(3)) & "," & CStr(maxlocation(3) * conversionnum) & "," & "0" & "," & "0"

            GoTo othername2
        End If


        dataoutput = CStr(maxnumbers(0) - minnumbers(0)) & "," & CStr(maxlocation(0) * conversionnum) & "," & CStr(maxnumbers(1) - minnumbers(1)) & "," & CStr(maxlocation(1) * conversionnum) & "," & CStr(maxnumbers(2) - minnumbers(2)) & "," & CStr(maxlocation(2) * conversionnum) & "," & CStr(maxnumbers(3) - minnumbers(3)) & "," & CStr(maxlocation(3) * conversionnum) & "," & CStr(maxnumbers(4) - minnumbers(4)) & "," & CStr(maxlocation(4) * conversionnum)
        'MsgBox(dataoutput)

othername2:

        dataoutput = dataoutput & "," & NumericUpDown11.Value & "," & maxlocation(0) & "," & minlocation(0) & "," & maxlocation(1) & "," & minlocation(1) & "," & maxlocation(2) & "," & minlocation(2) & "," & maxlocation(3) & "," & minlocation(3) & "," & maxlocation(4) & "," & minlocation(4) & ",,"

        'this does the interpeak latencies
        counter = 0
        Do Until counter = 4
            If maxlocation(counter) = 0 Then
                dataoutput = dataoutput & "0,"
            ElseIf maxlocation(counter + 1) = 0 Then
                dataoutput = dataoutput & "0,"
            Else
                dataoutput = dataoutput & CStr((maxlocation(counter + 1) - maxlocation(counter)) * conversionnum) & ","
                'MsgBox(conversionnum * 365)
            End If
            counter = counter + 1
        Loop

        'last 2 interpeak latencies
        If maxlocation(0) = 0 Then
            dataoutput = dataoutput & "0,"
        ElseIf maxlocation(3) = 0 Then
            dataoutput = dataoutput & "0,"
        Else
            dataoutput = dataoutput & CStr((maxlocation(3) - maxlocation(0)) * conversionnum) & ","
        End If

        If maxlocation(0) = 0 Then
            dataoutput = dataoutput & "0,"
        ElseIf maxlocation(4) = 0 Then
            dataoutput = dataoutput & "0,"
        Else
            dataoutput = dataoutput & CStr((maxlocation(4) - maxlocation(0)) * conversionnum) & ","
        End If

        dataoutput = dataoutput & ","

        'This does the Amplitude Ratios
        counter = 1
        Do Until counter = 5
            If maxlocation(0) = 0 Then
                dataoutput = dataoutput & "0,"
            ElseIf maxlocation(counter) = 0 Then
                dataoutput = dataoutput & "0,"
            Else
                dataoutput = dataoutput & CStr(maxnumbers(counter) - minnumbers(counter)) / (maxnumbers(0) - minnumbers(0)) & ","
            End If
            counter = counter + 1
        Loop

exitloop:


    End Sub

    Private Sub frequencynum()
        Dim linesstimulus_ As String() = IO.File.ReadAllLines(CurrentFile2)
        Dim line2_ As String
        Dim commas_ As Integer
        Dim linecounter_ As Integer
        Dim tempnamestim_ As String
        Dim tempname1stim_ As String
        Dim comma7_ As Integer
        Dim triggerpoint_ As Integer
        Dim matrixnum As Integer


        If stimulus = "Click" Then
            freq = "Click"
            GoTo endofprg
        End If

        If stimulus = "Tones" Then

            line2_ = linesstimulus_(linenum)

            linecounter_ = 1
            commas_ = 0
            comma7_ = 0
            triggerpoint_ = 0
            matrixnum = 0

            'this loop filters through the first lines in the CSV till it hits desired column
            While commas_ <= 12
                tempnamestim_ = Microsoft.VisualBasic.Left(line2_, linecounter_)
                tempname1stim_ = Microsoft.VisualBasic.Right(tempnamestim_, 1)

                If tempname1stim_ = "," Then
                    commas_ = commas_ + 1
                End If

                If commas_ = 11 Then
                    comma7_ = linecounter_
                End If

                linecounter_ = linecounter_ + 1
            End While

            linecounter_ = linecounter_ - 2

            freq = Microsoft.VisualBasic.Left(line2_, linecounter_)
            freq = Microsoft.VisualBasic.Right(freq, linecounter_ - comma7_ - 1)
            'MsgBox(freq)



            GoTo endofprg
        End If

        freq = "Unknown Stimulus"
endofprg:
    End Sub

    Private Sub SecondtryStimulusextract()

        Dim linesstimulus As String() = IO.File.ReadAllLines(CurrentFile2)
        Dim line2 As String

        line2 = linesstimulus(0)

        If InStr(line2, "Freq(Hz)") Then
            stimulus = "Tones"

        Else
            stimulus = "Click"

        End If

    End Sub

    Private Sub totalnumbercalc()
        Dim linesstimulus_ As String() = IO.File.ReadAllLines(CurrentFile2)
        Dim line2_ As String
        Dim commas_ As Integer
        Dim linecounter_ As Integer
        Dim tempnamestim_ As String
        Dim tempname1stim_ As String
        Dim comma7_ As Integer
        Dim triggerpoint_ As Integer
        Dim matrixnum As Integer
        Dim temp1 As String
        Dim testlineforcomma As String


        line2_ = linesstimulus_(linenum)

        linecounter_ = 1
        commas_ = 0
        comma7_ = 0
        triggerpoint_ = 0
        matrixnum = 0


        testlineforcomma = Microsoft.VisualBasic.Right(line2_, 1)
        If testlineforcomma <> "," Then
            line2_ = line2_ & ","
            'MsgBox("")
        End If


        'MsgBox("test")
        'this loop filters through the first lines in the CSV till it hits desired column
        While commas_ <= 45
            tempnamestim_ = Microsoft.VisualBasic.Left(line2_, linecounter_)
            tempname1stim_ = Microsoft.VisualBasic.Right(tempnamestim_, 1)

            If tempname1stim_ = "," Then
                commas_ = commas_ + 1
            End If

            If commas_ = 44 Then
                comma7_ = linecounter_
            End If

            linecounter_ = linecounter_ + 1
        End While

        linecounter_ = linecounter_ - 2
        'MsgBox(line2_)
        temp1 = Microsoft.VisualBasic.Left(line2_, linecounter_)
        'MsgBox(temp1)
        temp1 = Microsoft.VisualBasic.Right(temp1, linecounter_ - comma7_ - 1)
        'MsgBox(temp1)
        totalnumber = temp1 - 1
    End Sub

    Private Sub totalnumbercalcrewind()
        Dim linesstimulus_ As String() = IO.File.ReadAllLines(rewinduncalcfilename)
        Dim line2_ As String
        Dim commas_ As Integer
        Dim tempnamestim_ As String



        line2_ = linesstimulus_(rewindvalue)

        'MsgBox(line2_)
        commas_ = Len(line2_)

        tempnamestim_ = Replace(line2_, ",", "")

        totalnumber = commas_ - Len(tempnamestim_) - 8


        'MsgBox(totalnumber)
    End Sub

    Private Sub DecibleLevel()
        Dim linesstimulus_ As String() = IO.File.ReadAllLines(CurrentFile2)
        Dim line2_ As String
        Dim commas_ As Integer
        Dim linecounter_ As Integer
        Dim tempnamestim_ As String
        Dim tempname1stim_ As String
        Dim comma7_ As Integer
        Dim triggerpoint_ As Integer
        Dim Deciblenumber As Integer
        Dim temp1 As String

        On Error GoTo nodatafound
        line2_ = linesstimulus_(linenum)

        linecounter_ = 1
        commas_ = 0
        comma7_ = 0
        triggerpoint_ = 0


        If stimulus = "Tones" Then
            Deciblenumber = 13
        End If

        If stimulus = "Click" Then
            Deciblenumber = 12
        End If
        'MsgBox(Deciblenumber)
        'MsgBox("test")
        'this loop filters through the first lines in the CSV till it hits desired column
        While commas_ <= Deciblenumber
            tempnamestim_ = Microsoft.VisualBasic.Left(line2_, linecounter_)
            tempname1stim_ = Microsoft.VisualBasic.Right(tempnamestim_, 1)

            If tempname1stim_ = "," Then
                commas_ = commas_ + 1
            End If

            If commas_ = Deciblenumber - 1 Then
                comma7_ = linecounter_
            End If

            linecounter_ = linecounter_ + 1
        End While

        linecounter_ = linecounter_ - 2
        'MsgBox(line2_)
        temp1 = Microsoft.VisualBasic.Left(line2_, linecounter_)
        'MsgBox(temp1)
        temp1 = Microsoft.VisualBasic.Right(temp1, linecounter_ - comma7_ - 1)
        'MsgBox(temp1)


nodatafound:
        DecibleLevelnum = temp1

    End Sub

    Private Sub opencsv()
        Dim proc As New System.Diagnostics.Process()
        proc = Process.Start(CSVFILENAME)

    End Sub

    Private Sub movecsv()
        Dim testmgb As Integer
        Dim filename As String

        testmgb = InStrRev(CurrentFile, "\")

        filename = Microsoft.VisualBasic.Right(CurrentFile, Len(CurrentFile) - testmgb)
        File.Move(CurrentFile, TextBox1.Text & "\Old Results\" & thisdate & "_" & filename)
        File.Delete(CurrentFile)

    End Sub

    Private Sub voltagemultiply()

        If RadioButton3.Checked = True Then
            voltagemultipler = 1
        End If
        If RadioButton4.Checked = True Then
            voltagemultipler = 10 ^ 3
        End If
        If RadioButton5.Checked = True Then
            voltagemultipler = 10 ^ 6
        End If
        If RadioButton6.Checked = True Then
            voltagemultipler = 10 ^ 9
        End If

    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click


        If rewindvalue >= 1 Then
            Button9.Visible = True
        End If


        testflag = True

        rewindvalue = rewindvalue + 1
        previousplotgraphvalue = rewindvalue
        If rewindvalue - 1 = currentpart Then
            blnFlag = True
        ElseIf rewindvalue = currentpart Then

            Refreshflag = True
        Else

            refreshrewindgraph()

        End If

        Chart1.ChartAreas(0).AxisX.ScaleView.ZoomReset()
        Chart1.ChartAreas(0).AxisY.ScaleView.ZoomReset()

    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        testflag = True
        rewindvalue = rewindvalue + 1
        previousplotgraphvalue = rewindvalue

        Chart1.ChartAreas(0).AxisX.ScaleView.ZoomReset()
        Chart1.ChartAreas(0).AxisY.ScaleView.ZoomReset()
    End Sub



    Private Sub NumericUpDown1_ValueChanged(sender As Object, e As EventArgs) Handles NumericUpDown1.ValueChanged
        If currentpart = rewindvalue Then
            Refreshflag = True
        End If
        If rewindloopflag = True Then

            newrefreshflag = True
        End If
        If NumericUpDown1.Value >= totalnumber Then
            NumericUpDown1.Value = totalnumber
        End If
    End Sub
    Private Sub NumericUpDown2_ValueChanged(sender As Object, e As EventArgs) Handles NumericUpDown2.ValueChanged
        If currentpart = rewindvalue Then
            Refreshflag = True
        End If
        If rewindloopflag = True Then
            'MsgBox("MGB")
            newrefreshflag = True
        End If
        If NumericUpDown2.Value >= totalnumber Then
            NumericUpDown2.Value = totalnumber
        End If
    End Sub
    Private Sub NumericUpDown3_ValueChanged(sender As Object, e As EventArgs) Handles NumericUpDown3.ValueChanged
        If currentpart = rewindvalue Then
            Refreshflag = True
        End If
        If rewindloopflag = True Then
            'MsgBox("MGB")
            newrefreshflag = True
        End If
        If NumericUpDown3.Value >= totalnumber Then
            NumericUpDown3.Value = totalnumber
        End If
    End Sub
    Private Sub NumericUpDown4_ValueChanged(sender As Object, e As EventArgs) Handles NumericUpDown4.ValueChanged
        If currentpart = rewindvalue Then
            Refreshflag = True
        End If
        If rewindloopflag = True Then
            'MsgBox("MGB")
            newrefreshflag = True
        End If
        If NumericUpDown4.Value >= totalnumber Then
            NumericUpDown4.Value = totalnumber
        End If
    End Sub
    Private Sub NumericUpDown5_ValueChanged(sender As Object, e As EventArgs) Handles NumericUpDown5.ValueChanged
        If currentpart = rewindvalue Then
            Refreshflag = True
        End If
        If rewindloopflag = True Then
            'MsgBox("MGB")
            newrefreshflag = True
        End If
        If NumericUpDown5.Value >= totalnumber Then
            NumericUpDown5.Value = totalnumber
        End If
    End Sub
    Private Sub NumericUpDown6_ValueChanged(sender As Object, e As EventArgs) Handles NumericUpDown6.ValueChanged
        If currentpart = rewindvalue Then
            Refreshflag = True
        End If
        If rewindloopflag = True Then
            'MsgBox("MGB")
            newrefreshflag = True
        End If
        If NumericUpDown6.Value >= totalnumber Then
            NumericUpDown6.Value = totalnumber
        End If
    End Sub
    Private Sub NumericUpDown7_ValueChanged(sender As Object, e As EventArgs) Handles NumericUpDown7.ValueChanged
        If currentpart = rewindvalue Then
            Refreshflag = True
        End If
        If rewindloopflag = True Then
            'MsgBox("MGB")
            newrefreshflag = True
        End If
        If NumericUpDown7.Value >= totalnumber Then
            NumericUpDown7.Value = totalnumber
        End If
    End Sub
    Private Sub NumericUpDown8_ValueChanged(sender As Object, e As EventArgs) Handles NumericUpDown8.ValueChanged
        If currentpart = rewindvalue Then
            Refreshflag = True
        End If
        If rewindloopflag = True Then
            'MsgBox("MGB")
            newrefreshflag = True
        End If
        If NumericUpDown8.Value >= totalnumber Then
            NumericUpDown8.Value = totalnumber
        End If
    End Sub
    Private Sub NumericUpDown9_ValueChanged(sender As Object, e As EventArgs) Handles NumericUpDown9.ValueChanged
        If currentpart = rewindvalue Then
            Refreshflag = True
        End If
        If rewindloopflag = True Then
            'MsgBox("MGB")
            newrefreshflag = True
        End If
        If NumericUpDown9.Value >= totalnumber Then
            NumericUpDown9.Value = totalnumber
        End If
    End Sub
    Private Sub NumericUpDown10_ValueChanged(sender As Object, e As EventArgs) Handles NumericUpDown10.ValueChanged
        If currentpart = rewindvalue Then
            Refreshflag = True
        End If
        If rewindloopflag = True Then
            'MsgBox("MGB")
            newrefreshflag = True
        End If
        If NumericUpDown10.Value >= totalnumber Then
            NumericUpDown10.Value = totalnumber
        End If
    End Sub
    Private Sub Chart1_GetToolTipText(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataVisualization.Charting.ToolTipEventArgs) Handles Chart1.GetToolTipText

        Chart1.Series(0).ToolTip = "#VALX{0},#VALY{0.0}"
        Chart1.Series(1).ToolTip = "#VALX{0},#VALY{0.0}"
        Chart1.Series(2).ToolTip = "#VALX{0},#VALY{0.0}"
        Chart1.Series(3).ToolTip = "Previous MIN ,#VALX{0},#VALY{0.0}"
        Chart1.Series(4).ToolTip = "Previous MAX ,#VALX{0},#VALY{0.0}"

    End Sub

    Private Sub savesettings()

        If My.Computer.FileSystem.FileExists(CurDir() & "\Config.txt") = True Then
            My.Computer.FileSystem.DeleteFile(CurDir() & "\Config.txt")
        End If
        If My.Computer.FileSystem.FileExists(CurDir() & "\Config.txt") = False Then
            Dim fPath = CurDir() & "\Config.txt"
            Dim afile As New IO.StreamWriter(fPath, True)

            afile.WriteLine(RadioButton1.Checked)
            afile.WriteLine(RadioButton2.Checked)
            afile.WriteLine(NumericUpDown11.Value)
            afile.WriteLine(CheckBox1.Checked)
            afile.WriteLine(CheckBox2.Checked)
            afile.WriteLine(CheckBox3.Checked)
            afile.WriteLine(CheckBox4.Checked)
            afile.WriteLine(CheckBox5.Checked)
            afile.WriteLine(CheckBox6.Checked)
            afile.WriteLine(CheckBox7.Checked)
            afile.WriteLine(CheckBox8.Checked)
            afile.WriteLine(CheckBox9.Checked)
            afile.WriteLine(CheckBox10.Checked)
            afile.WriteLine(CheckBox11.Checked)
            afile.WriteLine(RadioButton3.Checked)
            afile.WriteLine(RadioButton4.Checked)
            afile.WriteLine(RadioButton5.Checked)
            afile.WriteLine(RadioButton6.Checked)
            afile.WriteLine(ComboBox2.Text)
            afile.Close()
        End If

    End Sub

    Private Sub loadsettings()
        If My.Computer.FileSystem.FileExists(CurDir() & "\" & settingsname & ".txt") = True Then
            Dim fPath = CurDir() & "\" & settingsname & ".txt"
            Dim afile As New IO.StreamReader(fPath, True)

            RadioButton1.Checked = afile.ReadLine
            RadioButton2.Checked = afile.ReadLine
            NumericUpDown11.Value = afile.ReadLine
            CheckBox1.Checked = afile.ReadLine
            CheckBox2.Checked = afile.ReadLine
            CheckBox3.Checked = afile.ReadLine
            CheckBox4.Checked = afile.ReadLine
            CheckBox5.Checked = afile.ReadLine
            CheckBox6.Checked = afile.ReadLine
            CheckBox7.Checked = afile.ReadLine
            CheckBox8.Checked = afile.ReadLine
            CheckBox9.Checked = afile.ReadLine
            CheckBox10.Checked = afile.ReadLine
            CheckBox11.Checked = afile.ReadLine
            RadioButton3.Checked = afile.ReadLine
            RadioButton4.Checked = afile.ReadLine
            RadioButton5.Checked = afile.ReadLine
            RadioButton6.Checked = afile.ReadLine
            ComboBox2.Text = afile.ReadLine
            afile.Close()
            'MsgBox("test")'
        End If



    End Sub

    Private Sub comboboxpopulate()
        Dim totalslashes As Integer
        Dim filename1 As String
        ComboBox1.Items.Clear()
        For Each File In IO.Directory.GetFiles(CurDir, "*.txt", IO.SearchOption.AllDirectories)
            'If InStr(File, "PRG^") = 0 And InStr(File, "PRG~") = 0 Then
            If InStr(File, "Config") = False And InStr(File, "Citation") = False Then
                If InStr(File, "currentfolder") = False Then
                    totalslashes = Len(File) - Len(Replace(File, "\", ""))
                    filename1 = File.Split("\"c)(totalslashes)
                    filename1 = Replace(filename1, ".txt", "")
                    ComboBox1.Items.Add(filename1)
                End If
            End If
            'End If
        Next
        'FileClose(1)

    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        'this is the sub routine to save presets
        'MsgBox(TextBox2.Text)
        If TextBox2.Text <> "" Then
            If TextBox2.Text <> "config" Or TextBox2.Text <> "Citation" Then
                If My.Computer.FileSystem.FileExists(CurDir() & "\" & TextBox2.Text & ".txt") = True Then
                    My.Computer.FileSystem.DeleteFile(CurDir() & "\" & TextBox2.Text & ".txt")
                End If
                If My.Computer.FileSystem.FileExists(CurDir() & "\" & TextBox2.Text & ".txt") = False Then
                    Dim fPath = CurDir() & "\" & TextBox2.Text & ".txt"
                    Dim afile As New IO.StreamWriter(fPath, True)

                    afile.WriteLine(RadioButton1.Checked)
                    afile.WriteLine(RadioButton2.Checked)
                    afile.WriteLine(NumericUpDown11.Value)
                    afile.WriteLine(CheckBox1.Checked)
                    afile.WriteLine(CheckBox2.Checked)
                    afile.WriteLine(CheckBox3.Checked)
                    afile.WriteLine(CheckBox4.Checked)
                    afile.WriteLine(CheckBox5.Checked)
                    afile.WriteLine(CheckBox6.Checked)
                    afile.WriteLine(CheckBox7.Checked)
                    afile.WriteLine(CheckBox8.Checked)
                    afile.WriteLine(CheckBox9.Checked)
                    afile.WriteLine(CheckBox10.Checked)
                    afile.WriteLine(CheckBox11.Checked)
                    afile.WriteLine(RadioButton3.Checked)
                    afile.WriteLine(RadioButton4.Checked)
                    afile.WriteLine(RadioButton5.Checked)
                    afile.WriteLine(RadioButton6.Checked)
                    afile.WriteLine(ComboBox2.Text)
                    afile.Close()
                    'MsgBox("test")
                    comboboxpopulate()


                    If UCase(TextBox2.Text) = "MATTBKE63" Then
                        MsgBox("hello friend")
                    End If
                    If UCase(TextBox2.Text) = "CAKE" Then
                        MsgBox("THE CAKE IS A LIE")
                    End If
                    If UCase(TextBox2.Text) = "1223" Then
                        MsgBox("HAPPY ANNIVERSARY!!")
                    End If
                End If
                MsgBox("Settings Saved")
            Else
                MsgBox("Settings can not be named this please pick another name")
            End If
        Else
                MsgBox("Please give presets a name")
        End If

    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click

        If My.Computer.FileSystem.FileExists(CurDir() & "\" & ComboBox1.Text & ".txt") = True Then
            Dim fPath = CurDir() & "\" & ComboBox1.Text & ".txt"
            Dim afile As New IO.StreamReader(fPath, True)

            RadioButton1.Checked = afile.ReadLine
            RadioButton2.Checked = afile.ReadLine
            NumericUpDown11.Value = afile.ReadLine
            CheckBox1.Checked = afile.ReadLine
            CheckBox2.Checked = afile.ReadLine
            CheckBox3.Checked = afile.ReadLine
            CheckBox4.Checked = afile.ReadLine
            CheckBox5.Checked = afile.ReadLine
            CheckBox6.Checked = afile.ReadLine
            CheckBox7.Checked = afile.ReadLine
            CheckBox8.Checked = afile.ReadLine
            CheckBox9.Checked = afile.ReadLine
            CheckBox10.Checked = afile.ReadLine
            CheckBox11.Checked = afile.ReadLine
            RadioButton3.Checked = afile.ReadLine
            RadioButton4.Checked = afile.ReadLine
            RadioButton5.Checked = afile.ReadLine
            RadioButton6.Checked = afile.ReadLine
            ComboBox2.Text = afile.ReadLine


            afile.Close()

        End If
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        'This Subroutine deletes the selected Preset
        My.Computer.FileSystem.DeleteFile(CurDir() & "\" & ComboBox1.Text & ".txt")
        ComboBox1.Text = ""
        comboboxpopulate()
    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        CSVFILENAME = (TextBox1.Text & "\RESULTS.CSV")
        opencsv()
    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        CSVFILENAME = (TextBox1.Text & "\RESULTS_uncalc.CSV")
        opencsv()
    End Sub

    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click
        Dim fd As OpenFileDialog = New OpenFileDialog()

        fd.Title = "Open File Dialog"
        fd.InitialDirectory = TextBox1.Text & "\Old Results"
        fd.Filter = "All Files (*.*)|*.*|CSV FIles (*.CSV)|*.CSV"
        fd.FilterIndex = 2
        fd.RestoreDirectory = True

        If fd.ShowDialog() = DialogResult.OK Then
            CSVFILENAME = fd.FileName
            opencsv()
        End If

    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        rewindloopflag = True

        rewindvalue = rewindvalue - 1
        previousplotgraphvalue = rewindvalue

        If rewindvalue = 1 Then
            Button9.Visible = False
        End If


        If rewindvalue >= 1 Then
            clearchart()

            refreshrewindgraph()

        End If

        Chart1.ChartAreas(0).AxisX.ScaleView.ZoomReset()
        Chart1.ChartAreas(0).AxisY.ScaleView.ZoomReset()



    End Sub

    Private Sub refreshrewindgraph()
        Dim resultscsvread As String() = IO.File.ReadAllLines(RewindCalcfilename)
        Dim resultsuncalccsvread As String() = IO.File.ReadAllLines(rewinduncalcfilename)
        Dim resultsline As String
        Dim resultsuncalcline As String
        Dim commas_ As Integer
        Dim linecounter_ As Integer
        Dim tempnamestim_ As String
        Dim tempname1stim_ As String
        Dim stimval_ As String
        Dim lastpoint As Integer

        Dim numbersmax(4) As Decimal
        Dim numbersmin(4) As Decimal
        Dim matrixnum As Integer
        Dim converttoint As Decimal
        Dim subidrewind As String
        Dim deciblerewind As String
        Dim freqrewind As String
        Dim writevalue As Integer
        Dim filenamerewind As String
        Dim specimenrewind As String
        Dim stimulusrewind As String
        Dim timepointrewind As String
        Dim timeoftest As Decimal
        Dim counter As Integer
        Dim newchartheader As String

        rewindrefreshflag = False
        testflag = False


        clearchart()

        resultsline = resultscsvread(rewindvalue)
        resultsuncalcline = resultsuncalccsvread(rewindvalue)
        writevalue = rewindvalue

        'These are counters
        linecounter_ = 1
        commas_ = 0
        matrixnum = 0

        'this loop filters through the first lines in the CSV till it hits desired column
        While commas_ <= 6
            tempnamestim_ = Microsoft.VisualBasic.Left(resultsuncalcline, linecounter_)
            tempname1stim_ = Microsoft.VisualBasic.Right(tempnamestim_, 1)

            If tempname1stim_ = "," Then
                commas_ = commas_ + 1
            End If

            linecounter_ = linecounter_ + 1
        End While
        lastpoint = linecounter_

        totalnumbercalcrewind()
        'MsgBox(totalnumber)
        binerrorcheck()
        Dim numbers(totalnumber) As Decimal

        'this loop puts all data for analysis into an array
        While commas_ <= totalnumber + 7

            tempnamestim_ = Microsoft.VisualBasic.Left(resultsuncalcline, linecounter_)
            tempname1stim_ = Microsoft.VisualBasic.Right(tempnamestim_, 1)


            If tempname1stim_ = "," Then
                On Error GoTo exitcount
                stimval_ = Microsoft.VisualBasic.Left(resultsuncalcline, linecounter_ - 1)
                stimval_ = Microsoft.VisualBasic.Right(stimval_, linecounter_ - lastpoint)
                converttoint = CDec(stimval_)
                numbers(matrixnum) = converttoint
                matrixnum = matrixnum + 1
                lastpoint = linecounter_ + 1
                commas_ = commas_ + 1

            End If
            linecounter_ = linecounter_ + 1

        End While
exitcount:
        linecounter_ = 1
        commas_ = 0
        matrixnum = 0
        lastpoint = 0

        'this grabs all varaible from calc file
        While commas_ <= 27

            tempnamestim_ = Microsoft.VisualBasic.Left(resultsline, linecounter_)
            tempname1stim_ = Microsoft.VisualBasic.Right(tempnamestim_, 1)


            If tempname1stim_ = "," Then
                stimval_ = Microsoft.VisualBasic.Left(resultsline, linecounter_ - 1)
                stimval_ = Microsoft.VisualBasic.Right(stimval_, linecounter_ - lastpoint)
                'converttoint = CDec(stimval_)
                'numbers2(matrixnum) = converttoint
                Select Case matrixnum
                    Case 1 - 1
                        filenamerewind = stimval_
                    Case 2 - 1
                        subidrewind = stimval_
                    Case 3 - 1
                        specimenrewind = stimval_
                    Case 4 - 1
                        stimulusrewind = stimval_
                    Case 5 - 1
                        timepointrewind = stimval_
                    Case 6 - 1
                        freqrewind = stimval_
                    Case 7 - 1
                        deciblerewind = stimval_
                    Case 18 - 1
                        ' MsgBox(stimval_)
                        timeoftest = CDec(stimval_)
                    Case 19 - 1
                        numbersmax(0) = CDec(stimval_)
                    Case 20 - 1
                        numbersmin(0) = CDec(stimval_)
                    Case 21 - 1
                        numbersmax(1) = CDec(stimval_)
                    Case 22 - 1
                        numbersmin(1) = CDec(stimval_)
                    Case 23 - 1
                        numbersmax(2) = CDec(stimval_)
                    Case 24 - 1
                        numbersmin(2) = CDec(stimval_)
                    Case 25 - 1
                        numbersmax(3) = CDec(stimval_)
                    Case 26 - 1
                        numbersmin(3) = CDec(stimval_)
                    Case 27 - 1
                        numbersmax(4) = CDec(stimval_)
                    Case 28 - 1
                        numbersmin(4) = CDec(stimval_)
                End Select


                matrixnum = matrixnum + 1
                lastpoint = linecounter_ + 1
                commas_ = commas_ + 1
                'MsgBox(stimval_)

            End If
            linecounter_ = linecounter_ + 1

        End While


repopgraph:
        Dim test As Integer
        test = 0
        While test < totalnumber - 1
            Me.Chart1.Series("Amplitude (nV)").Points.AddXY(test, numbers(test))
            test = test + 1
        End While

        Me.Chart1.Series("MAX").Points.AddXY(numbersmax(0), numbers(numbersmax(0)))
        Me.Chart1.Series("MAX").Points.AddXY(numbersmax(1), numbers(numbersmax(1)))
        Me.Chart1.Series("MAX").Points.AddXY(numbersmax(2), numbers(numbersmax(2)))
        Me.Chart1.Series("MAX").Points.AddXY(numbersmax(3), numbers(numbersmax(3)))
        Me.Chart1.Series("MAX").Points.AddXY(numbersmax(4), numbers(numbersmax(4)))

        Me.Chart1.Series("MIN").Points.AddXY(numbersmin(0), numbers(numbersmin(0)))
        Me.Chart1.Series("MIN").Points.AddXY(numbersmin(1), numbers(numbersmin(1)))
        Me.Chart1.Series("MIN").Points.AddXY(numbersmin(2), numbers(numbersmin(2)))
        Me.Chart1.Series("MIN").Points.AddXY(numbersmin(3), numbers(numbersmin(3)))
        Me.Chart1.Series("MIN").Points.AddXY(numbersmin(4), numbers(numbersmin(4)))



        newchartheader = Replace(deciblerewind, "+", "")
        If Len(newchartheader) <> Len(deciblerewind) Then
            newchartheader = newchartheader & " dB atten"
        End If

        Me.Chart1.Titles.Add("Sub ID: " & subidrewind & "         Decibel Level:" & newchartheader & "         Frequency Number: " & freqrewind)
        Me.Chart1.ChartAreas("ChartArea1").AxisX.Title = "Number"
        Me.Chart1.ChartAreas("ChartArea1").AxisY.Title = "Amplitude (nV)"

        Label13.Text = numbers(numbersmax(0))
        Label14.Text = numbers(numbersmax(1))
        Label15.Text = numbers(numbersmax(2))
        Label16.Text = numbers(numbersmax(3))
        Label17.Text = numbers(numbersmax(4))

        NumericUpDown1.Value = numbersmax(0)
        NumericUpDown2.Value = numbersmax(1)
        NumericUpDown3.Value = numbersmax(2)
        NumericUpDown4.Value = numbersmax(3)
        NumericUpDown5.Value = numbersmax(4)

        Label26.Text = numbers(numbersmin(0))
        Label25.Text = numbers(numbersmin(1))
        Label24.Text = numbers(numbersmin(2))
        Label23.Text = numbers(numbersmin(3))
        Label22.Text = numbers(numbersmin(4))

        NumericUpDown10.Value = numbersmin(0)
        NumericUpDown9.Value = numbersmin(1)
        NumericUpDown8.Value = numbersmin(2)
        NumericUpDown7.Value = numbersmin(3)
        NumericUpDown6.Value = numbersmin(4)


        Do Until testflag = True

            Me.Show()
            Application.DoEvents()

            numbersmax(0) = NumericUpDown1.Value
            numbersmax(1) = NumericUpDown2.Value
            numbersmax(2) = NumericUpDown3.Value
            numbersmax(3) = NumericUpDown4.Value
            numbersmax(4) = NumericUpDown5.Value


            numbersmin(0) = NumericUpDown10.Value
            numbersmin(1) = NumericUpDown9.Value
            numbersmin(2) = NumericUpDown8.Value
            numbersmin(3) = NumericUpDown7.Value
            numbersmin(4) = NumericUpDown6.Value

            Dim freq1 As Decimal
            Dim freq2 As Decimal
            Dim freq3 As Decimal
            Dim freq4 As Decimal
            Dim freq5 As Decimal

            freq1 = CStr(numbers(numbersmax(0)) - numbers(numbersmin(0)))
            freq2 = CStr(numbers(numbersmax(1)) - numbers(numbersmin(1)))
            freq3 = CStr(numbers(numbersmax(2)) - numbers(numbersmin(2)))
            freq4 = CStr(numbers(numbersmax(3)) - numbers(numbersmin(3)))
            freq5 = CStr(numbers(numbersmax(4)) - numbers(numbersmin(4)))


            If newrefreshflag = True Then
                Dim lines() As String = IO.File.ReadAllLines(RewindCalcfilename)
                Dim outputrewind As String
                outputrewind = filenamerewind & "," & subidrewind & "," & specimenrewind & "," & stimulusrewind & "," & timepointrewind & "," & freqrewind & "," & deciblerewind
                outputrewind = outputrewind & "," & freq1 & "," & numbersmax(0) * timeoftest / totalnumber & "," & freq2 & "," & numbersmax(1) * timeoftest / totalnumber & "," & freq3 & "," & numbersmax(2) * timeoftest / totalnumber & "," & freq4 & "," & numbersmax(3) * timeoftest / totalnumber & "," & freq5 & "," & numbersmax(4) * timeoftest / totalnumber
                outputrewind = outputrewind & "," & timeoftest & "," & numbersmax(0) & "," & numbersmin(0) & "," & numbersmax(1) & "," & numbersmin(1) & "," & numbersmax(2) & "," & numbersmin(2) & "," & numbersmax(3) & "," & numbersmin(3) & "," & numbersmax(4) & "," & numbersmin(4) & ",,"
                'this does the interpeak latencies
                counter = 0
                Do Until counter = 4
                    If numbersmax(counter) = 0 Then
                        outputrewind = outputrewind & "0,"
                    ElseIf numbersmax(counter + 1) = 0 Then
                        outputrewind = outputrewind & "0,"
                    Else
                        outputrewind = outputrewind & CStr((numbersmax(counter + 1) - numbersmax(counter)) * timeoftest / totalnumber) & ","
                    End If
                    counter = counter + 1
                Loop

                'last 2 interpeak latencies
                If numbersmax(0) = 0 Then
                    outputrewind = outputrewind & "0,"
                ElseIf numbersmax(3) = 0 Then
                    outputrewind = outputrewind & "0,"
                Else
                    outputrewind = outputrewind & CStr((numbersmax(3) - numbersmax(0)) * timeoftest / totalnumber) & ","
                End If

                If numbersmax(0) = 0 Then
                    outputrewind = outputrewind & "0,"
                ElseIf numbersmax(4) = 0 Then
                    outputrewind = outputrewind & "0,"
                Else

                    outputrewind = outputrewind & CStr((numbersmax(4) - numbersmax(0)) * timeoftest / totalnumber) & ","
                End If

                outputrewind = outputrewind & ","

                'This does the Amplitude Ratios
                counter = 1
                Do Until counter = 5
                    If numbersmax(0) = 0 Then
                        outputrewind = outputrewind & "0,"
                    ElseIf numbersmax(counter) = 0 Then
                        outputrewind = outputrewind & "0,"
                    Else
                        outputrewind = outputrewind & CStr(numbers(numbersmax(counter)) - numbers(numbersmin(counter))) / (numbers(numbersmax(0)) - numbers(numbersmin(0))) & ","
                    End If
                    counter = counter + 1
                Loop


                For i As Integer = 0 To UBound(lines)
                    If i = writevalue Then
                        lines(i) = outputrewind
                    End If
                Next
                'MsgBox(lines(writevalue))
                IO.File.WriteAllLines(RewindCalcfilename, lines)

                'MsgBox(rewindloopflag)

                clearchart()

                newrefreshflag = False
                GoTo repopgraph
            End If


            'MsgBox("pause")

            If stopexecution = True Then
                GoTo exitloop
            End If
        Loop


exitloop:



    End Sub


    Private Sub Button13_Click(sender As Object, e As EventArgs) Handles Button13.Click
        stopexecution = True
        'clears chart
        clearchart()

        Label38.Visible = False
        Label39.Visible = False

    End Sub

    Private Sub clearchart()
        Dim test As String

        Me.Chart1.Series("Amplitude (nV)").Points.Clear()
        Me.Chart1.Series("MAX").Points.Clear()
        Me.Chart1.Series("MIN").Points.Clear()
        'Me.Chart1.Series("Previous Results").Points.Clear()
        Me.Chart1.Series("Previous MAX").Points.Clear()
        Me.Chart1.Series("Previous MIN").Points.Clear()
        Me.Chart1.Titles.Clear()
        'Me.Chart1.Series("Previous Results").Legend.Remove()

        test = Me.Chart1.Series.Count
        'MsgBox(test)

        While test <> 5
            Me.Chart1.Series.RemoveAt(test - 1)
            test = test - 1
        End While
        'Me.Chart1.Series.RemoveAt(5)
        'Me.Chart1.Series.RemoveAt(6)

        If rewindvalue <> 1 Then
            Button19.Visible = True
        End If

    End Sub

    Private Sub Button14_Click(sender As Object, e As EventArgs) Handles Button14.Click


        NumericUpDown10.Value = 0
        NumericUpDown9.Value = 0
        NumericUpDown8.Value = 0
        NumericUpDown7.Value = 0
        NumericUpDown6.Value = 0

        NumericUpDown1.Value = 0
        NumericUpDown2.Value = 0
        NumericUpDown3.Value = 0
        NumericUpDown4.Value = 0
        NumericUpDown5.Value = 0




    End Sub

    Private Sub binerrorcheck()
        'This sub routine checks to make sure the first file evaluated has the same number of bins as any other file if it does not it does show text 
        'saying it might have a mismatch

        If fileline = 1 Then
            binvaluefirst = totalnumber
        End If
        If rewindvalue = 1 Then
            binvaluefirst = totalnumber
        End If


        If binvaluefirst <> totalnumber Then

            Label38.Visible = True
            Label39.Visible = True

        End If
        If binvaluefirst = totalnumber Then

            Label38.Visible = False
            Label39.Visible = False

        End If


    End Sub


    Private Sub Button15_Click(sender As Object, e As EventArgs) Handles Button15.Click
        Dim fd As OpenFileDialog = New OpenFileDialog()
        Dim commas As Integer

        fd.Title = "Open File Dialog"
        fd.InitialDirectory = TextBox1.Text & "\Old Results"
        fd.Filter = "All Files (*.*)|*.*|CSV Results Files (*Results_uncalc*.CSV)|*Results_uncalc*.CSV"
        fd.FilterIndex = 2
        fd.RestoreDirectory = True



        'Interrater check sub routine


        If fd.ShowDialog() = DialogResult.OK Then
            rewinduncalcfilename = fd.FileName
            RewindCalcfilename = Replace(fd.FileName, "_uncalc", "")
            If rewinduncalcfilename = RewindCalcfilename Then
                GoTo badfilename
            End If
            Dim commasfile As String() = IO.File.ReadAllLines(rewinduncalcfilename)
            Dim commasline As String


            If commasfile.Length = 1 Then
                GoTo blankfile
            End If
            commasline = commasfile(1)

            commas = Len(commasline) - Len(Replace(commasline, ",", ""))
            'MsgBox(rewinduncalcfilename)
            'MsgBox(commas)
            totalnumber = commas - 8
            rewindvalue = 1
            previousplotgraphvalue = rewindvalue

            rewindloopflag = True
            stopexecution = False
            Chart1.Visible = True
            Button6.Visible = True
            Button14.Visible = True

            'Turn off settings and execute button
            Button1.Visible = False
            'Button17.Visible = False
            Label28.Visible = False
            'Label29.Visible = False
            Label30.Visible = False
            Label37.Visible = False
            Label31.Visible = False
            RadioButton1.Visible = False
            RadioButton2.Visible = False
            NumericUpDown11.Visible = False
            CheckBox1.Visible = False
            CheckBox2.Visible = False
            CheckBox3.Visible = False
            CheckBox4.Visible = False
            CheckBox5.Visible = False
            CheckBox6.Visible = False
            CheckBox7.Visible = False
            CheckBox8.Visible = False
            CheckBox9.Visible = False
            CheckBox10.Visible = False
            TextBox2.Visible = False
            Button4.Visible = False
            Button7.Visible = False
            Button8.Visible = False
            ComboBox1.Visible = False
            Label32.Visible = False
            ComboBox2.Visible = False
            Label40.Visible = False

            'GroupBox1.Visible = False
            CheckBox11.Visible = False
            Button10.Visible = False
            Button11.Visible = False
            Button12.Visible = False
            Button18.Visible = False
            GroupBox2.Visible = False
            Label33.Visible = False
            Button15.Visible = False





            While commasfile.Length - 1 >= rewindvalue

                If rewindvalue = 1 Then
                    Button16.Visible = False
                Else
                    Button16.Visible = True
                End If

                Label35.Text = rewindvalue & "/" & commasfile.Length - 1
                refreshrewindgraph()

                If stopexecution = True Then
                    GoTo exitloop
                End If

            End While

        End If


exitloop:
        Chart1.Visible = False
        Button6.Visible = False
        Button16.Visible = False
        Button14.Visible = False
        Label39.Visible = False
        Label38.Visible = False
        Button19.Visible = False

        'turn on settings and execute button

        Button1.Visible = True
        'Button17.Visible = True
        Label28.Visible = True
        'Label29.Visible = True
        Label30.Visible = True
        Label37.Visible = True
        Label31.Visible = True
        RadioButton1.Visible = True
        RadioButton2.Visible = True
        NumericUpDown11.Visible = True
        CheckBox1.Visible = True
        CheckBox2.Visible = True
        CheckBox3.Visible = True
        CheckBox4.Visible = True
        CheckBox5.Visible = True
        CheckBox6.Visible = True
        CheckBox7.Visible = True
        CheckBox8.Visible = True
        CheckBox9.Visible = True
        CheckBox10.Visible = True
        TextBox2.Visible = True
        Button4.Visible = True
        Button7.Visible = True
        Button8.Visible = True
        ComboBox1.Visible = True
        Label32.Visible = True
        ComboBox2.Visible = True
        Label40.Visible = True
        'GroupBox1.Visible = True
        CheckBox11.Visible = True
        Button10.Visible = True
        Button11.Visible = True
        Button12.Visible = True
        Button18.Visible = True
        GroupBox2.Visible = True
        Label33.Visible = True
        Button15.Visible = True
        Label35.Text = "0/0"

        MsgBox("Data Point Review Complete")
        GoTo endloop
blankfile:
        MsgBox("results file " & rewinduncalcfilename & " is blank please select another file")
        GoTo endloop
badfilename:
        MsgBox("Cant find Calcualtion file for " & rewinduncalcfilename & "Please check file name")
endloop:
    End Sub

    Private Sub Button16_Click(sender As Object, e As EventArgs) Handles Button16.Click
        testflag = True
        rewindvalue = rewindvalue - 1
        previousplotgraphvalue = rewindvalue

        If rewindvalue = 1 Then
            Button19.Visible = False
        End If

        Chart1.ChartAreas(0).AxisX.ScaleView.ZoomReset()
        Chart1.ChartAreas(0).AxisY.ScaleView.ZoomReset()

    End Sub

    Private Sub lauralmfile()
        Dim newfilepath As String

        'This is the sub routine upon M conversion button click

        On Error GoTo LoopError
        'PrgCnt = 0
        FileList2.Clear()

        ErrorText = "enumerating file list"
        FileOpen(1, TextBox1.Text & "\Programs.txt", OpenMode.Output)
        For Each File In IO.Directory.GetFiles(PrgFrom, "*.m", IO.SearchOption.AllDirectories)
            'If InStr(File, "PRG^") = 0 And InStr(File, "PRG~") = 0 Then
            'PrgCnt += 1
            FileList2.Add(File)
            WriteLine(1, File)
            'End If
        Next
        FileClose(1)

        'PrgCnt = 0
        For Each File In FileList2
            CurrentFile = File
            If msgboxflag = False Then
                ErrorText = "Pulling the maximum elapsed time"
                pulltime()
                'PrgCnt = PrgCnt + 1
            End If

            ErrorText = "Pulling the data fields for m to CSV conversion"
            pullfiledata()

        Next
        GoTo endofprg
LoopError:
        MsgBox("Error while " & ErrorText & ":" & Chr(13) & CurrentFile & Chr(13) & Err.Description)
endofprg:
        MsgBox("File Conversion Completed, All m files that could be converted were and moved to Old Results. Files that could not be processed were not moved. Files will now be processed with BioSig RZ")
    End Sub
    Private Sub pulltime()
        Dim linesstimulus_ As String() = IO.File.ReadAllLines(CurrentFile)
        Dim line2_ As String
        Dim lengthoffile As Integer
        Dim counter As Integer
        Dim validdata As Boolean
        Dim validdataline As Integer
        Dim firstnumberpos As Integer
        Dim tempname1 As String
        Dim tempname2 As String
        Dim counter2 As Integer
        Dim endoffirstnumber As Integer


        'This sub routine pulls the data from .m Files

        counter = 0
        lengthoffile = UBound(linesstimulus_)
        validdata = False

        While counter < lengthoffile
            line2_ = linesstimulus_(counter)
            'MsgBox(InStr(line2_, "AverageData"))
            If InStr(line2_, "AverageData") Then
                validdata = True
                validdataline = counter
            End If
            counter = counter + 1
        End While

        'this look goes through first line of data
        If validdata = True Then
            counter = validdataline


            While counter < lengthoffile - 6
                line2_ = linesstimulus_(counter)
                firstnumberpos = InStr(line2_, "[")

                counter2 = firstnumberpos
                While counter2 < Len(line2_)
                    tempname1 = Microsoft.VisualBasic.Left(line2_, counter2)
                    tempname2 = Microsoft.VisualBasic.Right(tempname1, 1)
                    'MsgBox(tempname2)
                    If tempname2 = " " Then
                        endoffirstnumber = counter2
                        tempname1 = (Microsoft.VisualBasic.Left(line2_, endoffirstnumber))
                        firstnumberpos = InStrRev(tempname1, "[")
                        tempname2 = (Microsoft.VisualBasic.Right(tempname1, Len(tempname1) - firstnumberpos))
                        GoTo exitloop
                    End If
                    counter2 = counter2 + 1
                End While
exitloop:
                'MsgBox(tempname2)
                counter = counter + 1
            End While
            MsgBox("Time found in m file " & tempname2 & " updated elapsed time in script please ensure all files have same elapsed time before processing")
            NumericUpDown11.Value = tempname2
            msgboxflag = True
        End If

    End Sub
    Private Sub pullfiledata()
        Dim linesstimulus_ As String() = IO.File.ReadAllLines(CurrentFile)
        Dim line2_ As String
        Dim lengthoffile As Integer
        Dim counter As Integer
        Dim validdata As Boolean
        Dim stimline As String
        Dim stimvalue As Integer
        Dim mfilename As String
        Dim leveldb As String
        Dim leveldbint As Integer
        Dim validdataline As Integer
        Dim datavalues As String
        Dim csvwrite As String
        Dim numberofsamples As Integer
        Dim numberofsamplesstring As String


        'This sub routine pulls the data from .m Files


        counter = 0
        lengthoffile = UBound(linesstimulus_)
        validdata = False

        While counter < lengthoffile
            line2_ = linesstimulus_(counter)
            'MsgBox(InStr(line2_, "AverageData"))
            If InStr(line2_, "AverageData") Then
                validdata = True
                validdataline = counter
            End If
            counter = counter + 1
        End While

        'this look goes through first line of data
        If validdata = True Then
            counter = 0
            While counter < lengthoffile
                line2_ = linesstimulus_(counter)
                'MsgBox(InStr(line2_, "AverageData"))
                If InStr(line2_, ",'Stimuli', {struct('freq_hz', {") Then
                    stimline = line2_
                    stimline = Replace(stimline, ",'Stimuli', {struct('freq_hz', {", "")
                    stimline = Replace(stimline, " } ...", "")
                    stimvalue = CInt(stimline)
                End If
                counter = counter + 1
            End While

            mfilename = CurrentFile
            mfilename = Microsoft.VisualBasic.Right(mfilename, Len(mfilename) - InStrRev(mfilename, "\"))
            mfilename = Microsoft.VisualBasic.Left(mfilename, Len(mfilename) - 2)
            'MsgBox(mfilename)

            If File.Exists(Label2.Text & mfilename & ".CSV") Then
                File.Delete(Label2.Text & mfilename & ".CSV")
            End If


            If stimline = "0" Then
                File.Copy(CurDir() & "\mCSVconvertClicks.CSV", Label2.Text & mfilename & ".CSV")
                stimline = ""
            Else
                File.Copy(CurDir() & "\mCSVconvertTones.CSV", Label2.Text & mfilename & ".CSV")
                stimline = stimline & ","
            End If



            counter = 0
            While counter < lengthoffile
                line2_ = linesstimulus_(counter)
                'MsgBox(InStr(line2_, "AverageData"))
                If InStr(line2_, ",'db_atten', {") Then
                    leveldb = line2_
                    leveldb = Replace(leveldb, ",'db_atten', {", "")
                    leveldb = Replace(leveldb, " } ...", "")
                    leveldbint = CInt(leveldb)
                    leveldbint = leveldbint
                    leveldb = CStr(leveldbint)
                End If
                counter = counter + 1
            End While

            counter = validdataline
            While counter < lengthoffile - 6
                line2_ = linesstimulus_(counter)
                line2_ = Replace(line2_, " ]; ", "")
                line2_ = Microsoft.VisualBasic.Right(line2_, Len(line2_) - InStrRev(line2_, " "))
                counter = counter + 1
                datavalues = datavalues & "," & line2_
            End While


            numberofsamples = Len(datavalues) - Len(Replace(datavalues, ",", ""))
            numberofsamplesstring = CStr(numberofsamples)


            If stimline = "" Then
                csvwrite = ",,,,,,," & mfilename & ",,,,," & stimline & leveldb & "+" & ",,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,," & numberofsamplesstring & ",," & datavalues & ","
            Else
                csvwrite = ",,,,,,," & mfilename & ",,,,," & stimline & leveldb & "+" & ",,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,," & numberofsamplesstring & ",," & datavalues & ","

            End If

            Dim outFile12 As IO.StreamWriter = System.IO.File.AppendText(Label2.Text & mfilename & ".CSV")
            outFile12.WriteLine(csvwrite)
            outFile12.Close()

            File.Delete(Replace(CurrentFile, "Place CSV Here", "Old Results"))
            File.Move(CurrentFile, Replace(CurrentFile, "Place CSV Here", "Old Results"))


        End If

    End Sub

    Private Sub Button18_Click(sender As Object, e As EventArgs) Handles Button18.Click
        Citations.Show()
    End Sub

    Private Sub algorithmumset()
        ComboBox2.Items.Clear()
        For Each File In IO.Directory.GetFiles(CurDir(), "*.exe", IO.SearchOption.AllDirectories)
            If InStr(File, "Auditory Brainstem Response Waveform Analysis.exe") = False Then
                File = Replace(File, CurDir() & "\", "")
                File = Replace(File, ".exe", "")
                ComboBox2.Items.Add(File)
                'MsgBox(File)
            End If
        Next
        ComboBox2.Items.Add("Lauer Lab m File")
    End Sub

    Private Sub Button19_Click(sender As Object, e As EventArgs) Handles Button19.Click
        Dim plot As Integer
        Dim resultscsvread As String() = IO.File.ReadAllLines(RewindCalcfilename)
        Dim resultsuncalccsvread As String() = IO.File.ReadAllLines(rewinduncalcfilename)
        Dim currentlinecalc As String
        Dim currentlineuncalc As String
        Dim counter As Integer
        Dim plotxy As String
        Dim totalcommas As Integer
        Dim counter2 As Integer
        Dim maxplot(4) As Decimal
        Dim minplot(4) As Decimal
        Dim name As String



        previousplotgraphvalue = previousplotgraphvalue - 1

        currentlinecalc = resultscsvread(previousplotgraphvalue)
        currentlineuncalc = resultsuncalccsvread(previousplotgraphvalue)

        If currentlineuncalc.Split(","c)(5) = "Click" Then
            name = currentlineuncalc.Split(","c)(1) & "   " & " Click " & CDec(currentlineuncalc.Split(","c)(6)).ToString("####") & "dB"
        Else
            name = currentlineuncalc.Split(","c)(1) & "   " & CDec(currentlineuncalc.Split(","c)(5)).ToString("####") & "Hz  " & CDec(currentlineuncalc.Split(","c)(6)).ToString("####") & "dB"
        End If

        Me.Chart1.Series.Add(name).ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Spline
        'Series3.Color = System.Drawing.Color.Aqua
        If rewindvalue - previousplotgraphvalue = 1 Then
            Me.Chart1.Series.FindByName(name).Color = System.Drawing.Color.Magenta
        ElseIf rewindvalue - previousplotgraphvalue = 2 Then
            Me.Chart1.Series.FindByName(name).Color = System.Drawing.Color.Coral
        ElseIf rewindvalue - previousplotgraphvalue = 3 Then
            Me.Chart1.Series.FindByName(name).Color = System.Drawing.Color.Purple
        ElseIf rewindvalue - previousplotgraphvalue = 4 Then
            Me.Chart1.Series.FindByName(name).Color = System.Drawing.Color.Green
        ElseIf rewindvalue - previousplotgraphvalue = 5 Then
            Me.Chart1.Series.FindByName(name).Color = System.Drawing.Color.RosyBrown
        End If

        'Series1.ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Spline



        If Microsoft.VisualBasic.Right(currentlineuncalc, 1) <> "," Then
            currentlineuncalc = currentlineuncalc & ","
        End If


        totalcommas = Len(currentlineuncalc) - Len(Replace(currentlineuncalc, ",", ""))
        Dim makeplotarray(totalcommas - 7) As Decimal

        counter = 7
        counter2 = 0
        While counter <> totalcommas
            plotxy = currentlineuncalc.Split(","c)(counter)

            makeplotarray(counter2) = CDec(plotxy)
            counter = counter + 1
            counter2 = counter2 + 1
        End While




        plot = 0
        While plot < totalcommas - 8

            Me.Chart1.Series(name).Points.AddXY(plot, makeplotarray(plot))
            'MsgBox(makeplotarray(plot))
            plot = plot + 1
        End While

        '18
        'Latencies

        maxplot(0) = currentlinecalc.Split(","c)(18)
        minplot(0) = currentlinecalc.Split(","c)(19)
        maxplot(1) = currentlinecalc.Split(","c)(20)
        minplot(1) = currentlinecalc.Split(","c)(21)
        maxplot(2) = currentlinecalc.Split(","c)(22)
        minplot(2) = currentlinecalc.Split(","c)(23)
        maxplot(3) = currentlinecalc.Split(","c)(24)
        minplot(3) = currentlinecalc.Split(","c)(25)
        maxplot(4) = currentlinecalc.Split(","c)(26)
        minplot(4) = currentlinecalc.Split(","c)(27)

        Me.Chart1.Series("Previous MAX").Points.AddXY(maxplot(0), makeplotarray(maxplot(0)))
        Me.Chart1.Series("Previous MAX").Points.AddXY(maxplot(1), makeplotarray(maxplot(1)))
        Me.Chart1.Series("Previous MAX").Points.AddXY(maxplot(2), makeplotarray(maxplot(2)))
        Me.Chart1.Series("Previous MAX").Points.AddXY(maxplot(3), makeplotarray(maxplot(3)))
        Me.Chart1.Series("Previous MAX").Points.AddXY(maxplot(4), makeplotarray(maxplot(4)))

        Me.Chart1.Series("Previous MIN").Points.AddXY(minplot(0), makeplotarray(minplot(0)))
        Me.Chart1.Series("Previous MIN").Points.AddXY(minplot(1), makeplotarray(minplot(1)))
        Me.Chart1.Series("Previous MIN").Points.AddXY(minplot(2), makeplotarray(minplot(2)))
        Me.Chart1.Series("Previous MIN").Points.AddXY(minplot(3), makeplotarray(minplot(3)))
        Me.Chart1.Series("Previous MIN").Points.AddXY(minplot(4), makeplotarray(minplot(4)))





        If previousplotgraphvalue = 1 Then
            Button19.Visible = False
        End If

    End Sub

    Private Sub Chart1_MouseWheel(sender As Object, e As MouseEventArgs) Handles Chart1.MouseWheel
        Try
            Dim ZoomFactor As Double = 0.25   '0 to 1 Represent 0% to 100% Every Wheel Tick.

            Dim Current_xMin As Integer = Chart1.ChartAreas(0).AxisX.ScaleView.ViewMinimum
            Dim Current_xMax As Integer = Chart1.ChartAreas(0).AxisX.ScaleView.ViewMaximum
            Dim Current_yMin As Integer = Chart1.ChartAreas(0).AxisY.ScaleView.ViewMinimum
            Dim Current_yMax As Integer = Chart1.ChartAreas(0).AxisY.ScaleView.ViewMaximum

            Dim PointerPosOnChart_xAxis As Integer = Chart1.ChartAreas(0).AxisX.PixelPositionToValue(e.Location.X)
            Dim PointerPosOnChart_yAxis As Integer = Chart1.ChartAreas(0).AxisY.PixelPositionToValue(e.Location.Y)

            Dim New_xMin As Integer
            Dim New_xMax As Integer
            Dim New_yMin As Integer
            Dim New_yMax As Integer

            If Current_xMin <= PointerPosOnChart_xAxis And PointerPosOnChart_xAxis <= Current_xMax And
                Current_yMin <= PointerPosOnChart_yAxis And PointerPosOnChart_yAxis <= Current_yMax Then

                If e.Delta > 0 Then 'Zoom in While Moving The Mouse Wheel Forward (Up)

                    New_xMin = Current_xMin + ((PointerPosOnChart_xAxis - Current_xMin) * ZoomFactor)
                    New_xMax = Current_xMax - ((Current_xMax - PointerPosOnChart_xAxis) * ZoomFactor)
                    New_yMin = Current_yMin + ((PointerPosOnChart_yAxis - Current_yMin) * ZoomFactor)
                    New_yMax = Current_yMax - ((Current_yMax - PointerPosOnChart_yAxis) * ZoomFactor)

                ElseIf e.Delta < 0 Then 'Zoom out While Moving The Mouse Wheel Backward (Down)

                    New_xMin = Current_xMin - ((PointerPosOnChart_xAxis - Current_xMin) * ZoomFactor)
                    New_xMax = Current_xMax + ((Current_xMax - PointerPosOnChart_xAxis) * ZoomFactor)
                    New_yMin = Current_yMin - ((PointerPosOnChart_yAxis - Current_yMin) * ZoomFactor)
                    New_yMax = Current_yMax + ((Current_yMax - PointerPosOnChart_yAxis) * ZoomFactor)

                End If

                Chart1.ChartAreas(0).AxisX.ScaleView.Zoom(New_xMin, New_xMax)
                Chart1.ChartAreas(0).AxisY.ScaleView.Zoom(New_yMin, New_yMax)

            End If


        Catch ex As System.Exception
            MsgBox(ex.Message)
        End Try
    End Sub


    Private Sub Chart1_PAN(sender As Object, e As MouseEventArgs) Handles Chart1.MouseDown


        PointerPosOnChart_xAxis_pan = Chart1.ChartAreas(0).AxisX.PixelPositionToValue(e.Location.X)
        PointerPosOnChart_yAxis_pan = Chart1.ChartAreas(0).AxisY.PixelPositionToValue(e.Location.Y)


    End Sub

    Private Sub Chart1_PAN2(sender As Object, e As MouseEventArgs) Handles Chart1.MouseUp
        Dim xdelta As Integer
        Dim ydelta As Integer

        On Error GoTo skippan
        Dim PointerPosOnChart_xAxis_pan2 As Integer = Chart1.ChartAreas(0).AxisX.PixelPositionToValue(e.Location.X)
        Dim PointerPosOnChart_yAxis_pan2 As Integer = Chart1.ChartAreas(0).AxisY.PixelPositionToValue(e.Location.Y)

        Dim Current_xMin As Integer = Chart1.ChartAreas(0).AxisX.ScaleView.ViewMinimum
        Dim Current_xMax As Integer = Chart1.ChartAreas(0).AxisX.ScaleView.ViewMaximum
        Dim Current_yMin As Integer = Chart1.ChartAreas(0).AxisY.ScaleView.ViewMinimum
        Dim Current_yMax As Integer = Chart1.ChartAreas(0).AxisY.ScaleView.ViewMaximum


        xdelta = PointerPosOnChart_xAxis_pan - PointerPosOnChart_xAxis_pan2
        ydelta = PointerPosOnChart_yAxis_pan - PointerPosOnChart_yAxis_pan2

        If xdelta <> 0 And ydelta <> 0 Then
            Chart1.ChartAreas(0).AxisX.ScaleView.Zoom(Current_xMin + xdelta, Current_xMax + xdelta)
            Chart1.ChartAreas(0).AxisY.ScaleView.Zoom(Current_yMin + ydelta, Current_yMax + ydelta)
        End If

Skippan:

    End Sub


    Private Sub Chart1_resetzoom(sender As Object, e As MouseEventArgs) Handles Chart1.DoubleClick


        Chart1.ChartAreas(0).AxisX.ScaleView.ZoomReset()
        Chart1.ChartAreas(0).AxisY.ScaleView.ZoomReset()

        PointerPosOnChart_xAxis_pan = Chart1.ChartAreas(0).AxisX.PixelPositionToValue(e.Location.X)
        PointerPosOnChart_yAxis_pan = Chart1.ChartAreas(0).AxisY.PixelPositionToValue(e.Location.Y)


    End Sub

    Private Sub Button20_Click(sender As Object, e As EventArgs) Handles Button20.Click
        Dim test As Integer
        test = Me.Chart1.Series.Count
        'MsgBox(test)

        While test <> 5
            Me.Chart1.Series.RemoveAt(test - 1)
            test = test - 1
        End While

        previousplotgraphvalue = rewindvalue

        Me.Chart1.Series("Previous MAX").Points.Clear()
        Me.Chart1.Series("Previous MIN").Points.Clear()

        If previousplotgraphvalue <> 1 Then
            Button19.Visible = True
        End If

        Chart1.ChartAreas(0).AxisX.ScaleView.ZoomReset()
        Chart1.ChartAreas(0).AxisY.ScaleView.ZoomReset()

    End Sub


End Class
