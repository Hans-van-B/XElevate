Module CreateLaunchfile
    '---- Launch File Init, --------------------------------------------------
    ' only initializes the file And write the header
    ' The actual commands are still written in main
    Sub LaunchfileInit()

        xtrace("", 2)
        xtrace("LaunchfileInit", 2)
        xtrace(" - Exec Dir   = " & ExecDir, 2)

        '---- Check if the exe dir exists
        If (System.IO.Directory.Exists(ExecDir)) Then
            xtrace(" - ExecDir Exists", 2)
        Else
            xtrace(" - Create ExecDir", 2)
            System.IO.Directory.CreateDirectory(ExecDir)
        End If

        If IndexLaunchFile Then
            LaunchfileInitIndexed()
        Else
            LaunchfileInitNotIndexed()
        End If


    End Sub

    Sub LaunchfileInitIndexed()
        xtrace(" - Create indexed launch file", 2)
        '---- Create non-existing launch file
        ' This should support simultainious execution
        Dim NrTx As String = ""
        Dim FileNr As Integer

        xtrace(" - Seek non-existing launch file", 2)

        Dim TryPath As String

        For Nr As Integer = 1 To MaxFileNr
            NrTx = Nr.ToString
            TryPath = ExecDir & "\Launch_" & NrTx & ".bat"
            xtrace("    o Try " & TryPath, 3)
            If Not System.IO.File.Exists(TryPath) Then
                FileNr = Nr
                Launchfile = TryPath
                Exit For
            End If
        Next

        ' Check for max file number
        If Launchfile = "" Then
            xtrace(" - End of list reached, restart", 2)
            FileNr = 1
            NrTx = FileNr.ToString
            Launchfile = ExecDir & "\Launch_" & NrTx & ".bat"
        End If

        ' Write the file header
        xtrace(" - Write to Launchfile", 2)
        Try
            My.Computer.FileSystem.WriteAllText(Launchfile, "" & vbNewLine, False)
        Catch ex As Exception
            xtrace("Error 001: Failed to create the launch file, " & ex.Message, 1)
            ExitProgram = True
            Exit Sub
        End Try

        WriteLanchFile("@cls")
        If SetDebugInElevatedEnvironment Then
            WriteLanchFile("@echo on")
        Else
            WriteLanchFile("@if not ""%DEBUG%""==""TRUE"" echo off")
        End If
        WriteLanchFile("@echo XElevated process")

        '---- Move the log file
        xtrace(" - Move the log file", 2)
        Dim Logfile_2 As String
        Logfile_2 = ExecDir & "\Launch_" & NrTx & ".log"

        FileCopy(LogFile, Logfile_2)

        ' Delete the log file in %TEMP%
        My.Computer.FileSystem.DeleteFile(LogFile)

        ' From now on, log to ExecDir
        LogFile = Logfile_2
        xtrace(" - Log file moved", 2)
        ' No dash because it is also written to the console
        xtrace(" Logging to " & LogFile, tl)
        WriteLanchFile("set XElevate_Log=" & LogFile)

        '---- Delete next logfile in row
        ' This create a hole in the row of logfiles where the next one will be created.
        ' In that way the logfile will cicle through 1..MaxFileNr
        Dim Nextf As String
        NrTx = FileNr + 1.ToString

        Nextf = ExecDir & "\Launch_" & NrTx & ".bat"
        xtrace(" - Check next bat " & Nextf, 2)
        If System.IO.File.Exists(Nextf) Then
            xtrace(" - Delete " & Nextf, 2)
            My.Computer.FileSystem.DeleteFile(Nextf)
        End If

        Nextf = ExecDir & "\Launch_" & NrTx & ".log"
        xtrace(" - Check next log " & Nextf, 2)
        If System.IO.File.Exists(Nextf) Then
            xtrace(" - Delete " & Nextf, 2)
            My.Computer.FileSystem.DeleteFile(Nextf)
        End If
    End Sub

    Sub LaunchfileInitNotIndexed()
        xtrace(" - Create non-indexed launch file", 2)

        Dim Logfile_2 As String

        Launchfile = ExecDir & "\Launch_1.bat"
        Logfile_2 = ExecDir & "\Launch_1.log"

        ' Write the file header
        xtrace(" - Write to Launchfile", 2)
        Try
            My.Computer.FileSystem.WriteAllText(Launchfile, "" & vbNewLine, False)
        Catch ex As Exception
            xtrace("Error 002: Failed to create the launch file, " & ex.Message, 1)
            'ExitProgram = True
            'Exit Sub

        End Try
        WriteLanchFile("@cls")
        If SetDebugInElevatedEnvironment Then
            WriteLanchFile("@echo on")
        Else
            WriteLanchFile("@if not ""%DEBUG%""==""TRUE"" echo off")
        End If
        WriteLanchFile("@echo XElevated process")

        '---- Move the log file
        xtrace(" - Move the log file", 2)
        FileCopy(LogFile, Logfile_2)

        ' Delete the log file in %TEMP%
        My.Computer.FileSystem.DeleteFile(LogFile)

        ' From now on, log to ExecDir
        LogFile = Logfile_2
        xtrace(" - Log file moved", 2)
        ' No dash because it is also written to the console
        xtrace(" Logging to " & LogFile, tl)
    End Sub

    Sub WriteLanchFile(Line As String)
        My.Computer.FileSystem.WriteAllText(Launchfile, Line & vbNewLine, True)
    End Sub

End Module
