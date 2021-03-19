Module Log

    '---- Init ----------------------------------------------------------------

    Dim Err1 As String = ""
    Dim Err2 As String = ""
    Sub Log_Init()
        UTmp = Environment.GetEnvironmentVariable("TEMP")
        If (Environment.GetEnvironmentVariable("DEBUG") = "TRUE") Then Debug = True

        LogFile = UTmp & "\XElevate.log"
        ' Create new file
        Try
            My.Computer.FileSystem.WriteAllText(LogFile, vbNewLine, False)
        Catch ex As Exception
            Err1 = ex.Message
            Console.WriteLine("Error 003: Failed to create the log file. " & Err1)
            LogInitRecover()
        End Try

        xtrace(AppName, 2)
        xtrace(" - Version    = " & AppVersion, 2)
        xtrace(" - AppDate    = " & AppDate, 2)
        xtrace(" - Date       = " & DateTime.Now.ToString("yyyy-MM-dd"), 2)
        xtrace(" - Time       = " & DateAndTime.TimeString, 2)
        xtrace(" - StartupDir = " & StartupDir, 2)
        xtrace(" - Temp       = " & UTmp, 2)
        xtrace(" - Log level to con     = " & CTrace.ToString, 2)
        xtrace(" - Log level to logfile = " & LTrace.ToString, 2)
    End Sub

    '---- Xtrace --------------------------------------------------------------
    Sub xtrace(Msg As String, TV As Integer)
        If TV <= CTrace Then
            Console.Write(Msg & vbNewLine)
        End If

        If (TV <= LTrace) And LogFile <> "x" Then
            My.Computer.FileSystem.WriteAllText(LogFile, Msg & vbNewLine, True)
        End If
    End Sub

    Sub xtrace(Msg As String)
        If (2 <= LTrace) Then
            My.Computer.FileSystem.WriteAllText(LogFile, Msg & vbNewLine, True)
        End If
    End Sub

    '---- Write Error ---------------------------------------------------------

    Sub WriteErr(Msg As String)
        xtrace("", 0)
        xtrace("Error: " & Msg, 0)
    End Sub

    '---- Write Warning -------------------------------------------------------
    Sub WriteWarn(Msg As String)
        xtrace("Warning: " & Msg, 1)
    End Sub

    '---- Warning: No exe specified -------------------------------------------

    Sub NoExeSpecified()
        WriteWarn("The executable was not specified")
        If Not (System.IO.Directory.Exists(ExecDir)) Then
            xtrace(" - Create ExecDir", 2)
            System.IO.Directory.CreateDirectory(ExecDir)
        End If

        ToExecute = ExecDir & "\Dummy.bat"
        xtrace("Create " & ToExecute, 0)

        My.Computer.FileSystem.WriteAllText(ToExecute, "" & vbNewLine, False)
        WriteExeFile("@cls")
        WriteExeFile("@echo off")
        WriteExeFile("echo Dummy")
        'WriteExeFile("pause")
    End Sub

    Sub WriteExeFile(Line As String)
        My.Computer.FileSystem.WriteAllText(ToExecute, Line & vbNewLine, True)
    End Sub

    '==== Log Init Recover ====================================================

    Sub LogInitRecover()
        Dim Nr = 0
        Dim LogSucceeded As Boolean = False


        While Not LogSucceeded
            Nr = Nr + 1
            LogFile = UTmp & "\XElevate_" & Nr.ToString & ".log"
            Console.WriteLine("Try " & LogFile)

            Try
                My.Computer.FileSystem.WriteAllText(LogFile, vbNewLine, False)
                My.Computer.FileSystem.WriteAllText(LogFile, "Logfile recover from", True)
                My.Computer.FileSystem.WriteAllText(LogFile, "  " & Err1, True)
                If Err2 <> "" Then My.Computer.FileSystem.WriteAllText(LogFile, "  " & Err2, True)
                LogSucceeded = True

            Catch ex As Exception
                Err2 = ex.Message
                Console.WriteLine("Error 004: Failed to create log file. " & Err2)
                If Nr > 100 Then
                    ExitProgram = True
                    Console.WriteLine("Fatal error")
                    Exit Sub
                End If
            End Try
        End While

    End Sub

End Module
