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
        xtrace(" - Version    = " & AppVer, 2)
        xtrace(" - AppDate    = " & AppDate, 2)
        xtrace(" - Date       = " & DateTime.Now.ToString("yyyy-MM-dd"), 2)
        xtrace(" - Time       = " & DateAndTime.TimeString, 2)
        xtrace(" - StartupDir = " & StartupDir, 2)
        xtrace(" - Temp       = " & UTmp, 2)
        xtrace(" - Log level to con     = " & CTrace.ToString, 2)
        xtrace(" - Log level to logfile = " & LTrace.ToString, 2)
    End Sub

    '---- Xtrace --------------------------------------------------------------
    Sub xtrace_root(Msg As String, TV As Integer)
        xtrace_root(Msg, TV, "")
    End Sub
    Sub xtrace_root(Msg As String, TV As Integer, Extra As String)
        Dim Nr As Int16
        Dim Tab As String = ""

        ' If subroutines are Not logged then tabbing is also disabeled
        If LTrace >= G_TL_Sub Then
            Tab = "|"
            For Nr = 1 To SubLevel
                Tab = Tab + "  |"
            Next
        End If

        If TV <= CTrace Then
            Console.Write(Msg & vbCrLf)
        End If

        If TV <= LTrace Then
            Try ' Catch log conflict
                My.Computer.FileSystem.WriteAllText(LogFile, Tab + Msg & vbCrLf, True)
            Catch ex1 As Exception
                wait(1)
                Try
                    My.Computer.FileSystem.WriteAllText(LogFile, Tab + Msg & vbCrLf, True)
                Catch ex2 As Exception

                End Try
            End Try

        End If
    End Sub

    Sub xtrace(Msg As String)
        xtrace_root(" " & Msg, 2)
    End Sub

    Sub xtrace(Msg As String, TV As Integer)
        xtrace_root(" " & Msg, TV)
    End Sub

    '---- xtrace_i ----
    Sub xtrace_i(Msg As String)
        xtrace(" * " & Msg)
    End Sub

    Sub xtrace_i(Msg As String, TV As Integer)
        xtrace(" * " & Msg, TV)
    End Sub

    '--- xtrace Sub ----
    Sub xtrace_subs(Msg As String)
        xtrace_root("->" & Msg & " (" & (SubLevel + 1).ToString & ")", G_TL_Sub)
        SubLevel = SubLevel + 1
    End Sub

    Sub xtrace_sube(Msg As String)
        SubLevel = SubLevel - 1
        xtrace_root("<-" & Msg & " (" & (SubLevel + 1).ToString & ")", G_TL_Sub)
    End Sub

    Sub xtrace_sube()
        xtrace_sube("<NoName>")
    End Sub

    '---- Write Error ---------------------------------------------------------
    Sub xtrace_Err(Msg() As String)
        Dim Line As String
        Dim Nr As Integer = 0
        xtrace("", 0)
        For Each Line In Msg
            Nr += 1
            If Nr = 1 Then
                xtrace("Error: " & Line, 0)
            Else
                xtrace("       " & Line, 0)
            End If

        Next
    End Sub
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
