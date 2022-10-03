Module ReadCmdArguments
    Public Debug_Exe As Boolean = False

    '---- Read Command line arguments V2 ----
    ' Matches Help V2

    Sub ReadArgV2()
        xtrace_subs("Read Command-line parameters V2")

        Dim Sw As String
        Dim SwitchRow As String
        Dim l As String
        Dim P1 As Integer
        Dim SVal As String = ""
        Dim Name As String
        Dim ValS As String

        ' For each argument
        For Each Arg In My.Application.CommandLineArgs
            xtrace(" - Arg: " & Arg, 2)

            ' Win Style Arguments
            l = Left(Arg, 1)
            If l = "/" Then
                Sw = Mid(Arg, 2)
                xtrace(" - Sw : " & Sw, 2)

                If (Sw = "h") Or (Sw = "?") Then
                    ShowHelp()
                    ExitProgram = True
                End If

                Continue For
            End If

            ' Check for double-dash arguments
            If Left(Arg, 2) = "--" Then
                Sw = Mid(Arg, 3)
                xtrace_i("DDA:" & Sw)

                '---- Double-dash Assignment
                P1 = InStr(Sw, "=")
                If P1 > 0 Then
                    Name = Left(Sw, P1 - 1)
                    ValS = Mid(Sw, P1 + 1)
                    xtrace_i("Name = " & Name)
                    xtrace_i("Val  = " & ValS)

                    ' Also see the implicit read at the end
                    If Name = "ep" Then
                        TeEx_Arg = ValS
                        xtrace_i("Set ExeParam = " & TeEx_Arg)
                    End If

                    Continue For
                End If

                ' Check for assignment
                P1 = InStr(Sw, "=")
                If P1 > 0 Then
                    SVal = Mid(Sw, P1 + 1)
                    Sw = Left(Sw, P1 - 1)
                    xtrace(" - Sw=" & Sw & ", Val=" & SVal)
                End If

                If Sw = "cv" Then
                    xtrace(" Console verbose", 0)
                    CTrace = CTrace + 1
                End If

                If Sw = "de" Then
                    xtrace(" Debug Exe = True (Exe output goes to the log file)", 2)
                    Debug_Exe = True
                End If

                If Sw = "fmd" Then
                    xtrace(" Fix Mapped Network Drive", 2)
                    FixNetDrv = True
                End If

                If Sw = "w" Then
                    UseWait = Val(SVal)
                    xtrace(" Set Wait Time = " & SVal, 2)
                End If


                If Sw = "help" Then
                    ShowHelp()
                    ExitProgram = True
                End If

                If Sw = "xhelp" Then
                    ShowXHelp()
                    ExitProgram = True
                End If

                Continue For
            End If

            ' Check for single dash arguments
            l = Left(Arg, 1)
            If l = "-" Then
                SwitchRow = Mid(Arg, 2)
                While SwitchRow <> ""
                    xtrace(" - SwitchRow: " & SwitchRow, 3)
                    Sw = Left(SwitchRow, 1)
                    xtrace(" - Sw : " & Sw, 2)

                    If Sw = "h" Then
                        ShowHelp()
                        ExitProgram = True
                    End If

                    If Sw = "b" Then
                        If Debug Then
                            xtrace(" Be brief ignored", 2)
                        Else
                            xtrace(" Be brief", 2)
                            BeBrief = True
                            tl = 2
                        End If
                    End If

                    If Sw = "c" Then
                        xtrace(" Transfer the environment complete", 2)
                        TransferEnvironment = True
                        EnvironmentExceptions = False
                    End If

                    If Sw = "d" Then
                        xtrace(" Set Debug In Elevated Environment", 2)
                        SetDebugInElevatedEnvironment = True
                    End If

                    If Sw = "e" Then
                        xtrace(" Transfer the environment", 2)
                        TransferEnvironment = True
                    End If

                    If Sw = "i" Then
                        xtrace(" Disable Launch File Indexing", 2)
                        IndexLaunchFile = False
                    End If

                    If Sw = "k" Then
                        xtrace(" Use COMSPEC", 2)
                        UseCmd = True
                    End If

                    If Sw = "m" Then
                        xtrace(" Run Minimized", 2)
                        RunMinimized = True
                    End If

                    If Sw = "p" Then
                        xtrace(" Use Pause", 2)
                        UsePause = True
                    End If

                    If Sw = "r" Then
                        xtrace(" Redirect output", 2)
                        RedirectOutput = True
                    End If

                    If Sw = "v" Then
                        LTrace = LTrace + 1
                        xtrace(" Verbose " & LTrace.ToString, 2)
                        Console.WriteLine(" Log file = " & LogFile)
                    End If

                    If Sw = "q" Then
                        If Debug Then
                            xtrace(" Be quiet ignored", 2)
                        Else
                            xtrace(" Be quiet", 2)
                            BeQuiet = True
                            tl = 2
                        End If
                    End If

                    If Sw = "w" Then
                        xtrace(" Wait For Lanched Process", 2)
                        WaitForLanchedProcess = True
                    End If

                    If Sw = "x" Then
                        xtrace(" Force eXit", 2)
                        ForceExit = True
                    End If

                    SwitchRow = Mid(SwitchRow, 2)
                End While
                Continue For
            End If ' End SwitchRow

            '---- Executable and Exec-paramaters
            If ToExecute = "" Then
                ToExecute = Arg
                xtrace(" Executable = " & ToExecute, 2)
            Else
                If TeEx_Arg = "" Then
                    TeEx_Arg = Arg
                    xtrace(" Executable Arg = " & TeEx_Arg, 2)
                Else
                    TeEx_Arg = TeEx_Arg & " " & Arg
                    xtrace(" Executable Arg = " & TeEx_Arg, 2)
                End If

            End If
        Next

        xtrace(" ", 2)

        '---- Override Defaults by environment variables
        ' Sometimes, in rare cases, the subst command does not exit and causes xelevate to hang.
        If Environment.GetEnvironmentVariable("TA_DISABLE_FMD") = "TRUE" Then
            FixNetDrv = False
            xtrace_i("TA_DISABLE_FMD = TRUE!")
        End If

        xtrace_sube("Read Command-line parameters V2")
    End Sub


End Module
