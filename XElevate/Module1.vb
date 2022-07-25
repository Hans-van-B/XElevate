' 2021-09-23 Quotes added to exe string
Module Module1

    Sub Main()
        Dim ExeString As String = ""

        '==== Initialize ======================================================
        Log_Init()
        If ExitProgram Then Exit Sub

        '==== Read Command-line parameters ====================================

        xtrace("", 2)
        ReadArgV2()
        If ExitProgram Then Exit Sub

        '---- Separate write progname to con ----------------------------------
        ' after reading the param so that BeBrief will work

        If (Not BeQuiet) Then
            Console.Write(" " & AppName & " - " & AppVer & vbNewLine)
        End If

        '---- Check if the executable exists ----------------------------------

        If ToExecute = "" Then
            NoExeSpecified()
            'WriteErr("The executable was not specified")
            'Exit Sub
        Else

            If (System.IO.File.Exists(ToExecute)) Then
                xtrace("Executable exists", 2)
            Else
                xtrace("Executable does not exist, but it may be in the path", 2)
            End If
        End If

        '---- Check it the process is running in an already XElevated environment ----
        ' This will only work with log file indexing on!

        If Environment.GetEnvironmentVariable(XElevateVar) = "TRUE" Then
            xtrace("The process is already XElevated")
            LaunchfileInitIndexed()

            xtrace("Build ExeString", 2)
            ExeString = "CALL "
            ExeString = ExeString & """" & ToExecute & """ " & TeEx_Arg     ' 2021-09-23
            ExeString = ExeString & " >>" & LogFile & " 2>>&1"
            WriteLanchFile(" ")
            WriteLanchFile(ExeString)

            Execute_NotElevated(Launchfile, ToExecute)
            Exit Sub
        End If

        '---- List the environment --------------------------------------------

        LaunchfileInit()

        If TransferEnvironment Then
            xtrace("Transfer the environment", 2)
            WriteLanchFile(" ")

            Dim Name As String
            Dim Val As String
            Dim VLen As Integer = 0
            Dim EnvList(EnvBufferSize, 2) As String

            For Each EnvVar As DictionaryEntry In Environment.GetEnvironmentVariables()
                Name = EnvVar.Key
                Val = EnvVar.Value

                If EnvironmentExceptions Then
                    If Name = "PROCESSOR_ARCHITECTURE" Then Continue For
                    If Name = "USERNAME" Then Continue For
                    If Name = "PROCESSOR_ARCHITEW6432" Then Continue For
                End If

                VLen = VLen + 1
                If VLen = EnvBufferSize Then
                    WriteErr("Environment Buffer Too Small!")
                    Exit For
                End If

                EnvList(VLen, 1) = Name
                EnvList(VLen, 2) = Val
                xtrace(" " & VLen.ToString & " - " & Name & "=" & Val, 2)
            Next

            'Array.Sort(EnvList)

            For Nr As Integer = 1 To VLen
                Name = EnvList(Nr, 1)
                Val = EnvList(Nr, 2)
                WriteLanchFile("SET " & Name & "=" & Val)
            Next

            xtrace("Set current directory")
            WriteLanchFile("CD /D """ & Environment.CurrentDirectory & """")
        Else
            xtrace("Transfer The Environment Is Off", 2)
        End If

        '---- Extra

        WriteLanchFile(" ")
        'WriteLanchFile(" echo on")
        WriteLanchFile("SET " & XElevateVar & "=TRUE")
        WriteLanchFile("SET " & XElevateVar & "_LOG=" & LogFile)
        If ForceExit Then WriteLanchFile("SET EXITLEVEL=0")
        If SetDebugInElevatedEnvironment Then WriteLanchFile("SET DEBUG=TRUE")

        If FixNetDrv Then
            CheckDriveInElevatedProcess()
        Else
            xtrace("Check Drive In Elevated Process Is Off (--fmd)", 2)
        End If

        '---- Build execute string
        xtrace(" ", 2)
        xtrace("Compile ExeString", 2)
        If UseCmd Then
            ExeString = "%COMSPEC% /c "
        Else
            ExeString = "CALL """
        End If
        xtrace(" ExeStr = " & ExeString, 3)

        ExeString = ExeString & ToExecute & """ " & TeEx_Arg
        xtrace(" - ExeStr = " & ExeString, 3)

        If RedirectOutput Then
            ExeString = ExeString & " >>" & LogFile & " 2>>&1"
            xtrace(" - ExeStr = " & ExeString, 2)
        End If

        WriteLanchFile(" ")
        WriteLanchFile(ExeString)

        '---- Write Wait

        If UseWait > 0 And Not UsePause Then
            xtrace("Add Timeout", 2)
            WriteLanchFile("timeout /t " & UseWait.ToString)
        End If

        '---- Write Pause

        If UsePause Then
            xtrace("Add Pause", 2)
            WriteLanchFile("Pause")
        End If

        '---- Write Forced Exit

        If ForceExit Then
            xtrace("Add Exit", 2)
            WriteLanchFile("EXIT %EXITLEVEL%")
        End If

        WriteLanchFile(" ")
        WriteLanchFile("REM ____ Exit XElevate Launch File ____")

        '---- Start Elevated --------------------------------------------------

        Execute_Elevated(Launchfile, ExeString)

    End Sub

End Module
