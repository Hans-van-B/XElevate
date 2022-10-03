Imports System.IO
Module Execute
    Sub Execute_Elevated(Launchfile As String, ExeString As String)
        xtrace_subs("Execute_Elevated")

        xtrace_i("Set current Directory to make sure it is valid")
        Directory.SetCurrentDirectory(ExecDir)

        If Debug_Exe Then
            Execute_Elevated_De(Launchfile, ExeString)
        Else
            Execute_Elevated_Std(Launchfile, ExeString)
        End If

        xtrace_sube("Execute_Elevated")
    End Sub

    Sub Execute_Elevated_De(Launchfile As String, ExeString As String)
        xtrace_subs("Execute_Elevated_De")

        xtrace(" 0 Proces Start Elevated", 2)
        Dim Proc As New Process()
        Dim Output As String = ""
        With Proc.StartInfo
            .FileName = Launchfile
            .UseShellExecute = False
            .RedirectStandardError = True
            .RedirectStandardOutput = True
            .CreateNoWindow = True
            .Verb = "runas"
        End With

        Try
            If RunMinimized Then
                xtrace(" 1 RunMinimized")
                Proc.StartInfo.WindowStyle = ProcessWindowStyle.Minimized
            End If

            xtrace(" 2 Start " & Proc.StartInfo.FileName)
            Proc.Start()
            Output = Proc.StandardOutput.ReadToEnd()

            xtrace(" 3 Read output")
            Dim OutputLines() As String = Split(Output, vbNewLine)

            For Each Line As String In OutputLines
                xtrace(" | " & Line)
            Next

            xtrace(" 4 Wait for exit")
            If WaitForLanchedProcess Then Proc.WaitForExit()

        Catch ex As Exception
            WriteErr("Warning: Failed to start the elevated process!")
            xtrace(ex.Message.ToString, 1)
            Console.Write("Enter to continiue.")
            Console.ReadLine()
        End Try


        xtrace_sube("Execute_Elevated_De")
    End Sub

    Sub Execute_Elevated_Std(Launchfile As String, ExeString As String)
        xtrace_subs("Execute_Elevated_Std")

        Dim Proc As New Process()
        Dim ProcStartInfo As New ProcessStartInfo(Launchfile)

        xtrace("Proces Start Elevated And Wait", 2)

        Try
            If WaitForLanchedProcess Then xtrace(" Wait For Lanched Process to finish", 2)
            If RedirectOutput Then
                xtrace("", 2)
                xtrace("==== Output of Launced process below ======================================", 2)
                xtrace(ExeString)
            End If

            Proc.StartInfo = ProcStartInfo
            Proc.StartInfo.Verb = "runas"
            If RunMinimized Then Proc.StartInfo.WindowStyle = ProcessWindowStyle.Minimized
            Proc.Start()
            If WaitForLanchedProcess Then Proc.WaitForExit()

            'If RedirectOutput Then wait(2) ' Prevent output conflict

        Catch ex As Exception
            WriteErr("Warning: Failed to start the elevated process!")
            xtrace(ex.Message.ToString, 1)
            Console.Write("Enter to continiue.")
            Console.ReadLine()
        End Try

        xtrace_sube("Execute_Elevated_Std")
    End Sub

    '---- Process is already elevated, do not try to elevate again

    Sub Execute_NotElevated(Launchfile As String, ExeString As String)
        xtrace_subs("Execute_NotElevated")

        Dim Proc As New Process()
        Dim ProcStartInfo As New ProcessStartInfo(Launchfile)

        xtrace("Proces Start And Wait", 2)
        xtrace("ExeString = " & ExeString, 2)

        Try
            If WaitForLanchedProcess Then xtrace(" Wait For Lanched Process to finish", 2)
            If RedirectOutput Then
                xtrace("", 2)
                xtrace("==== Output of Launced process below ======================================", 2)
                xtrace(ExeString)
            End If

            Proc.StartInfo = ProcStartInfo
            If RunMinimized Then Proc.StartInfo.WindowStyle = ProcessWindowStyle.Minimized
            Proc.Start()
            If WaitForLanchedProcess Then Proc.WaitForExit()

        Catch ex As Exception
            WriteErr("Warning: Failed to start the process!")
            xtrace(ex.Message.ToString, 1)
            Console.Write("Enter to continiue.")
            Console.ReadLine()
        End Try

        xtrace_sube("Execute_NotElevated")
    End Sub
End Module
