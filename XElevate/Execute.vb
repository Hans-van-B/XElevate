Module Execute
    Sub Execute_Elevated(Launchfile As String, ExeString As String)
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

        Catch ex As Exception
            WriteErr("Warning: Failed to start the elevated process!")
            xtrace(ex.Message.ToString, 1)
            Console.Write("Enter to continiue.")
            Console.ReadLine()
        End Try
    End Sub

    Sub Execute_NotElevated(Launchfile As String, ExeString As String)
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
    End Sub
End Module
