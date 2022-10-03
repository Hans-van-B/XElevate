Module Util

    '==== Wait ================================================================
    Public Sub wait(ByVal interval As Integer)
        xtrace_subs("wait")

        xtrace_i("Wait " & interval.ToString)
        interval = interval * 1000

        Dim sw As New Stopwatch
        sw.Start()
        Do While sw.ElapsedMilliseconds < interval

        Loop
        sw.Stop()

        xtrace_sube("wait")
    End Sub

    '---- Get output from OS Command ------------------------------------------
    Function GetOutputFromCmd(ByVal Cmd As String) As String()
        xtrace_subs("GetOutputFromCmd")

        Dim proc As New Process
        Dim Output As String = ""
        With proc.StartInfo
            .FileName = "cmd.exe"
            .Arguments = "/c " & Cmd
            .UseShellExecute = False
            .RedirectStandardError = True
            .RedirectStandardInput = True
            .RedirectStandardOutput = True
            .WindowStyle = ProcessWindowStyle.Hidden
            .CreateNoWindow = True
        End With

        xtrace_i("Execute the command")
        Try
            proc.Start()
            Output = proc.StandardOutput.ReadToEnd()
            proc.WaitForExit()

            xtrace(Output)
        Catch ex As Exception
            xtrace(ex.Message)
        End Try

        GetOutputFromCmd = Split(Output, vbNewLine)

        xtrace_sube("GetOutputFromCmd")
    End Function
End Module
