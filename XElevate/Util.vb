Module Util

    '==== Wait ================================================================
    Public Sub wait(ByVal interval As Integer)
        xtrace_i("Wait " & interval.ToString)
        interval = interval * 1000

        Dim sw As New Stopwatch
        sw.Start()
        Do While sw.ElapsedMilliseconds < interval

        Loop
        sw.Stop()
    End Sub


End Module
