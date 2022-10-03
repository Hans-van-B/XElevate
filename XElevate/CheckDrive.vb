Module CheckDrive
    Dim Drv2 As String
    Dim Drive_Found As Boolean = False
    Sub CheckDriveInElevatedProcess()
        xtrace_subs("CheckDriveInElevatedProcess")

        xtrace("Prepare: Check Drive In Elevated Process", 2)
        xtrace(" - Executable = " & ToExecute, 2)

        ' Check if required
        Drv2 = Left(ToExecute, 2)

        xtrace(" - Drive = " & Drv2, 2)
        If Drv2 = "\\" Then
            xtrace(" - UNC path", 2)
            Exit Sub
        End If

        If Drv2 = "C:" Then
            xtrace(" - Local drive", 2)
            Exit Sub
        End If

        WriteLanchFile("if exist " & Drv2 & "\. if not exist """ & ToExecute & """ Net Use /d " & Drv2)

        CheckNetUse()
        If Not Drive_Found Then CheckSubst()

        xtrace_sube("CheckDriveInElevatedProcess")
    End Sub

    '==== Check Net-Use =======================================================
    ' 2022-10-03 | HvB | Improved read OS command output implemented
    Dim DriveTraceLevel As Integer = 3

    Sub CheckNetUse()
        xtrace_subs("CheckNetUse")

        xtrace(" - Get drive mapping", 2)
        Dim UNCPath As String = "-"

        'Dim proc As ProcessStartInfo = New ProcessStartInfo("cmd.exe")
        'Dim pr As Process
        Dim line As String
        'Dim cnt As Integer
        Dim Index1 As Integer
        Dim Index2 As Integer
        'Dim Index3 As Integer

        ' Debug drive detection
        Dim FixDriveDebugFile = ExecDir & "\FixDriveDebug.txt"

        If (System.IO.File.Exists(FixDriveDebugFile)) Then
            xtrace(" - Found: " & FixDriveDebugFile)
            DriveTraceLevel = 2
        End If

        Dim Output() As String = GetOutputFromCmd("net use")

        For Each line In Output
            Index1 = InStr(line, Drv2)
            Index2 = InStr(line, "\\")
            xtrace("    | " & line, 3)

            If (Index1 > 1) And (Index2 > 1) Then
                UNCPath = Mid(line, Index2)
                xtrace(" - UNCPath = " & UNCPath, 2)
                Exit For
            End If
        Next

        'proc.CreateNoWindow = True
        'proc.UseShellExecute = False
        'proc.RedirectStandardInput = True
        'proc.RedirectStandardOutput = True

        'pr = Process.Start(proc)
        'pr.StandardInput.WriteLine("net use")

        'For cnt = 1 To 200
        '    If pr.StandardOutput.EndOfStream Then Exit For
        '    line = pr.StandardOutput.ReadLine()
        '    Index1 = InStr(line, Drv2)
        '    Index2 = InStr(line, "\\")
        '    Index3 = InStr(line, "command completed successfully", CompareMethod.Text)
        '    xtrace(cnt.ToString & "|" & Index1.ToString & "|" & Index2.ToString & "|" & line, DriveTraceLevel)
        '
        '    If Index3 > 0 Then Exit For
        '
        '    If (Index1 > 1) And (Index2 > 1) Then
        '       UNCPath = Mid(line, Index2)
        '       xtrace(" - UNCPath = " & UNCPath, 2)
        '        Exit For
        '    End If
        '
        'Next
        'pr.StandardOutput.Close()
        'pr.Close()

        '---- Write Mount
        If UNCPath <> "-" Then
            xtrace(" - Create mount procedure", 2)
            line = "if not exist " & Drv2 & "\. net use " & Drv2 & " " & UNCPath & " /PERSISTENT:NO"
            WriteLanchFile(line)
            Drive_Found = True
        Else
            xtrace(" - Drive not found, this may be a subst drive", 2)
        End If

        xtrace_sube("CheckNetUse")
    End Sub

    '==== Check Subst =========================================================
    Sub CheckSubst()
        xtrace_subs("CheckSubst")

        xtrace(" - Get drive substitutions", 2)
        Dim SPath As String = "-"

        Dim proc As ProcessStartInfo = New ProcessStartInfo("cmd.exe")
        'Dim pr As Process
        Dim line As String
        'Dim cnt As Integer
        Dim Index1 As Integer
        Dim Index2 As Integer

        Dim Output() As String = GetOutputFromCmd("subst")

        For Each line In Output
            Index1 = InStr(line, Drv2)
            Index2 = InStr(line, "=> ")
            xtrace("    | " & line, 3)

            If (Index1 >= 1) And (Index2 > 1) Then
                SPath = Mid(line, Index2 + 3)
                xtrace(" - Path = " & SPath, 2)
                Exit For
            End If
        Next

        'proc.CreateNoWindow = True
        'proc.UseShellExecute = False
        'proc.RedirectStandardInput = True
        'proc.RedirectStandardOutput = True

        'Try

        '    pr = Process.Start(proc)
        '    pr.StandardInput.WriteLine("subst")
        '    xtrace_i("subst issued")

        '    For cnt = 1 To 20
        '        If pr.StandardOutput.EndOfStream Then Exit For
        '        line = pr.StandardOutput.ReadLine()
        '        Index1 = InStr(line, Drv2)
        '        Index2 = InStr(line, "=> ")
        '        xtrace(cnt.ToString & "|" & Index1.ToString & "|" & Index2.ToString & "|" & line, DriveTraceLevel)
        '        If (Index1 >= 1) And (Index2 > 1) Then
        '            SPath = Mid(line, Index2 + 3)
        '            xtrace(" - Path = " & SPath, 2)
        '            Exit For
        '        End If
        '    Next
        '    pr.StandardOutput.Close()
        '    pr.Close()
        'Catch ex As Exception
        '   xtrace_Err({"Failed to execute subst", ex.Message})
        'End Try

        '---- Write Mount
        If SPath <> "-" Then
            xtrace(" - Create substitute procedure", 2)

            WriteLanchFile("if not exist " & Drv2 & "\. subst " & Drv2 & " """ & SPath & """")
            Drive_Found = True
        Else
            xtrace(" - Drive not found...", 2)
        End If

        xtrace_sube("CheckSubst")
    End Sub
End Module
