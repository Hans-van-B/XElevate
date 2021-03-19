Public Module Glob
    Public StartupDir As String = My.Application.Info.DirectoryPath
    Public UTmp As String

    Public AppName As String = "XElevate"
    Public AppVersion As String = "V0.08b"
    Public AppDate As String = "2020-10-09"

    ' Logging
    Public LogFile As String
    Public CTrace As Integer = 1
    Public LTrace As Integer = 2
    Public Debug As Boolean = False

    Public ToExecute As String = ""
    Public TeEx_Arg As String = ""
    Public RedirectOutput As Boolean = False
    Public UseCmd As Boolean = False
    Public UsePause As Boolean = False
    Public MaxFileNr As Integer = 99
    Public UseWait As Integer = 3
    Public BeBrief As Boolean = False
    Public BeQuiet As Boolean = False
    Public tl As Integer = 1            ' Trace level influenced by BeBrief
    Public ForceExit As Boolean = False
    Public ExitProgram As Boolean = False
    Public RunMinimized As Boolean = False
    Public IndexLaunchFile As Boolean = True
    Public EnvBufferSize As Integer = 600
    Public SetDebugInElevatedEnvironment = False
    Public WaitForLanchedProcess = False
    Public FixNetDrv As Boolean = False
    Public XElevateVar As String = "XElevate"

    Public ExecDir As String = "C:\Temp\TA\Elevate"
    Public Launchfile As String = ""

    ' Command-line switches
    Public TransferEnvironment As Boolean = False
    Public EnvironmentExceptions As Boolean = True
End Module
