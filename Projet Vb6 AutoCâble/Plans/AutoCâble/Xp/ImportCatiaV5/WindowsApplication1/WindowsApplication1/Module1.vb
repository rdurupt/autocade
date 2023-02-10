Module Module1
    Public Event KeyDown As KeyEventHandler

    Public Const SWP_NOMOVE = &H2
    Public Const SWP_NOSIZE = &H1
    Public Const SWP_NOACTIVATE = &H10
    Public Const SWP_SHOWWINDOW = &H40
    Public Const HWND_TOPMOST = -1
    Public Const HWND_NOTOPMOST = -2
    Public Const Flags = SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
    Public Class user32
        Public Declare Function GetAsyncKeyState Lib "user32" Alias "GetAsyncKeyState" (ByVal vKey As Long) As Integer

        Public Declare Sub keybd_event Lib "user32.dll" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
        Public Declare Function SetWindowPos Lib "user32.dll" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
        Public Declare Function CallNextHookEx Lib "user32.dll" (ByVal hHook As Long, ByVal ncode As Long, ByVal wParam As Long, ByVal lParam As Object) As Long
        Public Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByVal hpvDest As Object, ByVal hpvSource As Object, ByVal cbCopy As Long)
        Public Declare Function SetForegroundWindow Lib "user32.dll" (ByVal hWnd As Long) As Long
        Public Declare Function UnhookWindowsHookEx Lib "user32.dll" (ByVal hHook As Long) As Long
        Public Declare Function SetWindowsHookEx Lib "user32.dll" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
        Public Declare Function GetWindowThreadProcessId Lib "user32.dll" (ByVal hWnd As Long, ByVal lpdwProcessId As Long) As Long
        Public Declare Function SetCapture Lib "user32.dll" (ByVal hWnd As Long) As Long

    End Class

    Public Class shell32
        Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
        Declare Function Shell_NotifyIconA Lib "shell32.dll" (ByVal dwMessage As Long, ByVal lpData As Object) As Integer

    End Class
    Public Class gdi32

        Public Declare Function GetPath Lib "gdi32.dll" (ByVal hdc As Long, ByVal lpPoint As Object, ByVal lpTypes As Byte, ByVal nSize As Long) As Long
    End Class



    Public Sub MyExecute(ByVal Fichier As String)

        Dim lapi As Long
        On Error Resume Next
        lapi = shell32.ShellExecute(100, "open", Fichier, "", vbNull, 5)
        'If Err Then MsgBox Err.Description
        Err.Clear()
        On Error GoTo 0
    End Sub

End Module
