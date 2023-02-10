Module Module2


    '#07/10/01#
    '
    Structure CWPSTRUCT
        Dim lParam As Long
        Dim wParam As Long
        Dim Message As Long
        Dim hWnd As Long
    End Structure

    Structure NOTIFYICONDATA
        Dim cbSize As Long
        Dim hWnd As Long
        Dim uID As Long
        Dim uFlags As Long
        Dim uCallbackMessage As Long
        Dim hIcon As Long
        Dim szTip As String
End Structure

        Const NIM_ADD = 0
        Const NIM_MODIFY = 1
        Const NIM_DELETE = 2
        Const NIF_MESSAGE = 1
        Const NIF_ICON = 2
        Const NIF_TIP = 4

        Public Class Win32
            Declare Auto Function MessageBox Lib "user32.dll" ( _
                ByVal hWnd As Integer, ByVal txt As String, _
                ByVal caption As String, ByVal Type As Integer) _
                As Integer
        End Class









        Public Const WH_CALLWNDPROC = 4
        Public Const WM_CREATE = &H1

        Public Const WM_MOUSEMOVE = &H200
        Public Const WM_LBUTTONDOWN = &H201
        Public Const WM_LBUTTONUP = &H202
        Public Const WM_LBUTTONDBLCLK = &H203
        Public Const WM_RBUTTONDOWN = &H204
        Public Const WM_RBUTTONUP = &H205
        Public Const WM_RBUTTONDBLCLK = &H206
        Public Const WM_MBUTTONDOWN = &H207
        Public Const WM_MBUTTONUP = &H208
        Public Const WM_MBUTTONDBLCLK = &H209

        Public hHook As Long

        Public Sub Iconify(ByVal FRM As Object, ByVal Tip As String)
            Dim I As Integer
            Dim nid As NOTIFYICONDATA

            nid = setNOTIFYICONDATA(FRM.hWnd, vbNull, NIF_MESSAGE Or NIF_ICON Or NIF_TIP, WM_MOUSEMOVE, FRM.Icon, Tip)
        I = shell32.Shell_NotifyIconA(NIM_ADD, nid)

        If FRM.WindowState <> 1 Then FRM.WindowState = 1
            FRM.Visible = False
        End Sub

    Public Sub DeIconify(ByVal FRM As Object, Optional ByVal ProgEnd As Boolean = False)
        Dim I As Integer
        Dim nid As NOTIFYICONDATA

        If Not ProgEnd Then
            FRM.WindowState = vbNormal
            FRM.Visible = True
        End If
        nid = setNOTIFYICONDATA(FRM.hWnd, vbNull, NIF_MESSAGE Or NIF_ICON Or NIF_TIP, vbNull, FRM.Icon, "")
        I = shell32.Shell_NotifyIconA(NIM_DELETE, nid)

    End Sub

        Public Sub UpdIcon(ByVal FRM As Form, ByVal Tip As String)
            Dim I As Integer
            Dim nid As NOTIFYICONDATA

        nid = setNOTIFYICONDATA(FRM.Handle, vbNull, NIF_MESSAGE Or NIF_ICON Or NIF_TIP, WM_MOUSEMOVE, 1, Tip)
        I = shell32.Shell_NotifyIconA(NIM_MODIFY, nid)
        End Sub

        Public Function setNOTIFYICONDATA(ByVal hWnd As Long, ByVal Id As Long, ByVal Flags As Long, ByVal CallbackMessage As Long, ByVal Icon As Long, ByVal Tip As String) As NOTIFYICONDATA
            Dim nidTemp As NOTIFYICONDATA


            nidTemp.hWnd = hWnd
            nidTemp.uID = Id
            nidTemp.uFlags = Flags
            nidTemp.uCallbackMessage = CallbackMessage
            nidTemp.hIcon = Icon
        nidTemp.szTip = Tip & Chr(0)

            setNOTIFYICONDATA = nidTemp
        End Function

        Public Function AppHook(ByVal idHook As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
            Dim CWP As CWPSTRUCT
        user32.CopyMemory(CWP, lParam, Len(CWP))
            Select Case CWP.Message
                Case WM_CREATE
                user32.SetForegroundWindow(CWP.hWnd)
                AppHook = user32.CallNextHookEx(hHook, idHook, wParam, lParam)
                user32.UnhookWindowsHookEx(hHook)
                    hHook = 0
                    Exit Function
            End Select
        AppHook = user32.CallNextHookEx(hHook, idHook, wParam, lParam)
        End Function


End Module
