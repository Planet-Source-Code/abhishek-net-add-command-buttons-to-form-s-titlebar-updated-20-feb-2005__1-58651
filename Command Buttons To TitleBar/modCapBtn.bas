Attribute VB_Name = "modCapBtn"
Option Explicit

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As Rect) As Long
Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook&, ByVal lpfn&, ByVal hmod&, ByVal dwThreadId&) As Long
Private Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook&) As Long

Private Type Rect
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type CWPSTRUCT
    lParam As Long
    wParam As Long
    Message As Long
    hwnd As Long
End Type

Private Const WM_MOVE = &H3
Private Const WM_SETCURSOR = &H20
Private Const WM_NCPAINT = &H85
Private Const WM_COMMAND = &H111

Private Const SWP_FRAMECHANGED = &H20
Private Const GWL_EXSTYLE = -20

Private WHook&

Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long

Private Const SM_CXSIZE = 30    'Width of title bar
Private Const SM_CYSIZE = 31    'height of title bar

Public Sub Init()
    WHook = SetWindowsHookEx(4, AddressOf HookProc, 0, App.ThreadID)
    Call SetWindowLong(frmMain.picbtn.hwnd, GWL_EXSTYLE, &H80)
    Call SetParent(frmMain.picbtn.hwnd, GetParent(frmMain.hwnd))
End Sub

Public Sub Terminate()
    Call UnhookWindowsHookEx(WHook)
    Call SetParent(frmMain.picbtn.hwnd, frmMain.hwnd)
End Sub

Public Function HookProc&(ByVal nCode&, ByVal wParam&, Inf As CWPSTRUCT)
    Dim R As Rect
    Static LastParam&

    If Inf.hwnd = GetParent(frmMain.picbtn.hwnd) Then
        If Inf.Message = WM_COMMAND Then
            Select Case LastParam
            Case frmMain.picbtn.hwnd: Call frmMain.Command1_Click
            End Select
        ElseIf Inf.Message = WM_SETCURSOR Then
            LastParam = Inf.wParam
        End If
    ElseIf Inf.hwnd = frmMain.hwnd Then
        If Inf.Message = WM_NCPAINT Or Inf.Message = WM_MOVE Then
            Call GetWindowRect(frmMain.hwnd, R)
            Call SetWindowPos(frmMain.picbtn.hwnd, 0, R.Right - 110, _
            R.Top + 4, Str$(GetSystemMetrics(SM_CXSIZE)), Str$(GetSystemMetrics(SM_CYSIZE)), SWP_FRAMECHANGED)
        End If
    End If
End Function
