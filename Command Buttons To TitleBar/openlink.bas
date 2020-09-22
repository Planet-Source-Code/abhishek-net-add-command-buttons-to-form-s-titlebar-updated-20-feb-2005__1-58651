Attribute VB_Name = "modOpenLink"
Option Explicit

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Const SW_SHOWNORMAL = 1

Public Sub OpenSite(UrlToOpen As String, ByVal hwnd As Long)
    On Error GoTo NoInternet
    ShellExecute hwnd, "open", UrlToOpen, "", 0, SW_SHOWNORMAL
    Exit Sub
NoInternet:
    MsgBox Err.Description, vbCritical, "Error"
End Sub
