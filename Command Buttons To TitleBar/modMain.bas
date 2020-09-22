Attribute VB_Name = "modMain"
Option Explicit
Private Declare Function InitCommonControls Lib "Comctl32.dll" () As Long

Public Sub Main()
    Dim x As Long
    x = InitCommonControls
    frmMain.Show
End Sub
