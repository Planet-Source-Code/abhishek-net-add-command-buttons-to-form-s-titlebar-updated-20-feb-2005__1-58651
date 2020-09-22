VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Command Buttons at TitleBar"
   ClientHeight    =   2490
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   4905
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   2490
   ScaleWidth      =   4905
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picbtn 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      ScaleHeight     =   375
      ScaleWidth      =   495
      TabIndex        =   0
      Top             =   120
      Width           =   495
      Begin VB.CommandButton Command1 
         Caption         =   "î"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   0
         TabIndex        =   1
         ToolTipText     =   "Minimize To SystemTray"
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "J"
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   26.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   570
      Left            =   240
      TabIndex        =   3
      Top             =   1680
      Width           =   465
   End
   Begin VB.Label Label1 
      Caption         =   $"frmMain.frx":0000
      Height          =   975
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   4695
   End
   Begin VB.Menu MnuPopUpMain 
      Caption         =   "PopUp"
      Visible         =   0   'False
      Begin VB.Menu MnuPopUp 
         Caption         =   "Show"
         Index           =   0
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Need To Be Public if u wanna call from outside!
Public Sub Command1_Click()
    Const strTitle As String = "This is One of the Uses of Adding Buttons To The TitleBar"
    Const strInfo As String = "Adding SysTray Icons is for Demostration Purposes Only and No Way Related to Adding Buttons to the TitleBar Project."
    
    Call modTray.AddTrayIcon(Me.hwnd, Me.Icon, "Hi!")
    
    Me.WindowState = vbMinimized
    Call modTray.ShowBalloon(Me.hwnd, strTitle, strInfo, [Information Icon])
    Me.Hide
End Sub

Private Sub Form_Load()
    Call Init
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If MsgBox("DO U WISH TO VOTE ?", vbQuestion + vbYesNo) = vbYes Then
        Call modOpenLink.OpenSite("http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=58651&lngWId=1", Me.hwnd)
    End If
    Call modTray.RemoveTrayIcon(Me.hwnd)
    Call Terminate
End Sub

Private Sub MnuPopUp_Click(Index As Integer)
    Me.Show
    Me.WindowState = vbNormal
End Sub

Private Sub picbtn_Resize()
    Command1.Width = picbtn.Width
    Command1.Height = picbtn.Height
End Sub
