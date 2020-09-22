Attribute VB_Name = "modTray"
Option Explicit
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Copyright ©1996-2004 VBnet, Randy Birch, All Rights Reserved.
' Some pages may also contain other copyrights by the author.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Distribution: You can freely use this code in your own
'               applications, but you may not reproduce
'               or publish this code on any web site,
'               online service, or distribute as source
'               on any media without express permission.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'Modified Version of the Orignal SysTray Code by Randy Birch (VBnet)
'Modified By Abhishek, abhishek007p@hotmail.com


'defWindowProc holds the address
'of the default window message processing
'procedure returned by SetWindowLong
Private defWindowProc As Long

'flag preventing re-creating the timer
Private tmrRunning As Boolean

'Get/SetWindowLong messages
Private Const GWL_WNDPROC As Long = (-4)
'Private Const GWL_HWNDPARENT As Long = (-8)
'Private Const GWL_ID As Long = (-12)
'Private Const GWL_STYLE As Long = (-16)
'Private Const GWL_EXSTYLE As Long = (-20)
'Private Const GWL_USERDATA As Long = (-21)

'general windows messages
Private Const WM_USER As Long = &H400
'Private Const WM_NOTIFY As Long = &H4E
'Private Const WM_COMMAND As Long = &H111
'Private Const WM_CLOSE As Long = &H10
Private Const WM_TIMER = &H113

'mouse constants for the callback
'Private Const WM_LBUTTONDOWN As Long = &H201
'Private Const WM_LBUTTONUP As Long = &H202
Private Const WM_LBUTTONDBLCLK As Long = &H203
'Private Const WM_MBUTTONDOWN As Long = &H207
'Private Const WM_MBUTTONUP As Long = &H208
'Private Const WM_MBUTTONDBLCLK As Long = &H209
'Private Const WM_RBUTTONDOWN As Long = &H204
Private Const WM_RBUTTONUP As Long = &H205
'Private Const WM_RBUTTONDBLCLK As Long = &H206

'private message the shell_notify api will pass
'to WindowProc when our systray icon is acted upon
Private Const WM_MYHOOK As Long = WM_USER + 1

'ID constant representing this
'application in the systray
Private Const APP_SYSTRAY_ID = 979

'ID constant representing this
'application for SetTimer
Private Const APP_TIMER_EVENT_ID As Long = 979

'const holding number of milliseconds to timeout
'10000=10 seconds
Private Const APP_TIMER_MILLISECONDS As Long = 10000

'balloon tip notification messages
Private Const NIN_BALLOONSHOW = (WM_USER + 2)
Private Const NIN_BALLOONHIDE = (WM_USER + 3)
Private Const NIN_BALLOONTIMEOUT = (WM_USER + 4)
Private Const NIN_BALLOONUSERCLICK = (WM_USER + 5)

'shell version / NOTIFIYICONDATA struct size constants
Private Const NOTIFYICONDATA_V1_SIZE As Long = 88  'pre-5.0 structure size
Private Const NOTIFYICONDATA_V2_SIZE As Long = 488 'pre-6.0 structure size
Private Const NOTIFYICONDATA_V3_SIZE As Long = 504 '6.0+ structure size
Private NOTIFYICONDATA_SIZE As Long

Private Const NOTIFYICON_VERSION = &H3

'shell_notify flags
Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4
'Private Const NIF_STATE = &H8
Private Const NIF_INFO = &H10
'shell_notify messages
Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
'Private Const NIM_SETFOCUS = &H3
Private Const NIM_SETVERSION = &H4
'Private Const NIM_VERSION = &H5
'shell_notify styles
'Private Const NIS_HIDDEN = &H1
Private Const NIS_SHAREDICON = &H2

Private Type GUID
   Data1 As Long
   Data2 As Integer
   Data3 As Integer
   Data4(7) As Byte
End Type

Private Type NOTIFYICONDATA
  cbSize As Long
  hwnd As Long
  uID As Long
  uFlags As Long
  uCallbackMessage As Long
  hIcon As Long
  szTip As String * 128
  dwState As Long
  dwStateMask As Long
  szInfo As String * 256
  uTimeoutAndVersion As Long
  szInfoTitle As String * 64
  dwInfoFlags As Long
  guidItem As GUID
End Type

Private Declare Function SetForegroundWindow Lib "user32" _
   (ByVal hwnd As Long) As Long
   
'Private Declare Function PostMessage Lib "user32" _
   Alias "PostMessageA" _
   (ByVal hwnd As Long, _
    ByVal wMsg As Long, _
    ByVal wParam As Long, _
    lParam As Any) As Long
    
Private Declare Function SetWindowLong Lib "user32" _
   Alias "SetWindowLongA" _
   (ByVal hwnd As Long, _
    ByVal nIndex As Long, _
    ByVal dwNewLong As Any) As Long

Private Declare Function CallWindowProc Lib "user32" _
   Alias "CallWindowProcA" _
   (ByVal lpPrevWndFunc As Long, _
    ByVal hwnd As Long, _
    ByVal uMsg As Long, _
    ByVal wParam As Long, _
    ByVal lParam As Long) As Long

Private Declare Function SetTimer Lib "user32" _
  (ByVal hwnd As Long, _
   ByVal nIDEvent As Long, _
   ByVal uElapse As Long, _
   ByVal lpTimerFunc As Long) As Long
   
Private Declare Function KillTimer Lib "user32" _
  (ByVal hwnd As Long, _
   ByVal nIDEvent As Long) As Long

Private Declare Function Shell_NotifyIcon Lib "shell32.dll" _
   Alias "Shell_NotifyIconA" _
  (ByVal dwMessage As Long, _
   lpData As NOTIFYICONDATA) As Long
   
Private Declare Function GetFileVersionInfoSize Lib "version.dll" _
   Alias "GetFileVersionInfoSizeA" _
  (ByVal lptstrFilename As String, _
   lpdwHandle As Long) As Long

Private Declare Function GetFileVersionInfo Lib "version.dll" _
   Alias "GetFileVersionInfoA" _
  (ByVal lptstrFilename As String, _
   ByVal dwHandle As Long, _
   ByVal dwLen As Long, _
   lpData As Any) As Long
   
Private Declare Function VerQueryValue Lib "version.dll" _
   Alias "VerQueryValueA" _
  (pBlock As Any, _
   ByVal lpSubBlock As String, _
   lpBuffer As Any, _
   nVerSize As Long) As Long

Private Declare Sub CopyMemory Lib "kernel32" _
   Alias "RtlMoveMemory" _
  (Destination As Any, _
   Source As Any, _
   ByVal Length As Long)
   
Public Enum eIconType
    [No Icon] = &H0
    [Information Icon] = &H1
    [Warning Icon] = &H2
    [Error Icon] = &H3
    [NIIF_GUID] = &H5
    [NIIF_ICON_MASK] = &HF
    [NIIF_NOSOUND] = &H10
End Enum

Private Function IsShellVersion(ByVal version As Long) As Boolean

  'returns True if the Shell version
  '(shell32.dll) is equal or later than
  'the value passed as 'version'
   Dim nBufferSize As Long
   Dim nUnused As Long
   Dim lpBuffer As Long
   Dim nVerMajor As Integer
   Dim bBuffer() As Byte
   
   Const sDLLFile As String = "shell32.dll"
   
   nBufferSize = GetFileVersionInfoSize(sDLLFile, nUnused)
   
   If nBufferSize > 0 Then
    
      ReDim bBuffer(nBufferSize - 1) As Byte
    
      Call GetFileVersionInfo(sDLLFile, 0&, nBufferSize, bBuffer(0))
    
      If VerQueryValue(bBuffer(0), "\", lpBuffer, nUnused) = 1 Then
         
         CopyMemory nVerMajor, ByVal lpBuffer + 10, 2
        
         IsShellVersion = nVerMajor >= version
      
      End If  'VerQueryValue
    
   End If  'nBufferSize
  
End Function

Private Sub SetShellVersion()

   Select Case True
      Case IsShellVersion(6)
         NOTIFYICONDATA_SIZE = NOTIFYICONDATA_V3_SIZE '6.0 structure size
      
      Case IsShellVersion(5)
         NOTIFYICONDATA_SIZE = NOTIFYICONDATA_V2_SIZE 'pre-6.0 structure size
      
      Case Else
         NOTIFYICONDATA_SIZE = NOTIFYICONDATA_V1_SIZE 'pre-5.0 structure size
   End Select

End Sub

Public Sub AddTrayIcon(hwnd As Long, Icon As StdPicture, sToolTip As String)
   
   Dim nid As NOTIFYICONDATA
   
   If NOTIFYICONDATA_SIZE = 0 Then SetShellVersion
     
   With nid
      .cbSize = NOTIFYICONDATA_SIZE
      .hwnd = hwnd
      .uID = APP_SYSTRAY_ID
      .uFlags = NIF_MESSAGE Or NIF_ICON Or NIF_TIP Or NIF_INFO
      .dwState = NIS_SHAREDICON
      .hIcon = Icon
      .szTip = sToolTip & vbNullChar
      .uTimeoutAndVersion = NOTIFYICON_VERSION
      .uCallbackMessage = WM_MYHOOK
   End With
   
  'add the icon ...
   If Shell_NotifyIcon(NIM_ADD, nid) = 1 Then
   
     '... and inform the system of the
     'NOTIFYICON version in use
      Call Shell_NotifyIcon(NIM_SETVERSION, nid)
      
     'prepare to receive the systray messages
      Call SubClass(hwnd)
      
   End If
       
End Sub

Public Sub RemoveTrayIcon(hwnd As Long)

   Dim nid As NOTIFYICONDATA
   
   If NOTIFYICONDATA_SIZE = 0 Then SetShellVersion
   
   With nid
      .cbSize = NOTIFYICONDATA_SIZE
      .hwnd = hwnd
      .uID = APP_SYSTRAY_ID
   End With
   
   If tmrRunning Then Call TimerStop(hwnd)
   Call Shell_NotifyIcon(NIM_DELETE, nid)

End Sub

Private Sub BalloonTipClose(hwnd As Long)

   Dim nid As NOTIFYICONDATA
   
   If NOTIFYICONDATA_SIZE = 0 Then SetShellVersion
   
   With nid
      .cbSize = NOTIFYICONDATA_SIZE
      .hwnd = hwnd
      .uID = APP_SYSTRAY_ID
      .uFlags = NIF_TIP Or NIF_INFO
      .szTip = vbNullChar
      .uTimeoutAndVersion = NOTIFYICON_VERSION
   End With
   
   Call Shell_NotifyIcon(NIM_MODIFY, nid)
   
End Sub

Public Sub ShowBalloon(hwnd As Long, sTitle As String, sMessage As String, eIcon As eIconType)

   Dim nid As NOTIFYICONDATA
   
   If NOTIFYICONDATA_SIZE = 0 Then SetShellVersion
   
   With nid
      .cbSize = NOTIFYICONDATA_SIZE
      .hwnd = hwnd
      .uID = APP_SYSTRAY_ID
      .uFlags = NIF_INFO
      .dwInfoFlags = eIcon
      .szInfoTitle = sTitle & vbNullChar
      .szInfo = sMessage & vbNullChar
   End With

   Call Shell_NotifyIcon(NIM_MODIFY, nid)

End Sub

Private Sub SubClass(hwnd As Long)

  'assign our own window message
  'procedure (WindowProc)
   On Error Resume Next
   defWindowProc = SetWindowLong(hwnd, GWL_WNDPROC, AddressOf WindowProc)
   
End Sub

'Private Sub UnSubClass(hwnd As Long)
'  'restore the default message handling
'  'before exiting
'   If defWindowProc <> 0 Then
'      SetWindowLong hwnd, GWL_WNDPROC, defWindowProc
'      defWindowProc = 0
'   End If
'End Sub

Private Sub TimerBegin(ByVal hwndOwner As Long, ByVal dwMilliseconds As Long)

   If Not tmrRunning Then

      If dwMilliseconds <> 0 Then

        'SetTimer returns the event ID we
        'assign if it starts successfully,
        'so this is assigned to the Boolean
        'flag to indicate the timer is running.
         tmrRunning = SetTimer(hwndOwner, _
                               APP_TIMER_EVENT_ID, _
                               dwMilliseconds, _
                               AddressOf TimerProc) = APP_TIMER_EVENT_ID
         
         Debug.Print "timer started"

      End If

   End If

End Sub

Private Function TimerProc(ByVal hwnd As Long, _
                          ByVal uMsg As Long, _
                          ByVal idEvent As Long, _
                          ByVal dwTime As Long) As Long

   Select Case uMsg
      Case WM_TIMER

         If idEvent = APP_TIMER_EVENT_ID Then
            If tmrRunning = True Then
            
               Debug.Print "timer proc fired"
               Debug.Print "  shutting down balloon"
               Call TimerStop(hwnd)
               Call BalloonTipClose(frmMain.hwnd)
               
            End If  'tmrRunning
         End If  'idEvent

      Case Else
   End Select

End Function

Private Sub TimerStop(ByVal hwnd As Long)

   If tmrRunning = True Then

      Debug.Print "timer stopped"
      Call KillTimer(hwnd, APP_TIMER_EVENT_ID)
      tmrRunning = False

   End If

End Sub

Private Function WindowProc(ByVal hwnd As Long, _
                           ByVal uMsg As Long, _
                           ByVal wParam As Long, _
                           ByVal lParam As Long) As Long

  'If the handle returned is to our form,
  'call a message handler to deal with
  'tray notifications. If it is a general
  'system message, pass it on to
  'the default window procedure.
  '
  'If destined for the form and equal to
  'our custom hook message (WM_MYHOOK),
  'examining lParam reveals the message
  'generated, to which we react appropriately.
   On Error Resume Next
  
   Select Case hwnd
   
     'form-specific handler
      Case frmMain.hwnd
         
         Select Case uMsg
          'check uMsg for the application-defined
          'identifier (NID.uID) assigned to the
          'systray icon in NOTIFYICONDATA (NID).
  
           'WM_MYHOOK was defined as the message sent
           'as the .uCallbackMessage member of
           'NOTIFYICONDATA the systray icon
            Case WM_MYHOOK
            
              'lParam is the value of the message
              'that generated the tray notification.
               Select Case lParam
                  Case WM_RBUTTONUP
                 'This assures that focus is restored to
                 'the form when the menu is closed. If the
                 'form is hidden, it (correctly) has no effect.
                  Call SetForegroundWindow(frmMain.hwnd)
                  frmMain.PopupMenu frmMain.MnuPopUpMain, , , , frmMain.MnuPopUp(0)
                  Debug.Print "WM_RBUTTONUP"
                  
                  Case WM_LBUTTONDBLCLK
                      frmMain.Show
                      App.TaskVisible = True
                      Debug.Print "WM_LBUTTONDBLCLK"
                      
                  Case NIN_BALLOONSHOW
                    'the balloon tip has just appeared so
                    'set the timer to automatically close it
                     Call TimerBegin(hwnd, APP_TIMER_MILLISECONDS)
                     Debug.Print "NIN_BALLOONSHOW"
     
                  Case NIN_BALLOONHIDE
                    'the balloon tip has just been hidden,
                    'either because of a user-click, the
                    'system timeout being reached, or our
                    'SetTimer timeout expiring, so ensure
                    'the timer has stopped.
                     Call TimerStop(hwnd)
                     Debug.Print "NIN_BALLOONHIDE"

                  Case NIN_BALLOONUSERCLICK
                    'the balloon tip was clicked so
                    'ensure the timer won't fire
                     Call TimerStop(hwnd)
                     Debug.Print "NIN_BALLOONUSERCLICK"
                            
                  Case NIN_BALLOONTIMEOUT
                    'the system timeout has been reached
                    'which causes the system to close the
                    'tip without intervention. The timer
                    'must also be stopped now. Note that
                    'this message does not fire if the
                    'balloon tip is closed through our
                    'SetTimer method!
                     Call TimerStop(hwnd)
                     Debug.Print "NIN_BALLOONTIMEOUT"
               
               End Select
            
           'handle any other form messages by
           'passing to the default message proc
            Case Else
            
               WindowProc = CallWindowProc(defWindowProc, _
                                            hwnd, _
                                            uMsg, _
                                            wParam, _
                                            lParam)
               Exit Function
            
         End Select
     
     'this takes care of messages when the
     'handle specified is not that of the form
      Case Else
      
          WindowProc = CallWindowProc(defWindowProc, _
                                      hwnd, _
                                      uMsg, _
                                      wParam, _
                                      lParam)
   End Select
End Function
