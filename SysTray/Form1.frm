VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "System Quick Tray"
   ClientHeight    =   1815
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6270
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1815
   ScaleWidth      =   6270
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   1  'Minimized
   Begin VB.Frame Frame1 
      Caption         =   "System Quick Tray"
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5895
      Begin VB.CommandButton Command1 
         Caption         =   "&OK"
         Height          =   375
         Left            =   4320
         TabIndex        =   2
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000005&
         X1              =   100
         X2              =   5760
         Y1              =   950
         Y2              =   950
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Steve Camm"
         Height          =   195
         Left            =   4680
         TabIndex        =   1
         Top             =   720
         Width           =   900
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000003&
         X1              =   100
         X2              =   5760
         Y1              =   930
         Y2              =   930
      End
      Begin VB.Image Image1 
         Height          =   240
         Left            =   360
         Picture         =   "Form1.frx":0442
         Top             =   360
         Width           =   240
      End
   End
   Begin VB.Menu mPopupSys 
      Caption         =   "&SysTray"
      Visible         =   0   'False
      Begin VB.Menu mPopODBC 
         Caption         =   "ODBC Admin"
      End
      Begin VB.Menu mPopVisualBasic 
         Caption         =   "&Visual Basic"
      End
      Begin VB.Menu mPopOutlook 
         Caption         =   "Outlook"
      End
      Begin VB.Menu mnuSpace2 
         Caption         =   "-"
      End
      Begin VB.Menu mPopRestore 
         Caption         =   "&About..."
      End
      Begin VB.Menu mnuSpace1 
         Caption         =   "-"
      End
      Begin VB.Menu mPopExit 
         Caption         =   "&Exit"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ------------------------------------------------------------------
' Author        S. Camm
' Date          10/01/2002
' Description   Hidden Form to add Menu for Systray
' Notes         My Apologies for not renaming the form and controls.
' ------------------------------------------------------------------

Private Sub Command1_Click()
 Me.Hide
 Me.WindowState = vbMinimized
End Sub

Private Sub Form_Load()
 'the form must by fully visble before calling shell_notifyIcon
 Me.Show
 Me.Refresh
 With nid
   .cbSize = Len(nid)
   .hwnd = Me.hwnd
   .uId = vbNull
   .uFlags = NIF_ICON Or NIT_TIP Or NIF_MESSAGE
   .uCallBackMessage = WM_MOUSEMOVE
   .hIcon = Me.Icon
   .szTip = "My Tooltip" & vbNullChar
 End With
 Shell_NotifyIcon NIM_ADD, nid
 
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 ' this procedure recieves the callbacks from the system tray icon
 Dim result As Long
 Dim msg As Long
 ' the value of X will vary depending upon the scalemode setting
 If Me.ScaleMode = vbPixels Then
    msg = X
 Else
   msg = X / Screen.TwipsPerPixelX
 End If
 Select Case msg
  Case WM_LBUTTONUP             '514 restore form window
   Me.WindowState = vbNormal
   result = SetForegroundWindow(Me.hwnd)
   Me.Show
  Case WM_LBUTTONDBCLICK        '515 restore form window
   Me.WindowState = vbNormal
   result = SetForegroundWindow(Me.hwnd)
   Me.Show
  Case WM_RBUTTONUP             '517 display popup menu
   result = SetForegroundWindow(Me.hwnd)
   Me.PopupMenu Me.mPopupSys
  End Select
End Sub

Private Sub Form_Resize()
 'this is necessary to assure  that the minimized windows is hidden
 If Me.WindowState = vbMinimized Then Me.Hide
 
End Sub

Private Sub Form_Unload(Cancel As Integer)
 ' this removes the icon from the system tray
 Shell_NotifyIcon NIM_DELETE, nid
End Sub

Private Sub mPopExit_CLick()
 ' called when user clicks the popup menu exit command
 Unload Me
End Sub

Private Sub mPopRestore_click()
 ' called when the user clicks the popup menu restore command
 Dim result As Long
 Me.WindowState = vbNormal
 result = SetForegroundWindow(Me.hwnd)
 Me.Show
End Sub

Private Sub mPopODBC_Click()
 lngRC = Shell("ODBCAD32.EXE", vbNormalFocus)
End Sub

Private Sub mPopVisualBasic_Click()
' Eeeek - Applications Paths hard coded
 lngRC = Shell("C:\Program Files\Microsoft Visual Studio\VB98\VB6.EXE", vbMaximizedFocus)
End Sub
Private Sub mPopOutlook_Click()
 ' Eeeek - Application Paths hard coded
 lngRC = Shell("C:\Program Files\Microsoft Office\Office\OUTLOOK.EXE", vbNormalFocus)
End Sub


