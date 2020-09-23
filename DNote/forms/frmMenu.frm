VERSION 5.00
Begin VB.Form frmMenu 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Menu"
   ClientHeight    =   2175
   ClientLeft      =   150
   ClientTop       =   840
   ClientWidth     =   6435
   Icon            =   "frmMenu.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2175
   ScaleWidth      =   6435
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Label Label1 
      Caption         =   $"frmMenu.frx":030A
      Height          =   945
      Left            =   225
      TabIndex        =   0
      Top             =   720
      Width           =   5655
   End
   Begin VB.Image imgTrayIcon 
      Height          =   480
      Left            =   210
      Picture         =   "frmMenu.frx":0456
      Top             =   120
      Width           =   480
   End
   Begin VB.Menu mnuPop 
      Caption         =   "Popup Menu"
      Begin VB.Menu mnuClose 
         Caption         =   "Close"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public sysTrayIcon As New TrayIcon

Private Sub Form_Load()
    ':: add tray icon
    If sysTrayIcon.addTrayIcon(Me.hWnd, Me.imgTrayIcon, "DNote - Saga") = False Then
        MsgBox "Unable to  add icon to system tray. Program closing", vbCritical, "DNote - Saga"
        Call KeyboardHook.unhookKeyboard
        Unload frmMain
        sysTrayIcon.removeTrayIcon
        Unload Me
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Select Case (X / Screen.TwipsPerPixelX)
        Case WM_RBUTTONUP: ':: when right click on the icon in the tray
            PopupMenu mnuPop ':: show pop up menu
        Case WM_LBUTTONUP:
            Call showNote
        End Select
End Sub

Private Sub mnuAbout_Click()
    frmAbout.Show
End Sub

Private Sub mnuClose_Click()
    If MsgBox("Are you sure you want to close DNote? This will terminate the program and its feature.", vbYesNo, "DNote - Saga") = vbYes Then
        Call mdlMain.terminateApplication
    End If
End Sub

