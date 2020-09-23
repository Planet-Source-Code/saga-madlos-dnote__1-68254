VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H80000009&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7230
   ClientLeft      =   3945
   ClientTop       =   1560
   ClientWidth     =   7740
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmMain.frx":0000
   ScaleHeight     =   7230
   ScaleWidth      =   7740
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtNote 
      BackColor       =   &H00E0E0E0&
      Height          =   5655
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   840
      Width           =   4215
   End
   Begin VB.Image imgPaste_down 
      Height          =   300
      Left            =   6000
      Picture         =   "frmMain.frx":62E4C
      Top             =   1485
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Image imgPaste_up 
      Height          =   300
      Left            =   4740
      Picture         =   "frmMain.frx":64150
      Top             =   1485
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Image imgCopy_down 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   5970
      Picture         =   "frmMain.frx":65454
      Top             =   1050
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Image imgCopy_up 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   4695
      Picture         =   "frmMain.frx":66758
      Top             =   1020
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Image imgClear_down 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   5985
      Picture         =   "frmMain.frx":67A5C
      Top             =   660
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Image imgClear_up 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   4665
      Picture         =   "frmMain.frx":68D60
      Top             =   645
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Image imgClose 
      Height          =   165
      Left            =   4260
      Picture         =   "frmMain.frx":6A064
      ToolTipText     =   "Hide Dnote"
      Top             =   90
      Width           =   135
   End
   Begin VB.Image imgClose_down 
      Height          =   165
      Left            =   6060
      Picture         =   "frmMain.frx":6A1DC
      Top             =   375
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Image imgClear 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   345
      Picture         =   "frmMain.frx":6A354
      Top             =   465
      Width           =   1200
   End
   Begin VB.Image imgCopy 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   1650
      Picture         =   "frmMain.frx":6B658
      Top             =   465
      Width           =   1200
   End
   Begin VB.Image imgPaste 
      Height          =   300
      Left            =   2955
      Picture         =   "frmMain.frx":6C95C
      Top             =   465
      Width           =   1200
   End
   Begin VB.Image imgClose_up 
      Appearance      =   0  'Flat
      Height          =   165
      Left            =   5715
      Picture         =   "frmMain.frx":6DC60
      Top             =   360
      Visible         =   0   'False
      Width           =   135
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Me.BackColor = RGB(0, 0, 255)   ':: set the background color of the form
                                    ':: this color will be rendered transparent
    ':: set the form transaparent
    If setWindowTransparent(Me.hWnd, RGB(0, 0, 255)) = False Then ':: if unsuccesfull
        MsgBox TransparentWindow.TRANSWIN_ERROR ':: display the error message
    End If
    ':: set the window as the topmost window
    Call SetWindowPos(Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE)
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
':: this procedure will  make the form dragable. Just click on the form surface
':: and drag to any position in the screen
    If Button = vbLeftButton Then
        Call ReleaseCapture
        Call SendMessage(Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0)
    End If

End Sub

Private Sub imgClear_Click()
    ':: clear/empty text box
    txtNote.Text = ""
End Sub

Private Sub imgClear_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgClear.Picture = imgClear_down.Picture
End Sub

Private Sub imgClear_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgClear.Picture = imgClear_up.Picture
End Sub

Private Sub imgClose_Click()
    Call showNote
End Sub

Private Sub imgClose_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgClose.Picture = imgClose_down.Picture
End Sub

Private Sub imgClose_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgClose.Picture = imgClose_up.Picture
End Sub

Private Sub imgCopy_Click()
    ':: copy text in the textbox to the clipboard
    Call Clipboard.Clear ':: empty clipboard first before copying any data into it
    Call Clipboard.SetText(txtNote.SelText) ':: copy selected text in textbox to clipboard
End Sub

Private Sub imgCopy_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgCopy.Picture = imgCopy_down.Picture
End Sub

Private Sub imgCopy_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgCopy.Picture = imgCopy_up.Picture
End Sub

Private Sub imgPaste_Click()
    ':: paste text from the clipboard to the location of the cursor in the textbox
    If Clipboard.GetFormat(CF_TEXT) Then ':: check first if format is text data
        txtNote.SelText = Clipboard.GetText
    End If
End Sub

Private Sub imgPaste_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgPaste.Picture = imgPaste_down.Picture
End Sub

Private Sub imgPaste_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgPaste.Picture = imgPaste_up.Picture
End Sub

':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'::                         CUSTOM PROCEDURES (SUB AND FUNCTION)
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
':: NAME        :
':: TYPE        :
':: SCOPE       :
':: PARAMETERS  :
':: RETURN      :
':: DESCRIPTION :
'::...................................................................................
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::


