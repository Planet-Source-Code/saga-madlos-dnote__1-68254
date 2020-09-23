Attribute VB_Name = "TransparentWindow"
Option Explicit
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'::                                    WIN 32 API
'::..................................................................................
':: function sets the opacity and transparency color key of a layered window
':: for more info http://msdn.microsoft.com/library/default.asp?url=/library/en-us/winui/winui/windowsuserinterface/windowing/windows/windowreference/windowfunctions/setlayeredwindowattributes.asp
Private Declare Function SetLayeredWindowAttributes Lib "user32.dll" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
':: function retrieves information about the specified window
':: for more info http://msdn.microsoft.com/library/default.asp?url=/library/en-us/winui/winui/windowsuserinterface/windowing/windowclasses/windowclassreference/windowclassfunctions/getwindowlong.asp
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
':: function changes an attribute of the specified window
':: for more info http://msdn.microsoft.com/library/default.asp?url=/library/en-us/winui/winui/windowsuserinterface/windowing/windowclasses/windowclassreference/windowclassfunctions/setwindowlong.asp
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'::                                     CONSTANTS
'::...................................................................................
Const LWA_COLORKEY = &H1 ':: Use crKey as the transparency color.
Const LWA_ALPHA = &H2 ':: Use bAlpha to determine the opacity of the layered window.
Const GWL_EXSTYLE = (-20) ':: windows extended style
Const WS_EX_LAYERED = &H80000 ':: windows extended style - layared
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'::                               STRUCTURES, ENUMS AND DATA TYPES
'::...................................................................................
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'::                                 OTHER VARIABLES
'::...................................................................................
Public TRANSWIN_ERROR As String ':: holds error information
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'::                         PROCEDURES (SUB AND FUNCTION)
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
':: NAME        : setWindowOpacity
':: TYPE        : Function
':: SCOPE       : Public
':: PARAMETERS  : hWnd - windows handle of the form to be set
'::               transpaColor - This color will be the one to be adjusted whether it
'::                 will be transparent or opaque depending on the value of opacity.
'::                 For example if this is set to RGB(255,255,255)
'::                 then any white color on the form will be transparent.
'::                 User RGB() function when setting this parameter.
':: RETURN      : Boolean
'::                 True - successfull in setting opacity
'::                 False - error in setting opacity
':: DESCRIPTION : this set the opacity of a form. From transparent to opaque
'::...................................................................................
Public Function setWindowTransparent(ByVal hWnd As Long, ByVal transpaColor As Long) As Boolean
Dim exStyle As Long
    If transpaColor > RGB(255, 255, 255) Or transpaColor < RGB(0, 0, 0) Then
        ':: check if the given transpaColor is valid
        ':: set error message if not and return false
        TRANSWIN_ERROR = "Invalid transparent color. Use RGB() function to generate COLORREF value."
        setWindowTransparent = False
    Else
        exStyle = GetWindowLong(hWnd, GWL_EXSTYLE) ':: get the extended style
        exStyle = exStyle Or WS_EX_LAYERED ':: set layared style in the extended style
        Call SetWindowLong(hWnd, GWL_EXSTYLE, exStyle) ':: set extended style of form
        ':: set window opacity
        Call SetLayeredWindowAttributes(hWnd, transpaColor, 0, LWA_COLORKEY)
        TRANSWIN_ERROR = "" ':: no error
        setWindowTransparent = True
    End If
End Function
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
