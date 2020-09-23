VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TrayIcon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'::                                    WIN 32 API
'::..................................................................................
':: Sends a message to the taskbar's status area.
':: for more info http://msdn.microsoft.com/library/default.asp?url=/library/en-us/shellcc/platform/shell/reference/functions/shell_notifyicon.asp
Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pNID As NOTIFYICONDATA) As Boolean
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'::                                     CONSTANTS
'::...................................................................................
':: specifies the action to be taken by the funcion Shell_NotifyIcon
Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
'::..................................................................................
':: Flags that indicate which of the other members contain valid data in the
':: structure NOTIFYICONDATA
Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4
'::..................................................................................
Private Const WM_MOUSEMOVE = &H200 ':: call back message to be send to the owner form
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'::                               STRUCTURES, ENUMS AND DATA TYPES
'::...................................................................................
':: Contains information that the system needs to process taskbar status area messages
':: use with Shell_NotifyIcon() function
Private Type NOTIFYICONDATA
        cbSize As Long
        hWnd As Long
        uID As Long
        uFlags As Long
        uCallbackMessage As Long
        hIcon As Long
        szTip As String * 64
End Type
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'::                                 OTHER VARIABLES
'::...................................................................................
Private nid As NOTIFYICONDATA ':: instance of the structure use in Shell_NotifyIcon
Private onTray As Boolean ':: identify whether icon is on tray or not
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'::                           CLASS PROPERTIES
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
':: NAME        :
':: TYPE        :
':: RETURN      :
':: DESCRIPTION :
'::...................................................................................
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'::                         PROCEDURES (SUB AND FUNCTION)
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
':: NAME        : addTrayIcon
':: TYPE        : Function
':: SCOPE       : Public
':: PARAMETERS  : hwnd as Long - handle of the form that is the owner of the tool tip
'::               imgIcon  as Image - icon to be added in the tray icon
'::               toolTip as toolTip - tool tip text for the icon in tray icon
':: RETURN      : Boolean
'::                 True - successful in adding icon in tray icon
'::                 False - unsuccessfull in adding icon in tray icon
':: DESCRIPTION : adds icon in tray icon
'::...................................................................................
Public Function addTrayIcon(ByVal hWnd As Long, ByRef imgIcon As Image, toolTip As String) As Boolean
    If onTray Then ':: check if tray icon is already in tray
        addTrayIcon = False
    Else ':: if not yet on tray add icon
        ':: initialize nid (NOTIFYICONDATA)
        nid.cbSize = Len(nid)
        nid.hWnd = hWnd
        nid.uID = vbNull
        nid.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
        nid.uCallbackMessage = WM_MOUSEMOVE
        nid.hIcon = imgIcon.Picture
        nid.szTip = toolTip & vbNullChar
        onTray = True
        addTrayIcon = True
        Call Shell_NotifyIcon(NIM_ADD, nid) ':: add icon into tray icon
    End If
End Function
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
':: NAME        : modifyToolTip
':: TYPE        : Function
':: SCOPE       : Public
':: PARAMETERS  : toolTip As String - tool tip string
':: RETURN      : Boolean
'::                 True - successful in modifying tool tip in tray icon
'::                 False - unsuccessfull in modifying tool tip in tray icon
':: DESCRIPTION : modify the tool tip for the icon in tray icon
'::...................................................................................
Public Function modifyToolTip(toolTip As String) As Boolean
    If onTray Then ':: icon should be on tray icon before it can be modified
        nid.szTip = toolTip & vbNullChar ':: set new tool tip text
        Call Shell_NotifyIcon(NIM_MODIFY, nid) ':: modify tray icon
        modifyToolTip = True
    Else
        modifyToolTip = False
    End If
End Function
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
':: NAME        : modifTrayIcon
':: TYPE        : Function
':: SCOPE       : Public
':: PARAMETERS  : imgIcon as Image - icon to be displayed in the tray icon
':: RETURN      : Boolean
'::                 True - successful in modifying icon in tray icon
'::                 False - unsuccessfull in modifying icon in tray icon
':: DESCRIPTION : modify the icon displayed in the tray icon
'::...................................................................................
Public Function modifyTrayIcon(ByRef imgIcon As Image) As Boolean
    If onTray Then ':: icon should be on tray icon before it can be modified
        nid.hIcon = imgIcon.Picture ':: set new icon
        Call Shell_NotifyIcon(NIM_MODIFY, nid) ':: modify tray icon
        modifyTrayIcon = True
    Else
        modifyTrayIcon = False
    End If
End Function
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
':: NAME        : removeTrayIcon
':: TYPE        : Function
':: SCOPE       : Public
':: PARAMETERS  : none
':: RETURN      : Boolean
'::                 True - successful in removing icon in tray icon
'::                 False - unsuccessfull in removing icon in tray icon
':: DESCRIPTION : remove icon in the tray icon
'::...................................................................................
Public Function removeTrayIcon() As Boolean
    If onTray Then ':: check if the icon is in tray icon
        Call Shell_NotifyIcon(NIM_DELETE, nid) ':: remove icon in tray icon
        onTray = False
        removeTrayIcon = True
    Else
        removeTrayIcon = False
    End If
End Function
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
':: NAME        : Class_Initialize
':: TYPE        : Sub
':: SCOPE       : Private
':: PARAMETERS  : none
':: RETURN      : none
':: DESCRIPTION : class constructor
'::...................................................................................
Private Sub Class_Initialize()
    onTray = False
End Sub
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
':: NAME        : Class_Terminate
':: TYPE        : Sub
':: SCOPE       : Private
':: PARAMETERS  : none
':: RETURN      : none
':: DESCRIPTION : class destructor
'::...................................................................................
Private Sub Class_Terminate()
    Call removeTrayIcon
End Sub
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

