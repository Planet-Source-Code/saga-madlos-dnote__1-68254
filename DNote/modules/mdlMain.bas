Attribute VB_Name = "mdlMain"
Option Explicit
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'::                                    WIN 32 API
'::..................................................................................
':: function releases the mouse capture from a window in the current thread and restores normal mouse input processing
':: for more info http://msdn2.microsoft.com/en-us/library/ms646261.aspx
Public Declare Function ReleaseCapture Lib "user32.dll" () As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function SetActiveWindow Lib "user32" (ByVal hWnd As Long) As Long
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'::                                     CONSTANTS
'::...................................................................................
Public Const CF_TEXT = 1 ':: Clip board date format
Public Const HWND_TOPMOST = -1 ':: as top most window
Public Const SWP_NOSIZE = &H1 ':: do not resize form
Public Const SWP_NOMOVE = &H2 ':: do not move form
Public Const WM_NCLBUTTONDOWN = &HA1 '::
Public Const HTCAPTION = 2 '::
Public Const HKEY_LOCAL_MACHINE = &H80000002 ':: registry key handle
Private Const REG_OPENED_EXISTING_KEY = &H2 ':: Existing Key opened
Private Const REG_SZ = 1    ':: Unicode nul terminated string
Private Const subKey = "SOFTWARE\MICROSOFT\WINDOWS\CURRENTVERSION\RUN" ':: sub key
Private Const valueName = "DNote" ':: value name in registry
Private regData As String ':: data to be save in registry
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'::                                 OTHER VARIABLES
'::...................................................................................
Public Const WM_RBUTTONUP = &H205 ':: right mouse button up
Public Const WM_LBUTTONUP = &H202 ':: left mouse button up
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'::                               STRUCTURES, ENUMS AND DATA TYPES
'::...................................................................................
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'::                         PROCEDURES (SUB AND FUNCTION)
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
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
':: NAME        : Main
':: TYPE        : Sub
':: SCOPE       : Global
':: PARAMETERS  : none
':: RETURN      : none
':: DESCRIPTION : program entry point
'::...................................................................................
Sub Main()
    App.TaskVisible = False ':: to hide the process in task manager
    If App.PrevInstance Then End ':: exit if application is already open
    regData = App.Path & "\" & App.EXEName & ".exe" ':: application path and name
    ':: since the low level keyboard hook does not work with non NT base version of
    ':: Windows (Win 98, 95) check first if it is NT base or not
    If isPlatformWin32NT = True Then
        Call setRunAtStartUp ':: set program to run at start up (registry)
        If KeyboardHook.hookKeyboard = False Then ':: hook keyboard
            MsgBox "Error in hooking keyboard. Program closing", vbCritical, "DNote - Saga"
        Else
            Load frmMain
            Load frmMenu
        End If
    Else
        MsgBox "Your version of windows does not support some functionality of this program. Program closing.", vbCritical, "DNote - Saga"
    End If
End Sub
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
':: NAME        : showNote
':: TYPE        : Sub
':: SCOPE       : Public
':: PARAMETERS  : none
':: RETURN      : none
':: DESCRIPTION : this is the function/procedure the system will call everytime the
'::                 PrintScreen key is pressed
'::...................................................................................
Public Sub showNote()
    If frmAbout.Visible = True Then Exit Sub
    If frmMain.Visible = True Then
        frmMain.Visible = False
        ':: save text to dnote.txt so that when computer or the program is closed
        ':: and then run again the last text in the text box will be displayed again
        Call saveToDisk
    Else
        frmMain.Visible = True
        ':: set the window as top most
        Call SetWindowPos(frmMain.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE)
        Call SetActiveWindow(frmMain.hWnd) ':: activate window
        frmMain.txtNote.SetFocus ':: set focus on text box
        ':: load text save in dnote.txt to the text box
        Call loadFromDisk
    End If
End Sub
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
':: NAME        : terminateApplication
':: TYPE        : Sub
':: SCOPE       : Public
':: PARAMETERS  : none
':: RETURN      : none
':: DESCRIPTION : function use to terminate process
'::...................................................................................
Public Sub terminateApplication()
    frmMenu.sysTrayIcon.removeTrayIcon
    KeyboardHook.unhookKeyboard
    Unload frmMain
    Unload frmMenu
End Sub
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
':: NAME        : setRunStartUp
':: TYPE        : Function
':: SCOPE       : Private
':: PARAMETERS  : none
':: RETURN      : none
':: DESCRIPTION : add DNote program to the list of programs to run during start up
'::...................................................................................
Private Function setRunAtStartUp()
Dim hKey As Long
    Call RegCreateKey(HKEY_LOCAL_MACHINE, subKey, hKey)
    Call RegSetValueEx(hKey, valueName, 0, REG_SZ, ByVal regData, Len(regData))
    Call RegCloseKey(hKey)
End Function
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
':: NAME        : saveToDisk
':: TYPE        : Function
':: SCOPE       : Private
':: PARAMETERS  : none
':: RETURN      : none
':: DESCRIPTION : save the text in the frmMain text box to the disk (dnote.txt)
'::...................................................................................
Private Function saveToDisk()
Dim fso As New FileSystemObject
Dim dnoteF As File
Dim txtStream As TextStream
    ':: create dnote.txt. If it already exist data are truncated
    Call fso.CreateTextFile(App.Path & "\dnote.txt")
    Set dnoteF = fso.GetFile(App.Path & "\dnote.txt") ':: get the file
    Set txtStream = dnoteF.OpenAsTextStream(ForWriting) ':: open a text stream for wrting
    Call txtStream.Write(Trim(frmMain.txtNote.Text)) ':: write text to dnote.txt
    txtStream.Close
End Function
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
':: NAME        : loadFromDisk
':: TYPE        : Function
':: SCOPE       : Private
':: PARAMETERS  : none
':: RETURN      : none
':: DESCRIPTION : load text to text box in frmMain from dnote.txt
'::...................................................................................
Private Function loadFromDisk()
On Error Resume Next
Dim fso As New FileSystemObject
Dim dnoteF As File
Dim txtStream As TextStream
    Set dnoteF = fso.GetFile(App.Path & "\dnote.txt") ':: get the file
    Set txtStream = dnoteF.OpenAsTextStream(ForReading) ':: open a text stream for reading
    frmMain.txtNote.Text = Trim(txtStream.ReadAll) ':: write text from dnote.txt to txtbox
    txtStream.Close
End Function
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
