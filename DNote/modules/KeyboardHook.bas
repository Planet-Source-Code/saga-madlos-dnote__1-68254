Attribute VB_Name = "KeyboardHook"
Option Explicit
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'::                                    WIN 32 API
'::..................................................................................
Private Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Private Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Private Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal nCode As Long, ByVal wParam As Long, lParam As Any) As Long
':: get the state of the given keycode (key down or key up state)
':: for more info http://msdn.microsoft.com/library/default.asp?url=/library/en-us/winui/winui/windowsuserinterface/userinput/keyboardinput/keyboardinputreference/keyboardinputfunctions/getasynckeystate.asp
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'::                                     CONSTANTS
'::...................................................................................
Private Const HC_ACTION = 0 ':: one of the 2 expected value return by 'code' in
                            ':: KeyboardProc procedure. The wParam and lParam
                            ':: parameters contain information about a keystroke
                            ':: message.
Private Const WH_KEYBOARD_LL = 13 ':: hook type
Private Const TRANSITION_STATE = &H80   ':: (bit 7) Specifies the transition state
                                        ':: The value is 0 if the key is being
                                        ':: pressed and 1 if it is being released.
Private Const VK_LWIN As Byte = &H5B ':: Left Windows key (Microsoft Natural keyboard)
Private Const VK_RWIN As Byte = &H5C ':: Right Windows key (Microsoft Natural keyboard)
Private Const VK_SPACE = &H20 ':: Space bar
Private Const VK_LCONTROL = &HA2 ':: left control
Private Const VK_RCONTROL = &HA3 ':: right control
Private Const VK_MENU = &H12 ':: alt key
Private Const WM_KEYUP = &H101 ':: windows keyboard message
Private Const WM_KEYDOWN = &H100 ':: windows keyboard message
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'::                               STRUCTURES, ENUMS AND DATA TYPES
'::...................................................................................
':: Structure contains information about a low-level keyboard input event.
':: For more info visit page http://msdn.microsoft.com/library/default.asp?url=/library/en-us/winui/winui/windowsuserinterface/windowing/hooks/hookreference/hookstructures/kbdllhookstruct.asp
Type KBDLLHOOKSTRUCT
    vkCode      As Long
    ScanCode    As Long
    Flags       As Long
    Time        As Long
    DwExtraInfo As Long
End Type
':: Contains operating system version information. user with GetVersionEx()
Private Type OSVERSIONINFO
        dwOSVersionInfoSize As Long
        dwMajorVersion As Long
        dwMinorVersion As Long
        dwBuildNumber As Long
        dwPlatformId As Long
        szCSDVersion As String * 128      '  Maintenance string for PSS usage
End Type

':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'::                                 OTHER VARIABLES
'::...................................................................................
Private hKeyHook As Long    ':: handle of keyboard hook
Private HookState As Boolean ':: state of the keyboard hook. hook or unhook?
Public HookError As Integer ':: 0 - no error in setting or unsetting hook
                            ':: 1 - hook already set. no need to set it again.
                            ':: 2 - hook not set. there is nothing to unset
                            ':: 3 - Win32 API error
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
':: NAME        : LowLevelKeyboardProc
':: TYPE        : Function
':: SCOPE       : Private
':: PARAMETERS  : ncode As Long - [in] Specifies a code the hook procedure uses to
'::                             determine how to process the message.
'::               wParam As Long - [in]Specifies the identifier of the keyboard message
'::                                (WM_KEYDOWN, WM_KEYUP, WM_SYSKEYDOWN, or WM_SYSKEYUP)
'::               lParam As Long - [in] Pointer to a KBDLLHOOKSTRUCT structure.
':: RETURN      : Long
'::                 If the hook procedure processed the message, it may return a
'::                 nonzero value to prevent the system from passing the message
'::                 to the rest of the hook chain or the target window procedure.
':: DESCRIPTION : The system calls this function every time a new keyboard input event
'::                 is about to be posted into a thread input queue. This function is
'::                 used with Win32 API SetWindowsHookEx.
'::                 For detailed informatiom visit this page:
'::                     http://msdn.microsoft.com/library/default.asp?url=/library/en-us/winui/winui/windowsuserinterface/windowing/hooks/hookreference/hookfunctions/lowlevelkeyboardproc.asp
'::...................................................................................
Private Function LowLevelKeyboardProc(ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Dim keyInput As KBDLLHOOKSTRUCT
    If nCode < 0 Then
        ':: If code is less than zero, the hook procedure must return the value
        ':: returned by CallNextHookEx.
        LowLevelKeyboardProc = CallNextHookEx(hKeyHook, nCode, wParam, lParam)
    ElseIf nCode = HC_ACTION Then
        ':: copy the value pointed by lParam to a KBDLLHOOKSTRUCT strucure
        Call CopyMemory(keyInput, ByVal lParam, Len(keyInput))
        ':: this will prevent from echoing the space when the control key is pressed
        If (keyInput.vkCode = VK_SPACE And wParam = WM_KEYDOWN) And (CBool(GetAsyncKeyState(VK_LCONTROL) And &H8000) = True Or CBool(GetAsyncKeyState(VK_RCONTROL) And &H8000) = True) Then
            LowLevelKeyboardProc = 1
            Exit Function
        End If
        ':: if left LeftControl + Space  + Alt or RightControl + Space + Alt is pressed
        ':: activates only when the Space key is released
        If (keyInput.vkCode = VK_SPACE And wParam = WM_KEYUP) And (CBool(GetAsyncKeyState(VK_LCONTROL) And &H8000) = True Or CBool(GetAsyncKeyState(VK_RCONTROL) And &H8000) = True) And CBool(GetAsyncKeyState(VK_MENU) And &H8000) Then
            Call mdlMain.showNote ':: hide or show main form
            LowLevelKeyboardProc = 1    ':: prevent the system from passing the message
                                        ':: to the rest of the hook chain or the target
                                        ':: window procedure.
            Exit Function
        End If
        LowLevelKeyboardProc = CallNextHookEx(hKeyHook, nCode, wParam, lParam)
    End If
End Function
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
':: NAME        : hookKeyboard
':: TYPE        : Function
':: SCOPE       : Public
':: PARAMETERS  : none
':: RETURN      : Boolean
'::                 True - hooking of keyboard  succesfull
'::                 False - failure in hooking keyboard
':: DESCRIPTION : hook keyboard
'::...................................................................................
Public Function hookKeyboard() As Boolean
    If HookState = True Then    ':: check if keyboard already hooked
        HookError = 1           ':: if already hooked set error info. already hooked
        hookKeyboard = False    ':: then return false
    Else
        ':: hook keyboard
        hKeyHook = SetWindowsHookEx(WH_KEYBOARD_LL, AddressOf LowLevelKeyboardProc, App.hInstance, 0)
        If IsNull(hKeyHook) Then ':: Win32 API error
            HookState = False       ':: set hook state (not hooked)
            HookError = 3           ':: set error info. Win32 API error
            hookKeyboard = False    ':: return false. unsuccessful in hooking
        Else ':: successfull in hooking keyboard
            HookState = True        ':: set hook state (hooked)
            HookError = 0           ':: set error info. no error
            hookKeyboard = True     ':: return true. succesful in hooking keyboard
        End If
    End If
End Function
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
':: NAME        : unhookKeyboard
':: TYPE        : Function
':: SCOPE       : Public
':: PARAMETERS  : none
':: RETURN      : Boolean
'::                 True    - succesfull in unhooking keyboard
'::                 False   - failure in unhooking keyboard
':: DESCRIPTION : unhook keyboard
'::...................................................................................
Public Function unhookKeyboard() As Boolean
    If HookState = False Then   ':: if there is no keyboard hook to unhook
        HookError = 2           ':: then set error info. nothing to unhook
        unhookKeyboard = False  ':: then return false
    Else
        ':: unhook keyboard hook
        If UnhookWindowsHookEx(hKeyHook) <> 0 Then ':: unhook succesfull
            HookState = False                   ':: set hook state to false (no hook)
            HookError = 0                       ':: set error info. to no error
            unhookKeyboard = True               ':: return true (success in unhooking)
        Else ':: Win32 API error
            HookError = 3                       ':: set error info. Win32 API error
            unhookKeyboard = True               ':: return true (success in unhooking)
        End If
    End If
End Function
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
':: NAME        : isPlatformWin32NT
':: TYPE        : Function
':: SCOPE       : Public
':: PARAMETERS  : none
':: RETURN      : Boolean
'::                 True - if platform is base on Win32 NT
'::                 False - if platform in not base on Win32 NT
':: DESCRIPTION : check if the OS is base on Win32 NT or not
'::...................................................................................
Public Function isPlatformWin32NT() As Boolean
Dim osvi As OSVERSIONINFO
    osvi.dwOSVersionInfoSize = Len(osvi)
    Call GetVersionEx(osvi)
    If osvi.dwPlatformId = 2 Then  ':: 2 - Win32 NT base. 1 - Win98, Win95 and WinME
        isPlatformWin32NT = True
    Else
        isPlatformWin32NT = False
    End If
    
End Function
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
