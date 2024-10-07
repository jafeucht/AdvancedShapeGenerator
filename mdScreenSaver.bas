Attribute VB_Name = "mdScreenSaver"
Option Explicit

Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function PwdChangePassword& Lib "mpr" Alias "PwdChangePasswordA" (ByVal lpcRegkeyname$, ByVal hwnd&, ByVal uiReserved1&, ByVal uiReserved2&)
Public Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Public Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long

Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Const HWND_TOP = 0
Public Const GWL_HWNDPARENT = (-8)
Public Const GWL_STYLE = (-16)
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_NOZORDER = &H4
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const SWP_SHOWWINDOW = &H40
Public Const WS_CHILD = &H40000000
Public Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2

Enum RunModes
    ScreenSaver = 1
    Configure
    Preview
    Password
End Enum

Global RunMode As RunModes
Const ModeStr = "SCPA"

Const APP_NAME = "Shape Screen Saver"
Const APP_SECTION = "Feuchtersoft"
Const SCT_LINEWIDTH = "Line Width"
Const SCT_PRINTNUM = "Print Quantity"
Const SCT_PAUSETIME = "Pause Length"
Const SCT_PHASESEP = "Phase Separation"

Public PreviewWindow As Long

Public LineWidth As Integer
Public PrintNum As Long
Public PauseTime As Double
Public PhaseSep As Double

Sub LoadSettings()
    LineWidth = Val(GetSetting(APP_SECTION, APP_NAME, SCT_LINEWIDTH, "1"))
    PrintNum = Val(GetSetting(APP_SECTION, APP_NAME, SCT_PRINTNUM, "10"))
    PauseTime = Val(GetSetting(APP_SECTION, APP_NAME, SCT_PAUSETIME, ".5"))
    PhaseSep = Val(GetSetting(APP_SECTION, APP_NAME, SCT_PHASESEP, "2"))
End Sub

Sub SaveSettings()
    SaveSetting APP_SECTION, APP_NAME, SCT_LINEWIDTH, LineWidth
    SaveSetting APP_SECTION, APP_NAME, SCT_PRINTNUM, PrintNum
    SaveSetting APP_SECTION, APP_NAME, SCT_PAUSETIME, PauseTime
    SaveSetting APP_SECTION, APP_NAME, SCT_PHASESEP, PhaseSep
End Sub

Public Sub Main()
Dim PreviewRect As RECT
Dim WindowStyle As Long
Dim rc As Long
    
    rc = InStr(1, ModeStr, Right(Left(UCase(Trim(Command)) & "  ", 2), 1))
    If rc = 0 Then rc = 2
    RunMode = rc
    
    If RunMode = Password Or RunMode = Preview Then
        PreviewWindow = CLng(Right(Command, Len(Command) - 3))
    End If
    
    LoadSettings
    
    With frmMain
        Select Case RunMode
            Case Configure
                'Configure
                RunMode = Configure
                frmConfig.Show
            Case Preview
                'Create a preview screen
                GetClientRect PreviewWindow, PreviewRect
                Load frmMain
                WindowStyle = GetWindowLong(.hwnd, GWL_STYLE)
                WindowStyle = (WindowStyle Or WS_CHILD)
                SetWindowLong .hwnd, GWL_STYLE, WindowStyle
                SetParent .hwnd, PreviewWindow
                SetWindowLong .hwnd, GWL_HWNDPARENT, PreviewWindow
                SetWindowPos .hwnd, HWND_TOP, 0&, 0&, PreviewRect.Right, PreviewRect.Bottom, SWP_NOZORDER Or SWP_NOACTIVATE Or SWP_SHOWWINDOW
            Case Password
                'Change the screensaver password
                On Error GoTo Error
                PwdChangePassword "SCRSAVE", PreviewWindow, 0, 0
            Case ScreenSaver
                'Run the screensaver
                CheckShouldRun
                .Show
        End Select
    End With
    Exit Sub

Error:
    MsgBox "Password could not be changed", vbOKOnly
    End
End Sub

Public Function UsePassword() As Boolean
'Check wether a password has been used or not
Dim lHandle As Long
Dim lResult As Long
Dim lValue As Long
    UsePassword = False
    lResult = RegOpenKeyEx(&H80000001, "Control Panel\Desktop", 0, 1, lHandle)
    If lResult = 0 Then
        lResult = RegQueryValueEx(lHandle, "ScreenSaveUsePassword", 0, 4, lValue, 32)
        If lResult = 0 Then
            UsePassword = lValue
            lResult = RegCloseKey(lHandle)
        End If
    End If
End Function

Public Sub ShowCurs(Optional Show As Boolean = True)
Dim sCursor As Long, sShow As Integer
    sShow = -Show * 2 - 1
    Do
        sCursor = ShowCursor(Show)
    Loop Until Abs(sCursor) * sShow = sCursor And sCursor <> 0
End Sub

Private Sub CheckShouldRun()
    If Not App.PrevInstance Then Exit Sub
    If FindWindow(vbNullString, APP_NAME) Then End
    frmMain.Caption = APP_NAME
End Sub
