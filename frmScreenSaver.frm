VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H80000007&
   BorderStyle     =   0  'None
   ClientHeight    =   4785
   ClientLeft      =   1695
   ClientTop       =   2025
   ClientWidth     =   4500
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   FillStyle       =   0  'Solid
   Icon            =   "frmScreenSaver.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   319
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   300
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Timer tmTimer 
      Interval        =   1
      Left            =   3975
      Top             =   75
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function VerifyScreenSavePwd Lib "password.cpl" (ByVal hwnd&) As Boolean

Dim WithEvents Gadget As csShape
Attribute Gadget.VB_VarHelpID = -1
Dim CurColor As cColor
Dim GradColor As cColor
Dim LineCnt As Integer

Const MIN_COLOR = &H64&
Const MAX_ACCEL = 3
Const MAX_VEL = 6

Private Type PointAPI
    X As Double
    Y As Double
End Type

'------------------------------------------------------
' Used in color functions. Contains R, G, and B values
' for a color.
'------------------------------------------------------
Private Type cColor
    R As Integer
    G As Integer
    B As Integer
End Type

Private Function SetPnt(X As Double, Y As Double) As PointAPI
    ' Allow PointAPI data change in one command
    With SetPnt
        .X = X
        .Y = Y
    End With
End Function

'----------------------------------------
' Waits for a given number of seconds.
' Warning: Do not execute over midnight.
'----------------------------------------
Private Sub Wait(ByVal Seconds As Currency, Optional ByRef Start As Double, Optional ByRef TotalTime As Double, Optional ByRef Finish As Double)
    Start = Timer
    Do While Timer < Start + Seconds
        DoEvents
    Loop
    Finish = Timer
    TotalTime = Finish - Start
End Sub

'--------------------------------------------------
' Color Function, converts a long color value into
' a cColor Type
' -------------------------------------------------
Private Function ColorRGB(ByVal LongColor As Long) As cColor
    ColorRGB.R = LongColor Mod 256
    ColorRGB.G = (LongColor And &HFF00FF00) / 256
    ColorRGB.B = (LongColor And &HFF0000) / 256 ^ 2
End Function

'-------------------------------------------------
' Color Function, finds a color value Fraction *
' the distance between two colors.
'-------------------------------------------------
Private Function ColorBetween(StartCol As cColor, EndCol As cColor, Fraction As Currency) As Long
Dim ResultColor As cColor
    ResultColor.R = Fraction * (StartCol.R - EndCol.R) + EndCol.R
    ResultColor.G = Fraction * (StartCol.G - EndCol.G) + EndCol.G
    ResultColor.B = Fraction * (StartCol.B - EndCol.B) + EndCol.B
    ColorBetween = RGB(ResultColor.R, ResultColor.G, ResultColor.B)
End Function

Private Sub Animate()
Static ColVelocity As cColor, ColAccel As cColor
Static GradVelocity As cColor, GradAccel As cColor
Dim i As Integer

    tmTimer.Enabled = False
    
    RandomElement CurColor.R, ColVelocity.R, ColAccel.R
    RandomElement CurColor.G, ColVelocity.G, ColAccel.G
    RandomElement CurColor.B, ColVelocity.B, ColAccel.B
    RandomElement GradColor.R, GradVelocity.R, GradAccel.R
    RandomElement GradColor.G, GradVelocity.G, GradAccel.G
    RandomElement GradColor.B, GradVelocity.B, GradAccel.B
    
    Gadget.Left = Int(Rnd * ScaleWidth / 2 + ScaleWidth / 4)
    Gadget.Top = Int(Rnd * ScaleHeight / 2 + ScaleWidth / 4)
    
    Gadget.ClearVectors
    For i = 1 To Int(Rnd * 19) + 1
        Gadget.AddVector i, Int(Rnd * ScaleWidth / 10) + 5, Int(Rnd * 360)
    Next i
    
    LineCnt = Gadget.NumSequences * Gadget.VectorCnt
    
    Gadget.DrawShape
    Wait PauseTime
    
    tmTimer.Enabled = True
End Sub

Function GetSign(Number As Integer) As Integer
    GetSign = Abs(Number) / Number
End Function

Function RandomElement(ByRef Element As Integer, ByRef Velocity As Integer, ByRef Acceleration As Integer) As Integer
    Acceleration = Acceleration - (Int(Rnd * 2) * 2 - 1)
    If Acceleration > MAX_ACCEL Then Acceleration = GetSign(Acceleration) * MAX_ACCEL
    Velocity = Velocity + Acceleration
    If Velocity > MAX_VEL Then
        Velocity = GetSign(Velocity) * MAX_VEL
        Acceleration = -Acceleration
    End If
    Element = Element + Velocity
    If Element > 255 Then
        Element = 255
        Velocity = 0
        Acceleration = 0
    ElseIf Element < MIN_COLOR Then
        Element = MIN_COLOR
        Velocity = 0
        Acceleration = 0
    End If
End Function

Private Sub Form_KeyPress(KeyAscii As Integer)
    If RunMode = ScreenSaver Then
        Unload Me
    End If
End Sub
Private Sub Form_Load()

    If RunMode <> Preview Then
        If UsePassword Then Call SystemParametersInfo(97, 1, 0, 0)
        ShowCurs False
        SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
    End If
    
    Set Gadget = New csShape
    Set Gadget.Field = Me
    
    DrawWidth = LineWidth
    
    CurColor.R = MIN_COLOR
    CurColor.G = MIN_COLOR
    CurColor.B = MIN_COLOR
    
    Randomize Timer

End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If RunMode = ScreenSaver Then
        Unload Me
        End
    End If
    
End Sub

Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Static Count As Integer
    Count = Count + 1 ' Give enough time for program to run
    
    If Count > 5 Then
        If RunMode = ScreenSaver Then
            Unload Me
        End If
    End If
    
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Static HitNum As Integer
Cancel = False

    HitNum = HitNum + 1
    If HitNum < 4 Then Cancel = True: Exit Sub
    HitNum = 0
    
    'If Windows is shut down close this application too
    If UnloadMode = vbAppWindows Then
        Exit Sub
    End If
    
    ShowCursor True
    
    'if a password is beeing used ask for it and check its validity
    If RunMode = ScreenSaver And UsePassword Then
        If (VerifyScreenSavePwd(Me.hwnd)) = False Then
            Cancel = True
        End If
    End If

    If Not Cancel Then End
    ShowCursor False
End Sub

Private Sub Gadget_DrawLine(LineNum As Long, Color As Long)
    DoEvents
    Color = ColorBetween(CurColor, GradColor, LineNum / LineCnt)
End Sub

Private Sub tmTimer_Timer()
Static PhaseCnt As Long
    PhaseCnt = PhaseCnt + 1
    Animate
    If PhaseCnt >= PrintNum Then
        Wait PhaseSep - PauseTime
        Cls
        PhaseCnt = 0
    End If
End Sub
