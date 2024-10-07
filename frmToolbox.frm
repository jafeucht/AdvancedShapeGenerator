VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmToolbox 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Tools"
   ClientHeight    =   8415
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   1455
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8415
   ScaleWidth      =   1455
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdRandom 
      Caption         =   "&Random"
      Height          =   375
      Left            =   75
      TabIndex        =   26
      Top             =   7950
      Width           =   1290
   End
   Begin VB.Frame fmColors 
      Caption         =   "Colors"
      Height          =   2475
      Left            =   60
      TabIndex        =   17
      Top             =   5400
      Width           =   1335
      Begin VB.OptionButton optColor 
         Caption         =   "Gradient"
         Height          =   255
         Index           =   1
         Left            =   75
         TabIndex        =   23
         Top             =   1575
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.OptionButton optColor 
         Caption         =   "Solid"
         Height          =   255
         Index           =   0
         Left            =   75
         TabIndex        =   20
         Top             =   750
         Width           =   1095
      End
      Begin VB.PictureBox pctBackColor 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   75
         ScaleHeight     =   225
         ScaleWidth      =   1140
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   450
         Width           =   1170
      End
      Begin VB.PictureBox pctGradColor 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   75
         ScaleHeight     =   225
         ScaleWidth      =   1140
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   2100
         Width           =   1170
      End
      Begin VB.PictureBox pctColor 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   75
         ScaleHeight     =   225
         ScaleWidth      =   1140
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   1275
         Width           =   1170
      End
      Begin VB.Label lblGradColor 
         Caption         =   "Gradient"
         Height          =   255
         Left            =   75
         TabIndex        =   24
         Top             =   1875
         Width           =   735
      End
      Begin VB.Label lblColor 
         Caption         =   "Main"
         Height          =   255
         Left            =   75
         TabIndex        =   21
         Top             =   1050
         Width           =   495
      End
      Begin VB.Label lblBackColor 
         Caption         =   "Background"
         Height          =   255
         Left            =   75
         TabIndex        =   18
         Top             =   225
         Width           =   1125
      End
   End
   Begin VB.TextBox txtRotate 
      Height          =   285
      Left            =   75
      TabIndex        =   16
      Top             =   5025
      Width           =   1290
   End
   Begin VB.TextBox txtZoom 
      Height          =   285
      Left            =   75
      TabIndex        =   14
      Top             =   4425
      Width           =   1290
   End
   Begin VB.TextBox txtLineWidth 
      Height          =   285
      Left            =   75
      TabIndex        =   12
      Top             =   3825
      Width           =   1290
   End
   Begin VB.Frame fmVectors 
      Caption         =   "Vectors"
      Height          =   2865
      Left            =   60
      TabIndex        =   2
      Top             =   675
      Width           =   1335
      Begin MSComDlg.CommonDialog CDialog 
         Left            =   120
         Top             =   120
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton cmdVectorStr 
         Caption         =   "&Vector String"
         Height          =   375
         Left            =   75
         TabIndex        =   10
         Top             =   2400
         Width           =   1170
      End
      Begin VB.TextBox txtAngle 
         Enabled         =   0   'False
         Height          =   285
         Left            =   75
         TabIndex        =   9
         Top             =   2025
         Width           =   1170
      End
      Begin VB.TextBox txtLength 
         Enabled         =   0   'False
         Height          =   285
         Left            =   75
         TabIndex        =   7
         Top             =   1425
         Width           =   1170
      End
      Begin VB.ComboBox cboVectors 
         Height          =   315
         Left            =   75
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   825
         Width           =   1170
      End
      Begin VB.TextBox txtVectorCnt 
         Height          =   285
         Left            =   75
         TabIndex        =   4
         Top             =   450
         Width           =   1170
      End
      Begin VB.Label lblAngle 
         Caption         =   "Angle:"
         Enabled         =   0   'False
         Height          =   255
         Left            =   75
         TabIndex        =   8
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label lblLength 
         Caption         =   "Length:"
         Enabled         =   0   'False
         Height          =   255
         Left            =   75
         TabIndex        =   6
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label lblVectors 
         AutoSize        =   -1  'True
         Caption         =   "Vector Count"
         Height          =   195
         Left            =   75
         TabIndex        =   3
         Top             =   225
         Width           =   930
      End
   End
   Begin VB.CheckBox chkGrid 
      Caption         =   "Grid"
      Height          =   255
      Left            =   75
      TabIndex        =   1
      Top             =   375
      Width           =   1290
   End
   Begin VB.CheckBox chkAxis 
      Caption         =   "X/Y Axis"
      Height          =   255
      Left            =   75
      TabIndex        =   0
      Top             =   75
      Width           =   1290
   End
   Begin VB.Label lblRotate 
      AutoSize        =   -1  'True
      Caption         =   "Rotate:"
      Height          =   195
      Left            =   75
      TabIndex        =   15
      Top             =   4800
      Width           =   525
   End
   Begin VB.Label lblZoom 
      AutoSize        =   -1  'True
      Caption         =   "Zoom:"
      Height          =   195
      Left            =   75
      TabIndex        =   13
      Top             =   4200
      Width           =   450
   End
   Begin VB.Label lblLineWidth 
      AutoSize        =   -1  'True
      Caption         =   "Line Width:"
      Height          =   195
      Left            =   75
      TabIndex        =   11
      Top             =   3600
      Width           =   810
   End
End
Attribute VB_Name = "frmToolbox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' frmToolbox
' This form allows the user to alter the properties of a selected
' shape.

Option Explicit

' Enable/Disable all controls on the toolbox
Public Sub Activate(Optional Enabled As Boolean = True)
Dim i As Integer
    On Error Resume Next
    For i = 0 To Controls.Count
        Controls(i).Enabled = Enabled
    Next i
End Sub

' Display information on the selected vector.
Private Sub cboVectors_Click()
Dim Length As Double, Angle As Double
    SelWindow.Gadget.GetVector cboVectors.ListIndex + 1, Length, Angle
    txtLength = Length
    txtAngle = Angle & "°"
End Sub

Private Sub chkAxis_Click()
    SelWindow.ShowAxis = chkAxis.Value
End Sub

Private Sub chkGrid_Click()
    SelWindow.ShowGrid = chkGrid.Value
End Sub

Function IsNothing(TestObj As Object) As Boolean
    On Error Resume Next
    ' Returns True if error
    If TestObj.Name = "" Then IsNothing = True
End Function

' Creates random shape and colors
Private Sub cmdRandom_Click()
Dim i As Integer
    Randomize Timer
    With SelWindow
        With .Gadget
            .ClearVectors
            .SetVectorCnt (Int(Rnd * 10) + 1)
            For i = 1 To .VectorCnt
                .SetVector i, Int(Rnd * 60) + 5, Int(Rnd * 360)
            Next i
        End With
        .GradColor = RGB(Int(Rnd * 156) + 100, Int(Rnd * 156) + 100, Int(Rnd * 156) + 100)
        .BaseColor = RGB(Int(Rnd * 156) + 100, Int(Rnd * 156) + 100, Int(Rnd * 156) + 100)
    End With
    GetSettings
End Sub

' Prompt for vector code.
Private Sub cmdVectorStr_Click()
Dim RetStr As String
    RetStr = InputBox("Type in your vector code here. Separate vectors by colons ("":""), and separate lengths from angles with a less-than sign (""<"").", , SelWindow.Gadget.CreateVectorStr)
    If RetStr = vbNullString Then Exit Sub
    SelWindow.Gadget.TakeVectorStr RetStr
    LoadVectors
End Sub

Private Sub Form_Load()
    frmMDI.ChangeToolBoxState True
    ' If no window is open, deactivate toolbox.
    If IsNothing(SelWindow) Then
        Activate False
    Else
        GetSettings
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmMDI.ChangeToolBoxState False
End Sub

Private Sub optColor_Click(Index As Integer)
Dim GradMode As Boolean
    GradMode = (Index = 1)
    SelWindow.Gradient = GradMode
    pctGradColor.Enabled = GradMode
    lblGradColor.Enabled = GradMode
End Sub

' Load all the shape's vectors to the cboVectors dropdown list.
Sub LoadVectors()
Dim i As Integer, OldIdx As Integer
    On Error Resume Next
    OldIdx = cboVectors.ListIndex
    If OldIdx = -1 Then OldIdx = 0
    cboVectors.Clear
    txtVectorCnt = SelWindow.Gadget.VectorCnt
    For i = 1 To SelWindow.Gadget.VectorCnt
        cboVectors.AddItem i
    Next i
    If cboVectors.ListIndex = -1 Then cboVectors.ListIndex = 0
    If cboVectors.ListCount > OldIdx Then cboVectors.ListIndex = OldIdx
End Sub

' Check a textbox if it is numeric and in bounds
Function Validate(StrVal As String, Min As Double, Max As Double) As Boolean
Const IGNORE = "%°"
    StrVal = Trim$(StrVal)
    If InStr(1, IGNORE, Right(StrVal, 1)) > 0 Then StrVal = Left(StrVal, Len(StrVal) - 1)
    Validate = IsNumeric(StrVal) And Val(StrVal) >= Min And Val(StrVal) <= Max
End Function

Public Sub GetSettings()
    ' Load all the settings from the selected instance of frmTest
    With SelWindow
        chkGrid.Value = -.ShowGrid
        chkAxis.Value = -.ShowAxis
        LoadVectors
        txtLineWidth = .LineWidth
        txtZoom = .Gadget.Zoom & "%"
        txtRotate = .Gadget.Rotate & "°"
        optColor(Abs(.Gradient)).Value = True
        pctBackColor.BackColor = .BkColor
        pctColor.BackColor = .BaseColor
        pctGradColor.BackColor = .GradColor
    End With
End Sub

Private Sub pctBackColor_Click()
    ShowColors pctBackColor
    SelWindow.BkColor = pctBackColor.BackColor
End Sub

Private Sub pctColor_Click()
    ShowColors pctColor
    SelWindow.BaseColor = pctColor.BackColor
End Sub

' Prompts a color dialog box for a selected PictureBox control.
Sub ShowColors(PBox As PictureBox)
    CDialog.Color = PBox.BackColor
    CDialog.ShowColor
    PBox.BackColor = CDialog.Color
End Sub

Private Sub pctGradColor_Click()
    ShowColors pctGradColor
    SelWindow.GradColor = pctGradColor.BackColor
End Sub

Private Sub txtAngle_Change()
Dim ValidNum As Boolean
    On Error Resume Next
    If cboVectors.ListIndex = -1 Then Exit Sub
    ValidNum = Validate(txtAngle, 0, 32766)
    txtAngle.ForeColor = 255 * (ValidNum + 1)
    If Not ValidNum Then Exit Sub
    SelWindow.Gadget.SetVector cboVectors.ListIndex + 1, txtLength, Val(txtAngle)
End Sub

Private Sub txtAngle_GotFocus()
    SetTBoxFocus txtAngle
End Sub

Private Sub txtAngle_KeyPress(KeyAscii As Integer)
    If cboVectors.ListIndex = -1 Then Beep: KeyAscii = 0
End Sub

Private Sub txtAngle_LostFocus()
    FixAngle txtAngle
End Sub

Private Sub txtLength_Change()
Dim ValidNum As Boolean
    ValidNum = Validate(txtLength, 0, 32766)
    txtLength.ForeColor = 255 * (ValidNum + 1)
    If Not ValidNum Then Exit Sub
    SelWindow.Gadget.SetVector cboVectors.ListIndex + 1, Val(txtLength), Val(txtAngle)
End Sub

Private Sub txtLength_GotFocus()
    SetTBoxFocus txtLength
End Sub

Private Sub txtLength_KeyPress(KeyAscii As Integer)
    If cboVectors.ListIndex = -1 Then Beep: KeyAscii = 0
End Sub

Private Sub txtLineWidth_Change()
Dim ValidNum As Boolean
    ValidNum = Validate(txtLineWidth, 0, 32766)
    txtLineWidth.ForeColor = 255 * (ValidNum + 1)
    If Not ValidNum Then Exit Sub
    If Not Validate(txtLineWidth, 1, 60) Then Exit Sub
    SelWindow.LineWidth = txtLineWidth
End Sub

Private Sub txtLineWidth_GotFocus()
    SetTBoxFocus txtLineWidth
End Sub

Private Sub txtRotate_GotFocus()
    SetTBoxFocus txtRotate
End Sub

Private Sub txtRotate_LostFocus()
    FixAngle txtRotate
End Sub

Private Sub txtVectorCnt_Change()
Dim ValidNum As Boolean
    ValidNum = Validate(txtVectorCnt, 0, 1000)
    txtVectorCnt.ForeColor = 255 * (ValidNum + 1)
End Sub

Private Sub txtVectorCnt_GotFocus()
    SetTBoxFocus txtVectorCnt
End Sub

Private Sub txtZoom_Change()
Dim ValidNum As Boolean, txtStr As String, i As Integer
    If txtZoom = "" Then Exit Sub
    i = InStr(1, txtZoom, "%")
    If i = 0 Then i = Len(txtZoom) + 1
    txtStr = Left(txtZoom, i - 1)
    ValidNum = Validate(txtStr, 0, 32766)
    txtZoom.ForeColor = 255 * (ValidNum + 1)
    If Not ValidNum Then Exit Sub
    SelWindow.Gadget.Zoom = Val(txtZoom)
End Sub

' Fix an angle so that it is from 0 to 359
Sub FixAngle(ByRef TBox As TextBox)
Dim Step As Integer, TNum As Integer
    On Error Resume Next
    TNum = Val(TBox)
    If Val(TNum) > 0 Then
        Step = -Abs(TNum) / TNum
        Do While TNum >= 360 Or TNum < 0
            TNum = TNum + Step * 360
        Loop
    End If
    TBox = TNum & "°"
End Sub

Private Sub txtRotate_Change()
Dim ValidNum As Boolean
    ValidNum = Validate(txtRotate, 0, 32766)
    txtRotate.ForeColor = 255 * (ValidNum + 1)
    If Not ValidNum Then Exit Sub
    SelWindow.Gadget.Rotate = Val(txtRotate)
End Sub

Private Sub SetTBoxFocus(TBox As TextBox)
    TBox.SelStart = 0
    TBox.SelLength = Len(TBox)
End Sub

Private Sub txtVectorCnt_LostFocus()
    If Not Validate(txtVectorCnt, 0, 1000) Then Exit Sub
    'txtAngle.Enabled = txtVectorCnt > 0
    'txtLength.Enabled = txtVectorCnt > 0
    SelWindow.Gadget.SetVectorCnt txtVectorCnt
    LoadVectors
End Sub

Private Sub txtZoom_GotFocus()
    SetTBoxFocus txtZoom
End Sub

Private Sub txtZoom_LostFocus()
    On Error Resume Next
    ' Add a percent sign onto txtZoom
    txtZoom = Val(txtZoom) & "%"
End Sub
