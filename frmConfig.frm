VERSION 5.00
Begin VB.Form frmConfig 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Screen Saver Setup"
   ClientHeight    =   3135
   ClientLeft      =   225
   ClientTop       =   1530
   ClientWidth     =   4575
   Icon            =   "frmConfig.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   4575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtPhaseSep 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   3000
      MaxLength       =   5
      TabIndex        =   11
      Top             =   1320
      Width           =   615
   End
   Begin VB.TextBox txtPauseTime 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   3000
      MaxLength       =   4
      TabIndex        =   9
      Top             =   840
      Width           =   615
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3240
      TabIndex        =   7
      Top             =   2520
      Width           =   1095
   End
   Begin VB.TextBox txtLineWidth 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   3000
      MaxLength       =   2
      TabIndex        =   5
      Top             =   1800
      Width           =   615
   End
   Begin VB.CommandButton cmdOkay 
      Caption         =   "&Okay"
      Default         =   -1  'True
      Height          =   375
      Left            =   2040
      TabIndex        =   6
      Top             =   2520
      Width           =   1095
   End
   Begin VB.TextBox txtPrintNum 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   3000
      MaxLength       =   4
      TabIndex        =   1
      Top             =   360
      Width           =   615
   End
   Begin VB.CommandButton cmdAbout 
      Caption         =   "&About"
      Height          =   375
      Left            =   240
      TabIndex        =   8
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      Caption         =   "points"
      Height          =   195
      Index           =   2
      Left            =   3840
      TabIndex        =   13
      Top             =   1800
      Width           =   420
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      Caption         =   "seconds"
      Height          =   195
      Index           =   1
      Left            =   3840
      TabIndex        =   12
      Top             =   1320
      Width           =   600
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      Caption         =   "seconds"
      Height          =   195
      Index           =   0
      Left            =   3840
      TabIndex        =   10
      Top             =   840
      Width           =   600
   End
   Begin VB.Label lblLineWidth 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Width of each line:"
      Height          =   195
      Left            =   1440
      TabIndex        =   4
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Label lblPhaseSep 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Seconds to pause between screens:"
      Height          =   195
      Left            =   180
      TabIndex        =   3
      Top             =   1320
      Width           =   2595
   End
   Begin VB.Label lblPauseTime 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Seconds to pause between shapes:"
      Height          =   195
      Left            =   240
      TabIndex        =   2
      Top             =   840
      Width           =   2550
   End
   Begin VB.Label lblPrintNum 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Number of shapes per screen:"
      Height          =   195
      Left            =   645
      TabIndex        =   0
      Top             =   360
      Width           =   2130
   End
End
Attribute VB_Name = "frmConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOkay_Click()
    If Not Validate(txtPrintNum, 1, 9999, "Number of shapes per screen can only range between 1 and 9999.") Then Exit Sub
    If Not Validate(txtPauseTime, 0, 10, "Pause time between each shape cannot exceed 10 seconds.") Then Exit Sub
    If Not Validate(txtPhaseSep, 0, 30, "Pause time between screens cannot exceed 30 seconds.") Then Exit Sub
    If Not Validate(txtLineWidth, 1, 10, "Line width can only range between 1 to 10 points.") Then Exit Sub
    PrintNum = Val(txtPrintNum)
    LineWidth = Val(txtLineWidth)
    PauseTime = Val(txtPauseTime)
    PhaseSep = Val(txtPhaseSep)
    SaveSettings
    End
End Sub

Sub SelectTB(TBox As TextBox)
    TBox.SelStart = 0
    TBox.SelLength = Len(TBox)
End Sub

Private Sub cmdCancel_Click()
    End
End Sub

Function Validate(TBox As TextBox, Min As Double, Max As Double, Optional ErrMsg As String) As Boolean
Const ERR_COL = 255
    If IsNumeric(TBox) Then
        If Val(TBox) >= Min And Val(TBox) <= Max Then
            Validate = True
        End If
    End If
    TBox.ForeColor = ERR_COL * (Validate + 1)
    If Not Validate And ErrMsg > "" Then
        MsgBox ErrMsg
        SelectTB TBox
    End If
End Function

Private Sub Form_Load()
    txtLineWidth = LineWidth
    txtPrintNum = PrintNum
    txtPauseTime = PauseTime
    txtPhaseSep = PhaseSep
End Sub

Private Sub cmdAbout_Click()
    frmAbout.Show
End Sub

Private Sub txtLineWidth_Change()
    Validate txtLineWidth, 1, 10
End Sub

Private Sub txtLineWidth_GotFocus()
    SelectTB txtLineWidth
End Sub

Private Sub txtPauseTime_Change()
    Validate txtPauseTime, 0, 10
End Sub

Private Sub txtPauseTime_GotFocus()
    SelectTB txtPauseTime
End Sub

Private Sub txtPhaseSep_Change()
    Validate txtPhaseSep, 0, 30
End Sub

Private Sub txtPhaseSep_GotFocus()
    SelectTB txtPhaseSep
End Sub

Private Sub txtPrintNum_Change()
    Validate txtPrintNum, 1, 9999
End Sub

Private Sub txtPrintNum_GotFocus()
    SelectTB txtPrintNum
End Sub
