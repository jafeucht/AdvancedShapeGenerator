VERSION 5.00
Begin VB.Form frmTest 
   AutoRedraw      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   3735
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3975
   DrawWidth       =   2
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   249
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   265
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' frmTest
' This is the main form, where the shape is exhibited. Multiple
' instances of this form can be loaded to view more than one shape.

Option Explicit

' Declare new csShape variable
Public WithEvents Gadget As csShape
Attribute Gadget.VB_VarHelpID = -1

' Properties to use with frmToolbox
Public ShowGrid As Boolean
Public ShowAxis As Boolean
Public Gradient As Boolean
Public BkColor As Long
Public BaseColor As Long
Public GradColor As Long
Public LineWidth As Integer

Dim NumLines As Long ' Stores the number of lines in a shape

Private Sub Form_Activate()
    ' Select this shape for viewing in the toolbox
    Set SelWindow = Me
    frmToolbox.GetSettings
    frmToolbox.Activate True
End Sub

Private Sub DrawShape()
    BackColor = BkColor
    Cls ' Clear work area
    With Gadget
        DrawGrid SetPnt(20 * (.Zoom / 100), 20 * (.Zoom / 100)) ' Draw grid
        DrawWidth = LineWidth
        
        ' Retrieve the number of sequences times the number of
        ' vectors. This is how many total lines would have to be drawn
        ' to create an enclosed shape. The NumLines variable is used
        ' in making the shape's color gradient.
        NumLines = .NumSequences * .VectorCnt
        
        .Color = BaseColor
        
        ' Draw the shape to the form
        .DrawShape
    End With
End Sub

Private Sub DrawGrid(Pos As PointAPI)
Dim i As Integer, CurPos As PointAPI
    On Error Resume Next
    If Pos.X = 0 Or Pos.Y = 0 Then Exit Sub
    Do While Pos.X < 5 Or Pos.Y < 5
        Pos = SetPnt(Pos.X * 10, Pos.Y * 10)
    Loop
    CurPos = SetPnt(Gadget.Left Mod Pos.X - Pos.X, Gadget.Top Mod Pos.Y - Pos.Y)
    ' Draw grid to drawing field.
    If ShowGrid Then
        DrawWidth = 1
        For i = 1 To ScaleHeight / Pos.Y
            Line (0, CurPos.Y + i * Pos.Y)-(ScaleWidth, CurPos.Y + i * Pos.Y), &HAAAAAA
        Next i
        For i = 1 To ScaleWidth / Pos.X
            Line (CurPos.X + i * Pos.X, 0)-(CurPos.X + i * Pos.X, ScaleHeight), &HAAAAAA
        Next i
    End If
    ' Draw Axis to drawing field.
    If ShowAxis Then
        DrawWidth = 2
        Line (Gadget.Left, 0)-(Gadget.Left, ScaleHeight), vbBlue
        Line (0, Gadget.Top)-(ScaleWidth, Gadget.Top), vbRed
    End If
End Sub

Private Sub Form_Load()
    Set Gadget = New csShape
    ' Initialization settings
    With Gadget
        Set .Field = Me
        .Left = ScaleWidth / 2
        .Top = ScaleHeight / 2
    End With
    ShowGrid = True
    ShowAxis = True
    Gradient = True
    BkColor = BackColor
    BaseColor = &HFF ' Red
    GradColor = &HFF00& ' Green
    LineWidth = 1
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        ' Center axis on mousepoint.
        Gadget.Left = CDbl(X)
        Gadget.Top = CDbl(Y)
    End If
    DrawShape
End Sub

Private Sub Form_Resize()
    ' Redraw shape
    DrawShape
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ' Disable toolbox.
    frmToolbox.Activate False
End Sub

' Notice that the Object menu at the top of this code window
' contains an object called Gadget, then there are no controls on
' the form. That is because the "WithEvents" part of the line:
'
' Public WithEvents Gadget As csShape
'
' Includes the Gadget object into the object window, so events
' contained in the control may be accessed.
Private Sub Gadget_DrawLine(LineNum As Long, Color As Long)
    If Not Gradient Then Exit Sub
    ' Find the color (Line*100/Numlines)% the way on a gradient
    ' between GradColor and BaseColor
    Color = ColorBetween(GradColor, BaseColor, LineNum / NumLines)
End Sub
