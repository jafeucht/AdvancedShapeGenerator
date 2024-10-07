VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "csShape"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' csShape
' This file is the actual shape class, used to sequence vectors in
' order to achieve a shape. Virtually any shape is possible and
' easy to create with this class.

' Please keep your code clean!

Option Explicit

Const PI = 3.14159265358979
Const MSK_SEP1 = "<"
Const MSK_SEP2 = ":"

Public Field As Object ' Form / PictureBox, etc. to draw to

' X/Y coordinate structure
Private Type PointAPI
    X As Double
    Y As Double
End Type

' Data structor for vector information
Private Type Vector
    Length As Double
    Angle As Double
End Type

' Each shape is an array of vectors
Dim Vectors() As Vector

' Shape property variables
Dim pStartPt As PointAPI ' Beginning Point
Dim pZoom As Double ' Percent Value
Dim pVisible As Boolean
Dim pColor As Long
Dim pRotate As Double

' Occurs before every time the class attempts to draw a line to the
' selected draw area
Event DrawLine(LineNum As Long, ByRef Color As Long)

Property Get Left() As Double
    Left = pStartPt.X
End Property

Property Let Left(NewValue As Double)
    pStartPt.X = NewValue
End Property

Property Get Top() As Double
    Top = pStartPt.Y
End Property

Property Let Top(NewValue As Double)
    pStartPt.Y = NewValue
End Property

Property Get Zoom() As Double
    Zoom = pZoom
End Property

Property Let Zoom(NewValue As Double)
    pZoom = NewValue
End Property

Property Get Visible() As Boolean
    Visible = pVisible
End Property

Property Let Visible(NewValue As Boolean)
    pVisible = NewValue
End Property

Property Get Color() As Long
    Color = pColor
End Property

Property Let Color(NewValue As Long)
    pColor = NewValue
End Property

Property Get Rotate() As Double
    Rotate = pRotate
End Property

Property Let Rotate(NewValue As Double)
    pRotate = NewValue
End Property

' Check a number to see if it is within certain bounds.
Private Function Validate(Index As Integer, Optional Min As Integer = 1, Optional Max As Integer, Optional RaiseErr As Boolean) As Boolean
    If Max < Min And Max = 0 Then Max = VectorCnt
    Validate = Index >= Min And Index <= Max
    If RaiseErr And Not Validate Then Err.Raise 9
End Function

Private Function Sine(Value As Double) As Double
    Sine = Sin((Value * 2 * PI) / 360)
End Function

Private Function Cosine(Value As Double) As Double
    Cosine = Cos((Value * 2 * PI) / 360)
End Function

Public Function NumSequences() As Integer
Dim i As Long, CurAng As Long, VectorAng As Long
    ' First, add up the angles of each vector.
    For i = 1 To VectorCnt
        VectorAng = VectorAng + Vectors(i).Angle
        If VectorAng >= 360 Then VectorAng = VectorAng - 360
    Next i
    ' Now, add up VectorAng until it hits 360
    i = 0
    Do
        i = i + 1
        CurAng = CurAng + VectorAng
        If CurAng >= 360 Then CurAng = CurAng - 360
    Loop Until CurAng = 0
    ' Return total number of lines
    NumSequences = i
End Function

' The main function, actually draws a shape.
Public Sub DrawShape(Optional Test As Boolean)
Dim i As Integer, j As Integer
Dim CurPnt As PointAPI, OldPnt As PointAPI
Dim PntMov As PointAPI, Ang As Double, LoopCnt As Long
    CurPnt = pStartPt
    If VectorCnt = 0 Then Exit Sub
    Ang = pRotate ' Rotate image
    Do
        OldPnt = CurPnt ' Store last point
        i = i + 1 ' Next vector
        If i > VectorCnt Then
            i = 1
            ' Check to see if we're heading off into eternity...
            If Ang = pRotate Then Exit Do
        End If
        Ang = Ang + Vectors(i).Angle
        Do While Ang >= 360
            ' Reduce angles as to prevent overflow
            Ang = Ang - 360
        Loop
        ' Count each loop
        LoopCnt = LoopCnt + 1
        ' Derive new coordinates based on new angle, vector length, and zoom
        PntMov = SetPnt(Cosine(Ang) * Vectors(i).Length * pZoom / 100, Sine(Ang) * Vectors(i).Length * pZoom / 100)
        ' Add new coordinates to current point
        CurPnt = SetPnt(CurPnt.X + PntMov.X, CurPnt.Y - PntMov.Y)
        ' Draw line from last point to new point
        If pVisible And Not Test Then
            RaiseEvent DrawLine(LoopCnt, pColor)
            Field.Line (OldPnt.X, OldPnt.Y)-(CurPnt.X, CurPnt.Y), pColor
        End If
    ' Loop until the shape completes complete circuit
    Loop Until DistDiff(pStartPt, CurPnt) = 0 And i = VectorCnt
End Sub

Private Function SetPnt(X As Double, Y As Double) As PointAPI
    ' Allow PointAPI data change in one command
    With SetPnt
        .X = X
        .Y = Y
    End With
End Function

Private Function DistDiff(Pt1 As PointAPI, Pt2 As PointAPI) As Double
    DistDiff = Abs(Sqr((Pt1.X - Pt2.X) ^ 2 + (Pt1.Y - Pt2.Y) ^ 2))
End Function

' Allows the user to enter vectors using a coded string, for example:
' 5<30:6<20:9<12. contains three vectors to load
Public Function TakeVectorStr(MaskStr As String) As Long
Dim i As Integer, TmpStr As String
Dim RetLength As Double, RetAngle As Double
Dim CharPos1 As Integer, CharPos2 As Integer
    Erase Vectors
    TmpStr = Trim$(MaskStr)
    On Error GoTo BadMaskStr
    Do ' Loop through vector string
        CharPos1 = InStr(1, TmpStr, MSK_SEP1) ' Search for "<"
        CharPos2 = InStr(1, TmpStr, MSK_SEP2) ' Search for ":"
        ' MaskStr isn't formated correctly, quit function
        If CharPos1 = 0 Or CharPos2 = 0 Then GoTo BadMaskStr
        ' Find Length parameter
        RetLength = Val(Mid(TmpStr, 1, CharPos1 - 1))
        ' Find Angle parameter
        RetAngle = Val(Mid(TmpStr, CharPos1 + 1, CharPos2 - 1 - CharPos1 + 1))
        ' Add data to vector array
        AddVector VectorCnt + 1, RetLength, RetAngle
        ' Remove read vector from the vector string
        TmpStr = Mid(TmpStr, CharPos2 + 1)
    Loop Until Len(TmpStr) = 0
    Exit Function
BadMaskStr:
    ' Return number of unread characters
    TakeVectorStr = Len(TmpStr)
End Function

' Creates a simple code describing all the vectors
Public Function CreateVectorStr() As String
Dim i As Integer, TmpStr As String
    For i = 1 To VectorCnt
        With Vectors(i)
            TmpStr = TmpStr & .Length & MSK_SEP1 & .Angle & MSK_SEP2
        End With
    Next i
    CreateVectorStr = TmpStr
End Function

Public Sub GetVector(Index As Integer, ByRef Length As Double, ByRef Angle As Double)
    If Not Validate(Index, , VectorCnt) Then Exit Sub
    With Vectors(Index)
        Length = .Length
        Angle = .Angle
    End With
End Sub

Public Sub SetVector(Index As Integer, Length As Double, Angle As Double)
    If Not Validate(Index, 1, VectorCnt) Then Exit Sub
    With Vectors(Index)
        .Length = Length
        .Angle = Angle
    End With
End Sub

Public Sub AddVector(Index As Integer, Length As Double, Angle As Double)
Dim i As Integer
    If Index = 0 Then Index = VectorCnt
    ReDim Preserve Vectors(1 To VectorCnt + 1) ' Create new vector
    For i = VectorCnt To Index + 1 Step -1 ' Shift vectors as to allow new vector
        Vectors(i) = Vectors(i - 1)
    Next i
    With Vectors(Index) ' Set up new vector
        .Length = Length
        .Angle = Angle
    End With
End Sub

Public Sub SetVectorCnt(Value As Integer)
Dim i As Integer
    If Value = 0 Then ClearVectors: Exit Sub
    If Value > VectorCnt Then
        For i = VectorCnt + 1 To Value
            AddVector VectorCnt + 1, 0, 0
        Next i
    Else
        ReDim Preserve Vectors(1 To Value)
    End If
End Sub

Public Sub RemoveVector(Index As Integer)
Dim i As Integer
    ' Move unwanted vector to back of stack
    For i = Index To VectorCnt
        Vectors(Index) = Vectors(Index + 1)
    Next i
    ' Delete last vector
    ReDim Vectors(1 To VectorCnt - 1)
End Sub

Public Sub ClearVectors()
    Erase Vectors
End Sub

Public Function VectorCnt() As Integer
    ' Returns the array count of Vectors().
    On Error Resume Next
    VectorCnt = UBound(Vectors)
End Function

Private Sub Class_Initialize()
    pZoom = 100
    pVisible = True
End Sub
