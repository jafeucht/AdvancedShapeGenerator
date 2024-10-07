Attribute VB_Name = "mdTest"
' mdTest
' This module contains some declarations, types, and functions
' used in other parts of the project.

Option Explicit

Type PointAPI
    X As Double
    Y As Double
End Type

' This variable stores information on the selected shape window.
Public SelWindow As frmTest

'------------------------------------------------------
' Used in color functions. Contains R, G, and B values
' for a color.
'------------------------------------------------------
Type cColor
    R As Integer
    G As Integer
    B As Integer
End Type

Public Function SetPnt(X As Double, Y As Double) As PointAPI
    ' Allow PointAPI data change in one command
    With SetPnt
        .X = X
        .Y = Y
    End With
End Function

'--------------------------------------------------
' Color Function, converts a long color value into
' a cColor Type
' -------------------------------------------------
Public Function ColorRGB(ByVal LongColor As Long) As cColor
    ColorRGB.R = LongColor Mod 256
    ColorRGB.G = (LongColor And &HFF00FF00) / 256
    ColorRGB.B = (LongColor And &HFF0000) / 256 ^ 2
End Function

'-------------------------------------------------
' Color Function, finds a color value Fraction *
' the distance between two colors.
'-------------------------------------------------
Public Function ColorBetween(StartCol As Long, EndCol As Long, Fraction As Currency) As Long
Dim ResultColor As cColor, Col1 As cColor, Col2 As cColor
    Col1 = ColorRGB(StartCol)
    Col2 = ColorRGB(EndCol)
    ResultColor.R = Fraction * (Col1.R - Col2.R) + Col2.R
    ResultColor.G = Fraction * (Col1.G - Col2.G) + Col2.G
    ResultColor.B = Fraction * (Col1.B - Col2.B) + Col2.B
    ColorBetween = RGB(ResultColor.R, ResultColor.G, ResultColor.B)
End Function
