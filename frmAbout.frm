VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About ..."
   ClientHeight    =   1695
   ClientLeft      =   1140
   ClientTop       =   1560
   ClientWidth     =   3675
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1695
   ScaleWidth      =   3675
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   390
      Left            =   135
      TabIndex        =   1
      Top             =   1170
      Width           =   3390
   End
   Begin VB.Label lblDescription 
      BackStyle       =   0  'Transparent
      Caption         =   "This screen saver was created from the skelleton screen saver code provided by Igguk at Planet Soucre Code."
      Height          =   720
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   3375
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Shape Screen Saver"
      Height          =   240
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3390
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()

    ' Unload and deallocate the about box.
    Unload frmAbout
End Sub
