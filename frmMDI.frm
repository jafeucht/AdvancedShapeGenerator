VERSION 5.00
Begin VB.MDIForm frmMDI 
   BackColor       =   &H8000000C&
   Caption         =   "Shape DLL Demonstration Environment"
   ClientHeight    =   5415
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   6150
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNewShp 
         Caption         =   "&New Shape"
      End
      Begin VB.Menu mnuFilePause 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuFileWindows 
      Caption         =   "&Windows"
      WindowList      =   -1  'True
      Begin VB.Menu mnuWindowsToolbox 
         Caption         =   "&Toolbox"
      End
   End
End
Attribute VB_Name = "frmMDI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' frmMDI
' This form contains all the instances of frmTest loaded.

' For best results, change the screen size to 1024*768 pixels
' or higher.

Option Explicit

Dim WindowIndex As Integer

Private Sub MDIForm_Load()
    CheckToolbox
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    End
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub

' Create new frmTest shape instance
Private Sub mnuFileNewShp_Click()
Dim NewForm As New frmTest
    WindowIndex = WindowIndex + 1
    NewForm.Show
    NewForm.Caption = "Shape " & WindowIndex
End Sub

' Hide/Show ToolBox
Private Sub mnuWindowsToolbox_Click()
    Select Case mnuWindowsToolbox.Checked
        Case True
            Unload frmToolbox
        Case False
            frmToolbox.Show
    End Select
End Sub

' Check/Uncheck the ToolBox item in the Windows menu.
Public Sub ChangeToolBoxState(Loaded As Boolean)
    mnuWindowsToolbox.Checked = Loaded
End Sub

 ' See if ToolBox is open.
Sub CheckToolbox()
    ChangeToolBoxState frmToolbox.Visible
End Sub
