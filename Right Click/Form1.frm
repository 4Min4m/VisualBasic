VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7965
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   9060
   LinkTopic       =   "Form1"
   ScaleHeight     =   7965
   ScaleWidth      =   9060
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu mnuTest 
      Caption         =   "&Test"
      Visible         =   0   'False
      Begin VB.Menu mnuTestCut 
         Caption         =   "Cut"
      End
      Begin VB.Menu mnuTestCopy 
         Caption         =   "Copy"
      End
      Begin VB.Menu mnuTestPaste 
         Caption         =   "Paste"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 And Shift = 1 Then
PopupMenu mnuTest, , X, Y
End If
End Sub

