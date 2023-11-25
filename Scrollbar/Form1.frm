VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4815
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6720
   LinkTopic       =   "Form1"
   ScaleHeight     =   4815
   ScaleWidth      =   6720
   StartUpPosition =   3  'Windows Default
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   4560
      Width           =   6735
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   4575
      Left            =   6480
      TabIndex        =   1
      Top             =   0
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   2640
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
HScroll1.Min = 0
HScroll1.Max = Me.Width - Text1.Width
HScroll1.Value = 0
HScroll1.SmallChange = Me.Width \ 5
HScroll1.LargeChange = Me.Width \ 5
VScroll1.Min = 0
VScroll1.Max = Me.Height - Text1.Height - 3 * (HScroll1.Height)
VScroll1.Value = 0
VScroll1.SmallChange = Me.Height \ 5
VScroll1.LargeChange = Me.Height \ 5
End Sub
Private Sub HScroll1_Change()
Text1.Left = HScroll1.Value
End Sub
Private Sub VScroll1_Change()
Text1.Top = VScroll1.Value
End Sub
