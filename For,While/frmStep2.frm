VERSION 5.00
Begin VB.Form frmStep2 
   Caption         =   "Form1"
   ClientHeight    =   5520
   ClientLeft      =   60
   ClientTop       =   435
   ClientWidth     =   6735
   LinkTopic       =   "Form1"
   ScaleHeight     =   5520
   ScaleWidth      =   6735
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "OK"
      Height          =   495
      Left            =   3960
      TabIndex        =   3
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Next"
      Height          =   495
      Left            =   3960
      TabIndex        =   2
      Top             =   4440
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "end"
      Height          =   495
      Left            =   1440
      TabIndex        =   1
      Top             =   3360
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   2415
      Left            =   1320
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   360
      Width           =   3855
   End
End
Attribute VB_Name = "frmStep2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
End
End Sub

Private Sub Command2_Click()
frmStep2.Hide
frmStep3.Show
End Sub

Private Sub Command3_Click()
For i = 1 To 9
 For j = 1 To i
 Text1.Text = Text1.Text + "*"
 Next j
Text1.Text = Text1.Text + vbCrLf
Next i
End Sub

Private Sub Form_Load()
Me.Height = Screen.Height / 2
Me.Width = Screen.Width / 2
Me.Top = (Screen.Height - Me.Height) / 2
Me.Left = (Screen.Width - Me.Width) / 2
End Sub
