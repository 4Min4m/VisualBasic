VERSION 5.00
Begin VB.Form frmStep 
   Caption         =   "Form1"
   ClientHeight    =   5400
   ClientLeft      =   60
   ClientTop       =   435
   ClientWidth     =   7050
   LinkTopic       =   "Form1"
   ScaleHeight     =   5400
   ScaleWidth      =   7050
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Next"
      Height          =   495
      Left            =   4200
      TabIndex        =   3
      Top             =   4320
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "End"
      Height          =   495
      Left            =   1800
      TabIndex        =   2
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   495
      Left            =   4080
      TabIndex        =   1
      Top             =   3120
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   2295
      Left            =   1320
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   360
      Width           =   4095
   End
End
Attribute VB_Name = "frmStep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
For i = 1 To 3
 For j = 1 To 3
 Text1.Text = Text1.Text + "" + Str(i * j)
 Next j
Text1.Text = Text1.Text + vbCrLf
Next i
End Sub

Private Sub Command2_Click()
End
End Sub

Private Sub Command3_Click()
frmStep.Hide
frmStep2.Show
End Sub

Private Sub Form_Load()
Me.Height = Screen.Height / 2
Me.Width = Screen.Width / 2
Me.Top = (Screen.Height - Me.Height) / 2
Me.Left = (Screen.Width - Me.Width) / 2
End Sub
