VERSION 5.00
Begin VB.Form frmStep3 
   Caption         =   "Form1"
   ClientHeight    =   6570
   ClientLeft      =   60
   ClientTop       =   435
   ClientWidth     =   7245
   LinkTopic       =   "Form1"
   ScaleHeight     =   6570
   ScaleWidth      =   7245
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "OK"
      Height          =   495
      Left            =   3840
      TabIndex        =   3
      Top             =   4200
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Next"
      Height          =   495
      Left            =   3840
      TabIndex        =   2
      Top             =   5520
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "end"
      Height          =   495
      Left            =   1200
      TabIndex        =   1
      Top             =   4200
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   2775
      Left            =   1080
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   480
      Width           =   4215
   End
End
Attribute VB_Name = "frmStep3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
End
End Sub

Private Sub Command3_Click()
i = 1
While i <= 9
 j = 1
 While j <= i
 Text1.Text = Text1.Text + Str(j)
 j = j + 1
 Wend
Text1.Text = Text1.Text + vbCrLf
i = i + 1
Wend
End Sub

Private Sub Form_Load()
Me.Height = Screen.Height / 2
Me.Width = Screen.Width / 2
Me.Top = (Screen.Height - Me.Height) / 2
Me.Left = (Screen.Width - Me.Width) / 2
End Sub
