VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Form"
   ClientHeight    =   6045
   ClientLeft      =   60
   ClientTop       =   435
   ClientWidth     =   6615
   LinkTopic       =   "Form1"
   ScaleHeight     =   6045
   ScaleWidth      =   6615
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Next"
      Height          =   495
      Left            =   4440
      TabIndex        =   11
      Top             =   4800
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   1320
      TabIndex        =   10
      Top             =   4080
      Width           =   1215
   End
   Begin VB.CheckBox Check1 
      Caption         =   "OK"
      Enabled         =   0   'False
      Height          =   495
      Left            =   4440
      TabIndex        =   9
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "OK"
      Height          =   495
      Left            =   4440
      TabIndex        =   8
      Top             =   4080
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   4440
      PasswordChar    =   "+"
      TabIndex        =   5
      Top             =   1800
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   4440
      TabIndex        =   4
      Top             =   480
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   4440
      TabIndex        =   1
      Top             =   2280
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   4440
      TabIndex        =   0
      Top             =   960
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Password"
      Height          =   495
      Left            =   1200
      TabIndex        =   7
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "UserName"
      Height          =   495
      Left            =   1200
      TabIndex        =   6
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Age:"
      Height          =   495
      Left            =   1320
      TabIndex        =   3
      Top             =   2280
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Name:"
      Height          =   495
      Left            =   1320
      TabIndex        =   2
      Top             =   1080
      Visible         =   0   'False
      Width           =   1215
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
End
End Sub

Private Sub Command2_Click()
If Text4.Text = "3" Then
Label3.Visible = False
Label4.Visible = False
Text3.Visible = False
Text4.Visible = False
Label1.Visible = True
Label2.Visible = True
Text1.Visible = True
Text2.Visible = True
Text1.Text = InputBox("INsert your Name")
Text2.Text = InputBox("Insert Your Age")
End If
If Text1.Text = "Amin" And Text2.Text = "20" Then
Check1.Enabled = True
Command2.Visible = False
Command3.Visible = True
End If
End Sub

Private Sub Command3_Click()
If Check1.Value = 1 Then
MsgBox "You Are,You Are"
frmMain.Hide
frmStep.Show
End If
End Sub

Private Sub Form_Load()
Me.Height = Screen.Height / 2
Me.Width = Screen.Width / 2
Me.Top = (Screen.Height - Me.Height) / 2
Me.Left = (Screen.Width - Me.Width) / 2
End Sub
