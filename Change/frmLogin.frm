VERSION 5.00
Begin VB.Form frmLogin 
   Caption         =   "Login"
   ClientHeight    =   7305
   ClientLeft      =   3090
   ClientTop       =   4860
   ClientWidth     =   9060
   LinkTopic       =   "Form1"
   ScaleHeight     =   7305
   ScaleWidth      =   9060
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   1680
      TabIndex        =   5
      Top             =   4560
      Width           =   1215
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Height          =   495
      Left            =   4320
      TabIndex        =   4
      Top             =   4560
      Width           =   1215
   End
   Begin VB.TextBox txtPassword 
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   5040
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   2400
      Width           =   1935
   End
   Begin VB.TextBox txtUsername 
      Height          =   495
      Left            =   5160
      TabIndex        =   0
      Top             =   720
      Width           =   1935
   End
   Begin VB.Label Label3 
      Height          =   495
      Left            =   3600
      TabIndex        =   6
      Top             =   3480
      Width           =   2295
   End
   Begin VB.Label Label2 
      Caption         =   "Password"
      Height          =   495
      Left            =   840
      TabIndex        =   3
      Top             =   2400
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "UserName"
      Height          =   495
      Left            =   840
      TabIndex        =   2
      Top             =   720
      Width           =   1935
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
End
End Sub

Private Sub cmdOk_Click()
If txtPassword.Text = "123" Then
frmLogin.Hide
frmMain.Show
Else
Label3.Caption = "Wrong Code"
End If
End Sub

Private Sub Form_Load()
Me.Height = frmMain.Height / 2
Me.Width = frmMain.Width / 2
Me.Top = (frmMain.Height - frmLogin.Height) / 2
Me.Left = (frmMain.Width - frmLogin.Width) / 2
End Sub
