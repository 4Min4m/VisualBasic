VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6480
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7305
   LinkTopic       =   "Form1"
   ScaleHeight     =   6480
   ScaleWidth      =   7305
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "ok"
      Height          =   495
      Left            =   3000
      TabIndex        =   2
      Top             =   3000
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   1575
      Left            =   2880
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   600
      Width           =   2655
   End
   Begin VB.Label Label1 
      Caption         =   "adad"
      Height          =   495
      Left            =   1320
      TabIndex        =   1
      Top             =   840
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
For i = 12 To 99 Step 3
Text1.Text = Text1.Text + "," + Str(i)
Next i
End Sub
