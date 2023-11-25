VERSION 5.00
Begin VB.Form frmBirthDate 
   Caption         =   "›—„ „Õ«”»Â ”‰"
   ClientHeight    =   6240
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6810
   LinkTopic       =   "Form1"
   ScaleHeight     =   6240
   ScaleWidth      =   6810
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnOk 
      Caption         =   "OK"
      Height          =   495
      Left            =   3000
      TabIndex        =   4
      Top             =   3600
      Width           =   1215
   End
   Begin VB.TextBox txtBirthDate 
      Height          =   495
      Left            =   3000
      TabIndex        =   3
      Top             =   1800
      Width           =   1215
   End
   Begin VB.TextBox txtName 
      Height          =   495
      Left            =   3000
      TabIndex        =   2
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "”«·  Ê·œ:"
      Height          =   495
      Left            =   600
      TabIndex        =   1
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "‰«„:"
      Height          =   495
      Left            =   600
      TabIndex        =   0
      Top             =   720
      Width           =   1215
   End
End
Attribute VB_Name = "frmBirthDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnOk_Click()
Sen = 1389 - Int(txtBirthDate.Text)
answer = MsgBox(" ”‰ " & Str(Sen) & " ”«· ", vbYesNo + vbQuestion, " ÊÃÂ")
If answer = 6 Then
txtName.Text = ""
txtBirthDate.Text = ""
ElseIf answer = 7 Then
End
End If
End Sub
