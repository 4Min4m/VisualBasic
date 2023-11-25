VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6525
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11625
   LinkTopic       =   "Form1"
   ScaleHeight     =   6525
   ScaleWidth      =   11625
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Caption         =   "Frame3"
      Height          =   4095
      Left            =   7440
      TabIndex        =   10
      Top             =   840
      Width           =   2895
      Begin VB.OptionButton Option9 
         Caption         =   "20"
         Height          =   615
         Left            =   360
         TabIndex        =   13
         Top             =   2640
         Width           =   1455
      End
      Begin VB.OptionButton Option8 
         Caption         =   "16"
         Height          =   855
         Left            =   360
         TabIndex        =   12
         Top             =   1440
         Width           =   1695
      End
      Begin VB.OptionButton Option7 
         Caption         =   "12"
         Height          =   495
         Left            =   360
         TabIndex        =   11
         Top             =   600
         Width           =   1215
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   3480
      TabIndex        =   8
      Top             =   5520
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   4215
      Left            =   4320
      TabIndex        =   1
      Top             =   840
      Width           =   2535
      Begin VB.OptionButton Option6 
         Caption         =   "blue"
         Height          =   495
         Left            =   240
         TabIndex        =   7
         Top             =   2400
         Width           =   1335
      End
      Begin VB.OptionButton Option5 
         Caption         =   "green"
         Height          =   495
         Left            =   240
         TabIndex        =   6
         Top             =   1440
         Width           =   1095
      End
      Begin VB.OptionButton Option4 
         Caption         =   "red"
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   480
         Value           =   -1  'True
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   4215
      Left            =   1080
      TabIndex        =   0
      Top             =   840
      Width           =   2895
      Begin VB.OptionButton Option3 
         Caption         =   "underline"
         Height          =   495
         Left            =   240
         TabIndex        =   4
         Top             =   2520
         Width           =   975
      End
      Begin VB.OptionButton Option2 
         Caption         =   "italic"
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   1560
         Width           =   975
      End
      Begin VB.OptionButton Option1 
         Caption         =   "bold"
         Height          =   495
         Left            =   240
         TabIndex        =   2
         Top             =   480
         Value           =   -1  'True
         Width           =   1215
      End
   End
   Begin VB.Label Label1 
      Caption         =   "this is a test"
      Height          =   495
      Left            =   3000
      TabIndex        =   9
      Top             =   120
      Width           =   2895
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Option1.Value = True Then
Label1.FontBold = True
Label1.FontItalic = False
Label1.FontUnderline = False
ElseIf Option2.Value = True Then
Label1.FontBold = False
Label1.FontItalic = True
Label1.FontUnderline = False
ElseIf Option3.Value Then
Label1.FontBold = False
Label1.FontItalic = False
Label1.FontUnderline = True
End If
If Option4.Value = True Then
Label1.ForeColor = vbRed
ElseIf Option5.Value = True Then
Label1.ForeColor = vbGreen
ElseIf Option6.Value = True Then
Label1.ForeColor = vbBlue
End If
If Option7.Value = True Then
Label1.FontSize = 12
ElseIf Option8.Value = True Then
Label1.FontSize = 16
ElseIf Option9.Value = True Then
Label1.FontSize = 20
End If
End Sub

Private Sub Form_Load()
Label1.ForeColor = vbRed
Label1.FontBold = True
End Sub
