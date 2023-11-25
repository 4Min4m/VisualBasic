VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5055
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6465
   LinkTopic       =   "Form1"
   ScaleHeight     =   5055
   ScaleWidth      =   6465
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refresh"
      Height          =   495
      Left            =   480
      TabIndex        =   13
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton cmdChap 
      Caption         =   "Chap"
      Height          =   495
      Left            =   480
      TabIndex        =   12
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton cmdMin 
      Caption         =   "Min"
      Height          =   495
      Left            =   480
      TabIndex        =   11
      Top             =   4200
      Width           =   1215
   End
   Begin VB.CommandButton cmdMax 
      Caption         =   "Max"
      Height          =   495
      Left            =   480
      TabIndex        =   10
      Top             =   3360
      Width           =   1215
   End
   Begin VB.TextBox Text5 
      Height          =   495
      Left            =   4080
      TabIndex        =   4
      Top             =   4200
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      Height          =   495
      Left            =   4080
      TabIndex        =   3
      Top             =   3360
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   4080
      TabIndex        =   2
      Top             =   2400
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   4080
      TabIndex        =   1
      Top             =   1440
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   4080
      TabIndex        =   0
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "Min"
      Height          =   495
      Left            =   2280
      TabIndex        =   9
      Top             =   4200
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Max"
      Height          =   495
      Left            =   2280
      TabIndex        =   8
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   495
      Left            =   2280
      TabIndex        =   7
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   495
      Left            =   2280
      TabIndex        =   6
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   495
      Left            =   2400
      TabIndex        =   5
      Top             =   480
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a, b, c, Max, Min As Integer
Private Sub cmdMax_Click()
a = Int(Text1.Text)
b = Int(Text2.Text)
c = Int(Text3.Text)
Max = a
If b > Max Then Max = b
If c > Max Then Max = c

End Sub

Private Sub cmdMin_Click()
a = Int(Text1.Text)
b = Int(Text2.Text)
c = Int(Text3.Text)
Min = a
If b < Min Then Min = b

End Sub
Private Sub cmdChap_Click()
Text4.Text = Str(Max)
Text5.Text = Str(Min)
End Sub

Private Sub cmdRefresh_Click()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
End Sub
