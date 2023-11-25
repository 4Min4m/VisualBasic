VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Sabt"
   ClientHeight    =   7575
   ClientLeft      =   165
   ClientTop       =   870
   ClientWidth     =   7725
   LinkTopic       =   "Form1"
   ScaleHeight     =   7575
   ScaleWidth      =   7725
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   2880
      TabIndex        =   9
      Top             =   6240
      Width           =   1215
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Next"
      Enabled         =   0   'False
      Height          =   495
      Left            =   5280
      TabIndex        =   8
      Top             =   6240
      Width           =   1215
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Married"
      Height          =   615
      Left            =   1200
      TabIndex        =   7
      Top             =   4800
      Width           =   2415
   End
   Begin VB.Frame Frame2 
      Caption         =   "Age"
      Height          =   2295
      Left            =   960
      TabIndex        =   3
      Top             =   2280
      Width           =   5535
      Begin VB.OptionButton Option5 
         Caption         =   "A-40"
         Height          =   495
         Left            =   240
         TabIndex        =   6
         Top             =   1440
         Width           =   1335
      End
      Begin VB.OptionButton Option4 
         Caption         =   "30-40"
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   960
         Width           =   1935
      End
      Begin VB.OptionButton Option3 
         Caption         =   "U-30"
         Height          =   495
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   2655
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Genre"
      Height          =   1575
      Left            =   960
      TabIndex        =   0
      Top             =   600
      Width           =   5535
      Begin VB.OptionButton Option2 
         Caption         =   "Woman"
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   960
         Width           =   1695
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Man"
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Value           =   -1  'True
         Width           =   1935
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileOpen 
         Caption         =   "Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFile_ 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "Exit"
         Shortcut        =   ^E
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
If Option1.Value = True And Option4.Value And Check1.Value = 1 Then
cmdOk.Enabled = True
MsgBox "You're Accept"
ElseIf Check1.Value = 0 Then
cmdOk.Enabled = False
ElseIf Option2.Value = True And Option3.Value Then
cmdOk.Enabled = True
MsgBox "You'll Accept"
ElseIf Check1.Value = 0 Then
cmdOk.Enabled = False
End If
End Sub

Private Sub cmdCancel_Click()
End
End Sub
Private Sub mnuFileExit_Click()
End
End Sub
