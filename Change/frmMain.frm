VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Main"
   ClientHeight    =   5910
   ClientLeft      =   165
   ClientTop       =   870
   ClientWidth     =   5460
   LinkTopic       =   "Form1"
   ScaleHeight     =   5910
   ScaleWidth      =   5460
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   1200
      TabIndex        =   7
      Top             =   4680
      Width           =   1215
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Height          =   495
      Left            =   3120
      TabIndex        =   6
      Top             =   4680
      Width           =   1215
   End
   Begin VB.TextBox txtHour 
      Height          =   615
      Left            =   3480
      TabIndex        =   2
      Top             =   3240
      Width           =   1695
   End
   Begin VB.TextBox txtMinute 
      Height          =   615
      Left            =   3480
      TabIndex        =   1
      Top             =   1920
      Width           =   1695
   End
   Begin VB.TextBox txtSecond 
      Height          =   615
      Left            =   3480
      TabIndex        =   0
      Top             =   600
      Width           =   1695
   End
   Begin VB.Label Label3 
      Caption         =   "Hour"
      Height          =   495
      Left            =   480
      TabIndex        =   5
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Minute"
      Height          =   495
      Left            =   480
      TabIndex        =   4
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Second"
      Height          =   495
      Left            =   480
      TabIndex        =   3
      Top             =   600
      Width           =   1215
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileOpen 
         Caption         =   "Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnu 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "Exit"
         Shortcut        =   ^E
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditFind 
         Caption         =   "Find"
         Shortcut        =   ^F
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuViewCode 
         Caption         =   "Code"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuViewObject 
         Caption         =   "Object"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
End
End Sub

Private Sub cmdOk_Click()
txtMinute.Text = txtSecond.Text / 60
txtHour.Text = txtMinute.Text / 60
End Sub

Private Sub Form_Load()
Me.Height = Screen.Height
Me.Width = Screen.Width
Me.Top = 0
Me.Left = 0
End Sub

Private Sub mnuFileExit_Click()
End
End Sub

Private Sub Text1_Change()

End Sub
