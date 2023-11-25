VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4905
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12195
   LinkTopic       =   "Form1"
   ScaleHeight     =   4905
   ScaleWidth      =   12195
   StartUpPosition =   3  'Windows Default
   Begin VB.FileListBox File1 
      Height          =   2235
      Left            =   3120
      Pattern         =   "*.jpg"
      TabIndex        =   2
      Top             =   960
      Width           =   2775
   End
   Begin VB.DirListBox Dir1 
      Height          =   2340
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   2535
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   2175
      Left            =   6480
      Stretch         =   -1  'True
      Top             =   960
      Width           =   2655
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Dir1_Change()
File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive
End Sub

Private Sub File1_Click()
Image1.Picture = LoadPicture(File1.Path + "\" + File1.FileName)
End Sub

