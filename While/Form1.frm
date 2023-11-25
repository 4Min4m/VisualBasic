VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5460
   ClientLeft      =   60
   ClientTop       =   435
   ClientWidth     =   6435
   LinkTopic       =   "Form1"
   ScaleHeight     =   5460
   ScaleWidth      =   6435
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   495
      Left            =   3720
      TabIndex        =   1
      Top             =   3600
      Width           =   1215
   End
   Begin VB.ListBox List1 
      Height          =   2400
      ItemData        =   "Form1.frx":0000
      Left            =   960
      List            =   "Form1.frx":0002
      TabIndex        =   0
      Top             =   240
      Width           =   3975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim Name, Family, Age As String
i = 1
While i <= 3
Name = InputBox("Enter Name")
Family = InputBox("Enter Family")
Age = InputBox("Enter Age")
List1.AddItem (Name + " " + Family + " " + Age)
i = i + 1
Wend
End Sub
