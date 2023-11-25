VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5730
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4530
   LinkTopic       =   "Form1"
   ScaleHeight     =   5730
   ScaleWidth      =   4530
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Ravie"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      Sorted          =   -1  'True
      TabIndex        =   4
      Top             =   240
      Width           =   1815
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Õ–› ò«„· ·Ì” "
      Height          =   495
      Left            =   2040
      TabIndex        =   3
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "«›“Êœ‰"
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   3960
      Width           =   1215
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "Õ–›"
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   4680
      Width           =   1215
   End
   Begin VB.CommandButton cmdNumber 
      Caption         =   "‘„«—‘"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   3120
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAdd_Click()
Element = InputBox("⁄‰’— „Ê—œ ‰Ÿ— —« Ê«—œ ‰„«ÌÌœ")
Combo1.AddItem Element
End Sub

Private Sub cmdClear_Click()
Combo1.Clear
End Sub

Private Sub cmdNumber_Click()
Tedad = Combo1.ListCount
MsgBox (" ⁄œ«œ ⁄‰«’— ·Ì”  »—«»— «”  »«:" & Tedad)
End Sub

Private Sub cmdRemove_Click()
Combo1.RemoveItem Combo1.ListIndex
MsgBox ":‘„«—Â ⁄‰’—" & Combo1.ListIndex
End Sub


Private Sub Form_Load()
Combo1.AddItem "roz"
Combo1.AddItem "orkideh"
Combo1.AddItem "maryam"
End Sub
