VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1230
   ClientLeft      =   60
   ClientTop       =   435
   ClientWidth     =   6150
   LinkTopic       =   "Form1"
   ScaleHeight     =   1230
   ScaleWidth      =   6150
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   1080
      Top             =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "This is a test"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   1800
      TabIndex        =   0
      Top             =   360
      Width           =   2700
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Timer1_Timer()
If Label1.ForeColor = &H80000012 Then
Label1.ForeColor = vbRed
ElseIf Label1.ForeColor = vbRed Then Label1.ForeColor = vbBlue
Else: Label1.ForeColor = &H80000012
End If
End Sub

