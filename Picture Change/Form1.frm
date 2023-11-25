VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6360
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8100
   LinkTopic       =   "Form1"
   ScaleHeight     =   6360
   ScaleWidth      =   8100
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   3000
      Left            =   0
      Top             =   5880
   End
   Begin VB.Image Image2 
      Height          =   4800
      Left            =   4320
      Stretch         =   -1  'True
      Top             =   480
      Width           =   3480
   End
   Begin VB.Image Image1 
      Height          =   4800
      Left            =   120
      Stretch         =   -1  'True
      Top             =   480
      Width           =   3840
   End
   Begin VB.Image tmpImage 
      Height          =   975
      Left            =   3240
      Top             =   5160
      Visible         =   0   'False
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Image1.Picture = LoadPicture("C:\Documents and Settings\All Users\Documents\My Pictures\Sample Pictures\sunset.jpg")
Image2.Picture = LoadPicture("C:\Documents and Settings\All Users\Documents\My Pictures\Sample Pictures\Blue hills.jpg")
End Sub

Private Sub Timer1_Timer()
tmpImage.Picture = LoadPicture("C:\Documents and Settings\All Users\Documents\My Pictures\Sample Pictures\sunset.jpg")
Image1.Picture = LoadPicture("C:\Documents and Settings\All Users\Documents\My Pictures\Sample Pictures\Blue hills.jpg")
Image2.Picture = LoadPicture("C:\Documents and Settings\All Users\Documents\My Pictures\Sample Pictures\sunset.jpg")
End Sub
