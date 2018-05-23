VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   9615
   ClientLeft      =   855
   ClientTop       =   1035
   ClientWidth     =   13950
   LinkTopic       =   "Form1"
   ScaleHeight     =   9615
   ScaleWidth      =   13950
   Begin VB.CommandButton cmdCopy 
      Caption         =   "Copy"
      Height          =   495
      Left            =   12360
      TabIndex        =   2
      Top             =   9000
      Width           =   1455
   End
   Begin VB.PictureBox pictSource 
      Height          =   3855
      Left            =   4320
      Picture         =   "BitBltTest.frx":0000
      ScaleHeight     =   3795
      ScaleWidth      =   3795
      TabIndex        =   1
      Top             =   240
      Width           =   3855
   End
   Begin VB.PictureBox pictDest 
      Height          =   3855
      Left            =   240
      Picture         =   "BitBltTest.frx":101C6
      ScaleHeight     =   3795
      ScaleWidth      =   3795
      TabIndex        =   0
      Top             =   240
      Width           =   3855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCopy_Click()
     Call TransparentBlt(pictDest, pictSource.Picture, 10, 10, QBColor(15))
End Sub
