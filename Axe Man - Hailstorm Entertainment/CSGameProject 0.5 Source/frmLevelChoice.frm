VERSION 5.00
Begin VB.Form frmTeleporter 
   BackColor       =   &H00000000&
   Caption         =   "Teleporter"
   ClientHeight    =   3480
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7035
   LinkTopic       =   "Form1"
   ScaleHeight     =   3480
   ScaleWidth      =   7035
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "The Evil Scientist is hidden in one of the levels.  Where would you like to search first?"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   975
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   6255
   End
End
Attribute VB_Name = "frmTeleporter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
