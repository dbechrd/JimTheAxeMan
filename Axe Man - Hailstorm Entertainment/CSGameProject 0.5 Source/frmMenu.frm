VERSION 5.00
Begin VB.Form frmMenu 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Jim the Axe Man"
   ClientHeight    =   9000
   ClientLeft      =   3390
   ClientTop       =   2775
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9000
   ScaleWidth      =   12000
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Bell MT"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   615
      Left            =   4920
      TabIndex        =   2
      Top             =   6600
      Width           =   2055
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Load Game"
      BeginProperty Font 
         Name            =   "Bell MT"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   615
      Left            =   4920
      TabIndex        =   1
      Top             =   5640
      Width           =   2055
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Start Game"
      BeginProperty Font 
         Name            =   "Bell MT"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   615
      Left            =   4920
      TabIndex        =   0
      Top             =   4680
      Width           =   2055
   End
   Begin VB.Image imgArrow 
      Height          =   240
      Left            =   4200
      Picture         =   "frmMenu.frx":0000
      Top             =   4800
      Width           =   750
   End
   Begin VB.Image Image4 
      Height          =   2550
      Left            =   1920
      Picture         =   "frmMenu.frx":09C2
      Top             =   600
      Width           =   7950
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub PlayGame()
    
    'MsgBox ("Starting...")
    Unload Me
    frmStoryline.Show
    
End Sub

Private Sub LoadGame()
    
    MsgBox ("Ha, you wish!  Start at the beginning like everyone else.")
    Dim linesofcode As Integer
    linesofcode = 120 + 319 + 201 + 112 + 285 + 93 + 308 + 63 + 1197 + 1183 + 1183 + 1183 + 1183 + 1179 + 1194 + 1181 + 1181 + 1181 + 1181 + 1181 + 1181 + 1179 + 1239 + 1196 + 1195 + 1183 + 1183 + 1179 + 1191 + 1186 + 1183 + 1183 + 1183 + 1177
    MsgBox ("This program containes: " & linesofcode & " lines of code.")
    
End Sub

Private Sub QuitGame()
    
    'MsgBox ("Quitting...")
    Unload Me
    
End Sub

Public Sub Form_Load()
    ArrowPos = 1
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
        If ArrowPos = 1 Then
            imgArrow.Top = imgArrow.Top + 960
            ArrowPos = 2
        ElseIf ArrowPos = 2 Then
            imgArrow.Top = imgArrow.Top + 960
            ArrowPos = 3
        ElseIf ArrowPos = 3 Then
            imgArrow.Top = imgArrow.Top - (960 * 2)
            ArrowPos = 1
        End If
    End If
    
    If KeyCode = vbKeyUp Then
        If ArrowPos = 1 Then
            imgArrow.Top = imgArrow.Top + (960 * 2)
            ArrowPos = 3
        ElseIf ArrowPos = 2 Then
            imgArrow.Top = imgArrow.Top - 960
            ArrowPos = 1
        ElseIf ArrowPos = 3 Then
            imgArrow.Top = imgArrow.Top - 960
            ArrowPos = 2
        End If
    End If
    
    If KeyCode = vbKeyReturn Then
        If ArrowPos = 1 Then
            Call PlayGame
        ElseIf ArrowPos = 2 Then
            Call LoadGame
        ElseIf ArrowPos = 3 Then
            Call QuitGame
        End If
    End If
    
End Sub
