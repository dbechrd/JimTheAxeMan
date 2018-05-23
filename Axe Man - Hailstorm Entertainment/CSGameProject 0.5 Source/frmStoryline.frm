VERSION 5.00
Begin VB.Form frmStoryline 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Jim the Axe Man"
   ClientHeight    =   9000
   ClientLeft      =   2205
   ClientTop       =   1410
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9000
   ScaleWidth      =   12000
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtTest 
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Text            =   "Debug Box"
      Top             =   120
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Timer tmrScroll 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   10800
      Top             =   1320
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Proceed"
      Default         =   -1  'True
      Height          =   495
      Left            =   10200
      TabIndex        =   1
      Top             =   7080
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Quit"
      Height          =   495
      Left            =   10200
      TabIndex        =   2
      Top             =   7680
      Width           =   1575
   End
   Begin VB.Label lblScroll 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Jim The Axe Man"
      BeginProperty Font 
         Name            =   "Bell MT"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   375
      Index           =   0
      Left            =   2160
      TabIndex        =   3
      Top             =   0
      Width           =   7815
   End
   Begin VB.Label lblBackground 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Bell MT"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   7935
      Left            =   2040
      TabIndex        =   0
      Top             =   360
      Width           =   8055
   End
End
Attribute VB_Name = "frmStoryline"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
'Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
'Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
'Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
'Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
'Private Type POINTAPI
'        X As Long
'        Y As Long
'End Type
Dim nFileNum As Integer, sText As String, sNextLine As String, lLineCount As Long, SPAM As Integer, RoflRabbit As Integer, LmaoLlama As Integer, CurScrollY As Integer, LabelQueue() As String
Const ScrollBoxTop = 360
Const ScrollBoxBottom = 7920

Private Sub Command2_Click()

    'If SPAM < 16 Then
    '    Call CreateLabel("Test SPAM#" & SPAM + 1, lblScroll(SPAM).Left, lblScroll(SPAM).Top + lblScroll(SPAM).Height)
    '    SPAM = SPAM + 1
    'End If
        
    'tmrScroll1.Enabled = Not tmrScroll1.Enabled
    '-------------------------
    'THIS PORTION WORKS
    'If tmrScroll.Enabled = False Then
    '    Call InsertStoryline
    'End If
    'END WORKING PORTION
    '-------------------------
    'Call CreateLabel("Test", lblScroll(0).Left, lblScroll(0).Top - lblScroll(0).Height)
    'tmrScroll.Enabled = True
    
    Me.Hide
    frmLevel1a.Show
    
End Sub

Private Sub Form_Load()

    'SetWindowPos Me.hwnd, -1, 100, 100, Me.Width * 0.06, Me.Height * 0.06, 0
    'SetWindowPos Me.hwnd, -1, 100, 100, 806, 632, 0
    'Ret = GetWindowLong(Me.hwnd, -20)
    'Ret = Ret Or &H80000
    'SetWindowLong Me.hwnd, -20, Ret
    'SetLayeredWindowAttributes Me.hwnd, vbBlack, 255, &H2
    
    If tmrScroll.Enabled = False Then
        Call InsertStoryline
    End If
    
    previousLevel = 0
    
End Sub

Private Sub Command1_Click()

    Unload Me
    
End Sub

'THIS FUNCTION WORKS - DO NOT EDIT
Public Function CreateLabel(text As String, X As Integer, Y As Integer)
    
    Dim L As Long, L1
    L = lblScroll.UBound
    L1 = L + 1
    Load lblScroll(L1)
    lblScroll(L1).Move X, Y
    lblScroll(L1).Visible = True
    lblScroll(L1).Caption = text
    lblScroll(L1).ZOrder (vbBringToFront)
    'txtTest.text = txtTest.text & "; " & LmaoLlama & " - " & LabelQueue(LmaoLlama)
    LmaoLlama = LmaoLlama + 1
    'txtTest.text = "DO5"
        
    '======================================
        
    'Dim ctlName As Control
    
    'Set ctlName = frmStoryline.Controls.Add("VB.Label", "Label1", frmStoryline)
        
    'ctlName.Visible = True
        
    'ctlName.Caption = "testing"
        
    'ctlName.Top = 600
        
    'ctlName.Left = 3600
        
    'ctlName.Alignment = 2
        
    'ctlName.BackStyle = 0
        
    'ctlName.ForeColor = &H8000000F
        
    'ctlName.Width = 7815
        
    'ctlName.Height = 375
  
End Function

'Private Sub tmrBackup_Timer()
'
'    Dim text As String
'    text = "Testing Scrolly Thingy :)"
'
'    If lblScroll1.Caption = text Then
'        lblScroll1.Caption = ""
'        lblScroll2.Caption = text
'    ElseIf lblScroll2.Caption = text Then
'        lblScroll2.Caption = ""
'        lblScroll3.Caption = text
'    ElseIf lblScroll3.Caption = text Then
'        lblScroll3.Caption = ""
'        lblScroll4.Caption = text
'    ElseIf lblScroll4.Caption = text Then
'        lblScroll4.Caption = ""
'        lblScroll5.Caption = text
'    ElseIf lblScroll5.Caption = text Then
'        lblScroll5.Caption = ""
'        lblScroll6.Caption = text
'    ElseIf lblScroll6.Caption = text Then
'        lblScroll6.Caption = ""
'        lblScroll7.Caption = text
'    ElseIf lblScroll7.Caption = text Then
'        lblScroll7.Caption = ""
'        lblScroll8.Caption = text
'    ElseIf lblScroll8.Caption = text Then
'        lblScroll8.Caption = ""
'        lblScroll9.Caption = text
'    ElseIf lblScroll9.Caption = text Then
'        lblScroll9.Caption = ""
'        lblScroll10.Caption = text
'    ElseIf lblScroll10.Caption = text Then
'        lblScroll10.Caption = ""
'        lblScroll1.Caption = text
'    End If
'
'End Sub

Private Sub tmrScroll_Timer()
    
    'Do
    '    If lblScroll1.Visible = True Then
    '
    '        Dim CuteLilKittens As Integer
    '        CuteLilKittens = lblScroll1.Top
    '
    '        Dim UglyLilKittens As Integer
    '        UglyLilKittens = 10920
    '
    '        If CuteLilKittens < UglyLilKittens Then
    '            lblScroll1.Top = lblScroll1.Top + 20
    '        Else
    '            lblScroll1.Visible = False
    '            'lblScroll1.Top = 240
    '        End If
    '
    '    End If
    'While RoflRabbit < lblScroll.UBound
    
    'txtTest.text = UBound(LabelQueue)
    
    'If lblScroll(RoflRabbit).Top >= (lblScroll(RoflRabbit).Height + ScrollBoxTop) And LmaoLlama < UBound(LabelQueue) Then
    '    Call CreateLabel(LabelQueue(RoflRabbit), lblScroll(RoflRabbit).Left, lblScroll(RoflRabbit).Top - lblScroll(RoflRabbit).Height)
        'LmaoLlama = LmaoLlama + 1
        'tmrScroll.Enabled = False
    'End If
    
    '==============================================================================
    
    'txtTest.text = "DO7"
    
    Dim CuteLilKittens As Integer
    Dim UglyLilKittens As Integer
    RoflRabbit = 1
    
    Do
        'txtTest.text = "DO1"
        If lblScroll(RoflRabbit).Visible = True Then
            lblScroll(RoflRabbit).Visible = False
        End If
        RoflRabbit = RoflRabbit + 1
    Loop While RoflRabbit <= lblScroll.UBound
    RoflRabbit = 1
    
    'ERROR HERE
    'ERROR HERE
    'ERROR HERE
    Do
        'txtTest.text = "DO2"
        lblScroll(RoflRabbit).Top = lblScroll(RoflRabbit).Top - 10
        If lblScroll(RoflRabbit).Top <= ScrollBoxTop Then
            LabelQueue(RoflRabbit) = "0"
        End If
        RoflRabbit = RoflRabbit + 1
    Loop While RoflRabbit <= lblScroll.UBound
    RoflRabbit = 1
    'ERROR HERE
    'ERROR HERE
    'ERROR HERE
    
    If lblScroll(LmaoLlama).Top <= (ScrollBoxBottom - lblScroll(LmaoLlama).Height) And LmaoLlama < UBound(LabelQueue) Then
        Call CreateLabel(LabelQueue(LmaoLlama), lblScroll(LmaoLlama).Left, lblScroll(LmaoLlama).Top + lblScroll(LmaoLlama).Height)
    End If
    
    Do
        'txtTest.text = "DO3"
        If Not LabelQueue(RoflRabbit) = "0" Then
            lblScroll(RoflRabbit).Visible = True
        End If
        RoflRabbit = RoflRabbit + 1
    Loop While RoflRabbit <= lblScroll.UBound
    RoflRabbit = 1
    
End Sub

Public Function ReadStoryFile() As String

    Dim sText As String

    ' Get a free file number
    nFileNum = FreeFile

    ' Open a text file for input. inputbox returns the path to read the file
    'Open "C:\Documents and Settings\Sudeep\My Documents\3.txt" For Input As nFileNum
    Open "story.txt" For Input As nFileNum
    lLineCount = 1
    ' Read the contents of the file
    Do While Not EOF(nFileNum)

        Line Input #nFileNum, sNextLine
        'do something with it
        'add line numbers to it, in this case!
        sNextLine = sNextLine & vbCrLf
        sText = sText & sNextLine

    Loop
    'Text1.Text = sText
    ReadStoryFile = sText

    ' Close the file
    Close nFileNum

End Function


Public Sub InsertStoryline()

    Dim Storyline As String
    '============================================
    'Dim StorylineSplit() As String
    'Dim TempString As String
    ''TempString = "test "
    'Dim blah As Integer
    'blah = 0
    '============================================
    
    Storyline = ReadStoryFile()
    'LabelQueue() = Split(Storyline, vbCrLf)
    LabelQueue() = Split(Storyline, vbCrLf)
    'txtTest.text = LabelQueue(0) & ", " & lblScroll(0).Left & ", 0"
    Call CreateLabel(LabelQueue(0), lblScroll(0).Left, ScrollBoxBottom)
    '============================================
    'StorylineSplit() = Split(Storyline, vbCrLf)
    '============================================
    
    'txtTest.text = StorylineSplit(1)
    'txtTest.text = ""
    'txtTest.text = UBound(StorylineSplit)
    
    'Do
    '    TempString = TempString & StorylineSplit(blah)
    '    blah = blah + 1
    'Loop While blah < UBound(StorylineSplit)
    'txtTest.text = TempString
    
    '============================================
    'Do
    '    CurScrollY = CurScrollY - lblScroll(0).Height
    '    Call CreateLabel(StorylineSplit(blah), lblScroll(0).Left, CurScrollY)
    '    blah = blah + 1
    'Loop While blah < UBound(StorylineSplit)
    '============================================
    
    'Load lblScroll(0)
    'lblScroll(0).Top = 240 * (0 + 1)
    'lblScroll(0).Caption = "Label #" & "0"
    'lblScroll(0).Visible = True
    
    'THIS ONE
    'Load lblScroll(1)
    'lblScroll(1).Top = 240 * (1 + 1)
    'lblScroll(1).Caption = "Label #" & "1"
    'lblScroll(1).Visible = True
    
    'Dim I As Long
    'For I = 0 To UBound(StorylineSplit)
    '    Load lblScroll(I)
    '    With lblScroll(I)
    '        .Move 240 * (I + 1), 3600
    '        .Caption = "Label #" & I & StorylineSplit(I)
    '        .Visible = True
    '    End With
    'Next
    'txtTest.text = "DO6"
    tmrScroll.Enabled = True

End Sub

