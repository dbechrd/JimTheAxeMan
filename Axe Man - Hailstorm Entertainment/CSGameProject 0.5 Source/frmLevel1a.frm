VERSION 5.00
Begin VB.Form frmLevel1a 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Jim the Axe Man"
   ClientHeight    =   9300
   ClientLeft      =   1770
   ClientTop       =   975
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9300
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrExplode 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   4920
      Top             =   3120
   End
   Begin VB.Timer tmrGame 
      Interval        =   1000
      Left            =   4680
      Top             =   3600
   End
   Begin VB.Timer tmrWalk 
      Interval        =   75
      Left            =   4440
      Top             =   3120
   End
   Begin VB.Timer tmrUserText 
      Enabled         =   0   'False
      Left            =   4200
      Top             =   3600
   End
   Begin VB.Timer tmrGravity 
      Interval        =   5
      Left            =   3960
      Top             =   3120
   End
   Begin VB.CommandButton cmdDebug 
      Caption         =   "Go"
      Default         =   -1  'True
      Height          =   255
      Left            =   2160
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtDebug 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Timer tmrMoveRight 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   3480
      Top             =   3120
   End
   Begin VB.Timer tmrMoveLeft 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   3720
      Top             =   3600
   End
   Begin VB.Label HealthObj 
      Caption         =   "healthbox.bmp"
      Height          =   615
      Index           =   0
      Left            =   6120
      TabIndex        =   10
      Top             =   6720
      Width           =   615
   End
   Begin VB.Label lblEnemyLife 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Enemy Life: 100"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   9960
      TabIndex        =   9
      Top             =   9000
      Width           =   2055
   End
   Begin VB.Label lblCharLife 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Player Life: 100"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   0
      TabIndex        =   8
      Top             =   9000
      Width           =   1935
   End
   Begin VB.Label lblUserText 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1920
      TabIndex        =   7
      Top             =   9000
      Width           =   8055
   End
   Begin VB.Label ChestObj 
      Caption         =   "chest.bmp"
      Height          =   1095
      Index           =   0
      Left            =   2280
      TabIndex        =   6
      Top             =   6720
      Width           =   975
   End
   Begin VB.Label ChestObj 
      Caption         =   "chest.bmp"
      Height          =   1095
      Index           =   1
      Left            =   8040
      TabIndex        =   5
      Top             =   6840
      Width           =   975
   End
   Begin VB.Label Collision 
      Caption         =   "Col1"
      Height          =   495
      Index           =   1
      Left            =   5880
      TabIndex        =   4
      Top             =   7440
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Collision 
      Caption         =   "Col0"
      Height          =   495
      Index           =   0
      Left            =   2280
      TabIndex        =   3
      Top             =   7800
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Image imgAxe1 
      Height          =   705
      Left            =   10440
      Picture         =   "frmLevel1.frx":0000
      Top             =   480
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.Image imgAxe0 
      Height          =   705
      Left            =   10080
      Picture         =   "frmLevel1.frx":0E28
      Top             =   480
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.Label lblTeabag 
      Alignment       =   2  'Center
      BackColor       =   &H80000003&
      BackStyle       =   0  'Transparent
      Caption         =   "Teabag!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10080
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Image Image1 
      Height          =   9000
      Left            =   0
      Picture         =   "frmLevel1.frx":1B31
      Top             =   0
      Visible         =   0   'False
      Width           =   12000
   End
End
Attribute VB_Name = "frmLevel1a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit

Private cResChar As cSpriteBitmaps
Private cChar As cSprite
Private cResEnemy1 As cSpriteBitmaps
Private cEnemy1 As cSprite
Dim cResChestObj() As cSpriteBitmaps
Dim cChestObj() As cSprite
Dim cResHealthObj() As cSpriteBitmaps
Dim cHealthObj() As cSprite

Private cStage As cBitmap
'Excluded Private cBack As cBitmap

' Get Key presses:
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

' General for game:
Private m_bInGame As Boolean

Dim jumpnumber As Integer, jumpdir As Integer, jumpheight As Integer, jumpland As Integer, jumptop As Integer, charJumping As Boolean
Dim Cheat_superjump As Boolean, Cheat_teabag As Boolean, Cheat_speed As Boolean, Cheat_swirl As Boolean
Dim gravityCheck As Integer, charOnPlatform As Boolean
Const groundtop As Integer = 544
Dim AttackTimer As Integer, Blerg As Integer, YouveBeenHit As Integer
Dim ChestJustOpened As Boolean

Private Sub CreateSpriteResource( _
        ByRef cR As cSpriteBitmaps, _
        ByVal sFile As String, _
        ByVal cx As Long, _
        ByVal cy As Long, _
        ByVal lTransColor As Long _
    )
    Set cR = New cSpriteBitmaps
    cR.CreateFromFile sFile, cx, cy, , lTransColor
End Sub

Private Sub CreateSprite( _
        ByRef cR As cSpriteBitmaps, _
        ByRef cS As cSprite _
    )
    Set cS = New cSprite
    cS.SpriteData = cR
    cS.Create Me.hDC
End Sub

Public Sub Form_Load()

    Unload frmStoryline

    '======================================================
    jumpnumber = 0
    'jumpheight = 48
    jumpheight = 80
    'platform1top = 312
    'platform1right = 144
    charJumping = False
    gravityCheck = True
    Blerg = 0
    AttackTimer = 0
    ChestJustOpened = False
    charWalking = 0
    timeleft = 100
    YouveBeenHit = 0
    '======================================================

    Me.Width = 800 * Screen.TwipsPerPixelX
    Me.Height = 620 * Screen.TwipsPerPixelY
    '--Dim i As Integer
    
    ' Create sprites:
    CreateSpriteResource cResChar, App.Path & "\lumberjack.bmp", 19, 1, &HFF00FF
    CreateSprite cResChar, cChar
    cChar.X = 0
    cChar.Y = 0
    
    'CreateSpriteResource cResChar, App.Path & "\Tree1S.bmp", 1, 1, &HFF00FF
    'CreateSprite cResChar, cChar
    'cChar.X = 0
    'cChar.Y = 0

    If (ChestObj.UBound > 0) Then

        ReDim cResChestObj(0 To 0)
        ReDim cChestObj(0 To 0)
    
        Do While i <= ChestObj.UBound
            ReDim Preserve cResChestObj(LBound(cResChestObj) To UBound(cResChestObj) + 1)
            Set cResChestObj(i) = New cSpriteBitmaps
            ReDim Preserve cChestObj(LBound(cChestObj) To UBound(cChestObj) + 1)
            Set cChestObj(i) = New cSprite
            'If Object(i).Caption = "chest" Then
            '    CreateSpriteResource cResChestObj(i), App.Path & "\chest_closed.bmp", 1, 1, &HFF00FF
            'End If
            CreateSpriteResource cResChestObj(i), App.Path & "\" & ChestObj(i).Caption, 2, 1, &HFF00FF
            CreateSprite cResChestObj(i), cChestObj(i)
            cChestObj(i).X = ChestObj(i).Left \ Screen.TwipsPerPixelX
            cChestObj(i).Y = ChestObj(i).Top \ Screen.TwipsPerPixelY
            cChestObj(i).Cell = 1
            i = i + 1
        Loop
        
    End If

    Dim lW As Long, lH As Long
    Set cStage = New cBitmap
    
    lW = Screen.Width \ Screen.TwipsPerPixelX
    lH = Screen.Height \ Screen.TwipsPerPixelY
    cStage.CreateAtSize lW, lH
    
    'Excluded Set cBack = New cBitmap
    'Excluded lW = Screen.Width \ Screen.TwipsPerPixelX
    'Excluded lH = Screen.Height \ Screen.TwipsPerPixelY

    'cBack.CreateFromFile App.Path & "\background1.bmp"
    'Call levelLoad("background1.bmp")
    'cBack.CreateFromFile App.Path & "\sideplat.bmp"
    'Call levelLoad("sideplat.bmp")
    'Excluded cBack.RenderBitmap cStage.hDC, 0, 0
    
    Call showUserText("Level One - The Forest", 3000)
    
    Me.Show
    ' ? Me.Refresh
    GameLoop
    
End Sub

'===========================

Public Sub GameLoop()
    Static bGameLoop As Boolean
    Dim i As Long
    Dim iSpriteNum As Long
    Dim lHDC As Long
    Dim lH As Long
    Dim lW As Long

    bGameLoop = Not (bGameLoop)
    lHDC = Me.hDC
    lW = Me.ScaleWidth \ Screen.TwipsPerPixelX
    lH = Me.ScaleHeight \ Screen.TwipsPerPixelY
    cChar.X = (lW - cChar.Width) \ 2
    cChar.Y = (lH - cChar.Height) \ 2
    cChar.Cell = 1
    
    If (bGameLoop) Then
        cStage.RenderBitmap lHDC, 0, 0
    End If
    m_bInGame = bGameLoop
    
    Do While bGameLoop
        
        ' ******************************************************
        ' 1) Firstly, we restore the stage bitmap to its original
        ' state:
        
        Do While i <= ChestObj.UBound
            cChestObj(i).RestoreBackground cStage.hDC
            i = i + 1
        Loop
        
        If Cheat_swirl = False Then
            cChar.RestoreBackground cStage.hDC
        End If
        
        ' ******************************************************
        
        ' (At this point you could modify the background in cStage)
        If blah = 0 Then
            levelLoad ("sideplat.bmp")
            'cStage.CreateFromFile App.Path & "\background1.bmp"
            'Excluded cBack.RenderBitmap cStage.hDC, 0, 0
                    
            'Me.Show
            'Me.Refresh
            'cStage.RenderBitmap lHDC, 0, 0
            blah = 1
        End If
        
        If YouveBeenHit = 1 Then
            If Blerg = 0 Then
                Call levelLoad("red.bmp")
                Blerg = 1
            ElseIf Blerg = 1 Then
                Call levelLoad("white.bmp")
                Blerg = 0
            End If
            AttackTimer = AttackTimer + 1
            If AttackTimer = 4 Then
                Call levelLoad("sideplat.bmp")
                AttackTimer = 0
                YouveBeenHit = 0
            End If
        End If
        
        
        ' ******************************************************
        ' 2) Secondly, we move all the sprites to their new position
        ' on the stage bitmap and copy the stage at that point:
        
        'Character:
        'gravityCheck = False
        If (GetAsyncKeyState(vbKeyUp) <> 0) And (jumpnumber = 0) Then
            'jump
            charWalking = 0
            jumpland = cChar.Y
            jumptop = cChar.Y - jumpheight
            gravityCheck = False
            jumpdir = 1
            jumpnumber = 1
            charJumping = True
            
                'Move Character
                '-cChar.Y = cChar.Y - 1
                'Check for collision
                '-If charCollision = False Then
                
                    'Move Teabag Label
                    If lblTeabag.Visible = True Then
                        lblTeabag.Top = lblTeabag.Top - 1
                    End If
                    
                    'Move Axe
                    If imgAxe0.Visible = True Then
                        imgAxe0.Top = cChar.Y + 480
                    ElseIf imgAxe1.Visible = True Then
                        imgAxe1.Top = cChar.Y + 480
                    End If
                
                '-Else
                    'tmrMoveRight.Enabled = False
                    'cChar.Y = cChar.Y + 1
                    'MsgBox ("Collide Up")
                '-End If
        ElseIf (GetAsyncKeyState(vbKeyRight) <> 0) Then
            'move right
            Call charMoveRight
        ElseIf (GetAsyncKeyState(vbKeyLeft) <> 0) Then
            'move left
            Call charMoveLeft
        'ElseIf (GetAsyncKeyState(vbKeyDown) <> 0) Then
        '    'crouch
        '    '--cChar.Y = cChar.Y + 1
        '    If charCollision = False Then
        '
        '        cChar.Y = cChar.Y + 1
        '        If lblTeabag.Visible = True Then
        '            lblTeabag.Top = lblTeabag.Top + 1
        '        End If
        '        If imgAxe0.Visible = True Then
        '            imgAxe0.Top = cChar.Y + 480
        '        ElseIf imgAxe1.Visible = True Then
        '            imgAxe1.Top = cChar.Y + 480
        '        End If
        '
        '    Else
        '        'MsgBox ("Collide Down")
        '    End If
        'ElseIf (GetAsyncKeyState(vbKeyF11) <> 0) Then
        '    txtDebug.Visible = True
        '    cmdDebug.Visible = True
        ElseIf (GetAsyncKeyState(vbKeySpace) <> 0) Then
            Dim ObjCollidedWith As Integer
            ObjCollidedWith = charChestObjCollision()
            If ObjCollidedWith >= 0 Then
                'DEBUG
                'MsgBox ("Chest Collision")
                If (cChestObj(ObjCollidedWith).Cell = 1) Then
                    cChestObj(ObjCollidedWith).Cell = 2
                    Call levelLoad("sideplat.bmp")
                    'EXPLODE
                    tmrExplode.Enabled = True
                    'Call showUserText("The chest was empty.", 2000)
                    Call showUserText("The chest contained a pipe bomb.", 3500)
                    ChestJustOpened = True
                ElseIf (ChestJustOpened = False) Then
                    'Call showUserText("You have already looted this chest.  Nothing remains.", 3000)
                    Call showUserText("You peek in and see some of your past life's charred remains.", 4000)
                    'Call showUserText("Greedy Bastard.", 2000)
                End If
            End If
        ElseIf (GetAsyncKeyState(vbKeyF10) <> 0) Then
            'tmrExplode.Enabled = True
            
            'If (cChar.Cell < 6) Then
            '    cChar.Cell = cChar.Cell + 1
            'Else
            '    cChar.Cell = 1
            'End If
            
            'MsgBox (cChar.X)
            'MsgBox (cChar.Width)
        '    If (blah = 1) Then
        '        'MsgBox ("Loading sideplat.bmp")
        '        Call levelLoad("sideplat.bmp")
        '
        '        blah = 2
        '        'cStage.RenderBitmap lHDC, 0, 0
        '    ElseIf (blah = 2) Then
        '        'MsgBox ("Loading background1.bmp")
        '        Call levelLoad("background1.bmp")
        '
        '        blah = 1
        '        'cStage.RenderBitmap lHDC, 0, 0
        '    End If
        '    'Call levelLoad("background1.bmp")
        '    'cStage.RenderBitmap lHDC, 0, 0
        ElseIf (GetAsyncKeyState(vbKeyF12) <> 0) Then
            YouveBeenHit = 1
        '    If (blah = 1) Then
        '        'MsgBox ("Loading sideplat.bmp")
        '        Call levelLoad("sideplat.bmp")
        '
        '        blah = 2
        '        'cStage.RenderBitmap lHDC, 0, 0
        '    ElseIf (blah = 2) Then
        '        'MsgBox ("Loading background1.bmp")
        '        Call levelLoad("background1.bmp")
        '
        '        blah = 1
        '        'cStage.RenderBitmap lHDC, 0, 0
        '    End If
        '    'Call levelLoad("background1.bmp")
        '    'cStage.RenderBitmap lHDC, 0, 0
        ElseIf (GetAsyncKeyState(vbKeyEscape) <> 0) Then
        '    ' accelerate in current direction:
            Unload Me
            Exit Sub
        ElseIf (GetAsyncKeyState(vbKeyRight) = 0 And GetAsyncKeyState(vbKeyLeft) = 0) Then
            charWalking = 0
        End If
        
        If charJumping = True Then
            Call charJump
        End If
        
        If gravityCheck = True Then
            Call gravity
        End If
        
        'Animation:
        
        'Objects:
        Dim h As Integer
        For h = 0 To ChestObj.UBound Step 1
            cChestObj(h).StoreBackground cStage.hDC, cChestObj(h).X, cChestObj(h).Y
        Next h
        
        'Enemies:
        
        cChar.StoreBackground cStage.hDC, cChar.X, cChar.Y
        If Cheat_swirl = True Then
            'Excluded cChar.RestoreBackground cBack.hDC
        End If
        
        ' ******************************************************
        
        ' ******************************************************
        ' 3) Next we draw all the sprites onto the stage:
        
        Dim g As Integer
        For g = 0 To ChestObj.UBound Step 1
            cChestObj(g).TransparentDraw cStage.hDC, cChestObj(g).X, cChestObj(g).Y, cChestObj(g).Cell, False
        Next g
        
        cChar.TransparentDraw cStage.hDC, cChar.X, cChar.Y, cChar.Cell, False

        
        ' ******************************************************

        ' ******************************************************
        
        ' ******************************************************
        ' 3) Finally we transfer the changes in the stage onto
        ' the screen, minimising the number of visible screen
        ' blits as best as we can:
        ' ? cChar.StageToScreen lHDC, cStage.hDC
        
        ' ******************************************************
        If Cheat_speed = False Then
            'Excluded cBack.RenderBitmap cStage.hDC, 0, 0
        End If
        cStage.RenderBitmap lHDC, 0, 0
        
        DoEvents
    Loop
        
End Sub

Public Sub levelLoad(FileName As String)
'Public Sub levelLoad()

    'Dim lW As Long, lH As Long
    'Set cStage = New cBitmap
    
    'lW = Screen.Width \ Screen.TwipsPerPixelX
    'lH = Screen.Height \ Screen.TwipsPerPixelY
    'cStage.CreateAtSize lW, lH
    
    'Set cBack = New cBitmap
    'lW = Screen.Width \ Screen.TwipsPerPixelX
    'lH = Screen.Height \ Screen.TwipsPerPixelY

    'cBack.CreateFromFile App.Path & "\" & FileName
    'cBack.RenderBitmap cStage.hDC, 0, 0
    
    'Me.Show
    'Me.Refresh
    
    'cBack.RenderBitmap cStage.hDC, 0, 0
    '====================================================
    
    cStage.CreateFromFile App.Path & "\" & FileName
    'Excluded cBack.RenderBitmap cStage.hDC, 0, 0
                    
    ' ? Me.Show
    ' ? Me.Refresh
    cStage.RenderBitmap Me.hDC, 0, 0
    
End Sub

Public Sub objectLoad(FileName As String)
    
    cResChestObj.CreateFromFile App.Path & "\" & FileName
    cResChestObj.RenderBitmap Me.hDC, 0, 0
    
End Sub

'===========================

Public Sub gravity()
    If charCollision = False Then
        'charOnPlatform = False
        If cChar.Y < groundtop - cChar.Height Then
            cChar.Y = cChar.Y + 1
        ElseIf cChar.Y > groundtop - cChar.Height Then
            cChar.Y = groundtop - cChar.Height
        Else
            jumpnumber = 0
        End If
    Else
        'charOnPlatform = True
        jumpnumber = 0
    End If
End Sub

'===========================

Public Sub charJump()
    
    If jumpdir = 1 Then
        If cChar.Y >= 0 Then
            If cChar.Y > jumptop Then
                cChar.Y = cChar.Y - 1
            ElseIf cChar.Y <= jumptop Then
                'MsgBox (jumptop & " " & cChar.Y)
                jumpdir = 2
            End If
        ElseIf cChar.Y < 0 Then
            jumpdir = 2
        End If
    ElseIf jumpdir = 2 Then
        gravityCheck = True
        charJumping = False
        jumpdir = 1
    End If
    'If imgAxe0.Visible = True Then
    '    imgAxe0.Top = imgCharacter.Top
    'ElseIf imgAxe1.Visible = True Then
    '    imgAxe1.Top = imgCharacter.Top
    'End If
    
End Sub

Public Sub charMoveRight()

    charWalking = 1
    cChar.X = cChar.X + 1
    'If charCollision = False Or (charCollision = True And charOnPlatform = True) Then
    '
    '    If lblTeabag.Visible = True Then
    '        lblTeabag.Left = lblTeabag.Left + 1
    '    End If
    '    If imgAxe0.Visible = True Then
    '        imgAxe0.Left = cChar.X + 480
    '    ElseIf imgAxe1.Visible = True Then
    '        imgAxe1.Left = cChar.X + 480
    '    End If
    
    If charCollision = False Then
        
        'MsgBox ("1")
        If lblTeabag.Visible = True Then
            lblTeabag.Left = lblTeabag.Left + 1
        End If
        If imgAxe0.Visible = True Then
            imgAxe0.Left = cChar.X + 480
        ElseIf imgAxe1.Visible = True Then
            imgAxe1.Left = cChar.X + 480
        End If
        
    ElseIf (charCollision = True And charOnPlatform = True) Then
        
        'MsgBox ("Platform = True ;)")
        If lblTeabag.Visible = True Then
            lblTeabag.Left = lblTeabag.Left + 1
        End If
        If imgAxe0.Visible = True Then
            imgAxe0.Left = cChar.X + 480
        ElseIf imgAxe1.Visible = True Then
            imgAxe1.Left = cChar.X + 480
        End If
    
    Else
        'Collision move char back
        cChar.X = cChar.X - 1
        charWalking = 0
    End If

End Sub

Public Sub charMoveLeft()

    charWalking = 2
    cChar.X = cChar.X - 1
    If charCollision = False Or (charCollision = True And charOnPlatform = True) Then
                
        If lblTeabag.Visible = True Then
            lblTeabag.Left = lblTeabag.Left - 1
        End If
        If imgAxe0.Visible = True Then
            imgAxe0.Left = cChar.X + 480
        ElseIf imgAxe1.Visible = True Then
            imgAxe1.Left = cChar.X + 480
        End If
                        
    Else
        'Collision move char back
        cChar.X = cChar.X + 1
        charWalking = 0
    End If

End Sub

Public Function charCollision() As Boolean

    'MsgBox ("False")
    charOnPlatform = False
    Dim i As Integer
    If cChar.X + cChar.Width > 800 Or cChar.X < 0 Then
        'MsgBox ("BLAH U WIN")
        charCollision = True
    End If
        Do While i <= Collision.UBound
            'Left-Right Collision
            If (cChar.X * Screen.TwipsPerPixelX) < Collision(i).Left + Collision(i).Width And (cChar.X * Screen.TwipsPerPixelX) + (cChar.Width * Screen.TwipsPerPixelX) > Collision(i).Left Then
                'Top-Bottom Collision
                'DEBUG THIS
                If (cChar.Y * Screen.TwipsPerPixelY) < Collision(i).Top + Collision(i).Height And (cChar.Y * Screen.TwipsPerPixelY) + (cChar.Height * Screen.TwipsPerPixelY) > Collision(i).Top Then
                    If (cChar.Y * Screen.TwipsPerPixelY) + (cChar.Height * Screen.TwipsPerPixelY) = Collision(i).Top + 15 Then
                        charOnPlatform = True
                        'MsgBox ("Please Work..")
                    Else
                        'MsgBox ((cChar.Y * Screen.TwipsPerPixelY) + (cChar.Height * Screen.TwipsPerPixelY))
                        'MsgBox (Collision(i).Top)
                        '= MsgBox ("False")
                        '= charOnPlatform = False
                        'MsgBox ("Char on platformz")
                    End If
                    charCollision = True
                End If
            End If
            i = i + 1
        Loop
        
End Function

Public Function charChestObjCollision() As Integer
    
    charChestObjCollision = -1
    Dim i As Integer
    Do While i <= ChestObj.UBound
        'Left-Right Collision
        If (cChar.X * Screen.TwipsPerPixelX) < ChestObj(i).Left + ChestObj(i).Width And (cChar.X * Screen.TwipsPerPixelX) + (cChar.Width * Screen.TwipsPerPixelX) > ChestObj(i).Left Then
            'Top-Bottom Collision
            If (cChar.Y * Screen.TwipsPerPixelY) < ChestObj(i).Top + ChestObj(i).Height And (cChar.Y * Screen.TwipsPerPixelY) + (cChar.Height * Screen.TwipsPerPixelY) > ChestObj(i).Top Then
                'If (cChar.Y * Screen.TwipsPerPixelY) + (cChar.Height * Screen.TwipsPerPixelY) = Object(i).Top + 15 Then
                    charChestObjCollision = i
                'End If
            End If
        End If
        i = i + 1
    Loop
    
End Function

Private Sub tmrGame_Timer()
    
    lblEnemyLife.Caption = "Enemy Life: " & timeleft
    timeleft = timeleft - 1
    'Call showUserText("Enemy Life: " & timeleft, 300)
    
End Sub

Private Sub tmrWalk_Timer()

    If (charWalking = 1) Then
        If (cChar.Cell < 6) Then
            cChar.Cell = cChar.Cell + 1
        Else
            cChar.Cell = 1
        End If
    ElseIf (charWalking = 2) Then
        If (cChar.Cell < 12 And cChar.Cell > 6) Then
            cChar.Cell = cChar.Cell + 1
        Else
            cChar.Cell = 7
        End If
    ElseIf (charWalking = 0) Then
    End If

End Sub

Private Sub tmrExplode_Timer()

    If (cChar.Cell < 13) Then
        cChar.Cell = 13
        YouveBeenHit = 1
    ElseIf (cChar.Cell < 19) Then
        cChar.Cell = cChar.Cell + 1
    Else
        tmrExplode.Enabled = False
        'cChar.Cell = 1
    End If

End Sub

Public Function showUserText(text As String, Interval As Integer)
    
    lblUserText.Caption = text
    tmrUserText.Interval = Interval
    lblUserText.ZOrder (0)
    lblUserText.Visible = True
    tmrUserText.Enabled = True
    
End Function

Private Sub tmrUserText_Timer()

    lblUserText.Visible = False
    tmrUserText.Enabled = False
    ChestJustOpened = False
    
End Sub

Private Sub cmdDebug_Click()
    Call DoCheat
End Sub

Public Function DoCheat()

    If txtDebug.text = "superjump" Then
        If Cheat_superjump = False Then
            Cheat_superjump = True
            jumpheight = groundtop - cChar.Height
            MsgBox ("0x1001" & jumpheight)
        ElseIf Cheat_superjump = True Then
            Cheat_superjump = False
            jumpheight = 48
            MsgBox ("0x0001")
        End If
    ElseIf txtDebug.text = "teabag" Then
        If Cheat_teabag = False Then
            Cheat_teabag = True
            MsgBox ("0x1002")
        ElseIf Cheat_teabag = True Then
            Cheat_teabag = False
            MsgBox ("0x0002")
        End If
    ElseIf txtDebug.text = "speed" Then
        If Cheat_speed = False Then
            Cheat_speed = True
        ElseIf Cheat_speed = True Then
            Cheat_speed = False
        End If
    ElseIf txtDebug.text = "swirl" Then
        If Cheat_swirl = False Then
            Cheat_swirl = True
        ElseIf Cheat_swirl = True Then
            Cheat_swirl = False
        End If
    Else
        MsgBox ("0x0000")
    End If
    txtDebug.text = ""
    txtDebug.Visible = False
    cmdDebug.Visible = False

End Function

Public Function CharAttack(AttackType As Integer, KeyPos As Integer)
    Select Case AttackType
        Case 0
            If KeyPos Then
                If imgAxe0.Visible = False Then
                    If imgAxe1.Visible = True Then
                        imgAxe1.Visible = False
                    End If
                    imgAxe0.Top = imgCharacter.Top
                    imgAxe0.Left = imgCharacter.Left + 480
                    imgAxe0.Visible = True
                End If
            Else
                imgAxe0.Visible = False
            End If
        Case 1
            If KeyPos Then
                If imgAxe1.Visible = False Then
                    If imgAxe0.Visible = True Then
                        imgAxe0.Visible = False
                    End If
                    imgAxe1.Top = imgCharacter.Top
                    imgAxe1.Left = imgCharacter.Left + 480
                    imgAxe1.Visible = True
                End If
            Else
                imgAxe1.Visible = False
            End If
    End Select
End Function

'=================

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If (m_bInGame) Then
        GameLoop
    End If
End Sub
