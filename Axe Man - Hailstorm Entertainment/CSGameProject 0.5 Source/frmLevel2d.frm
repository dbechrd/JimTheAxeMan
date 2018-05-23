VERSION 5.00
Begin VB.Form frmLevel2d 
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
   Begin VB.Timer tmrMoveLeft 
      Enabled         =   0   'False
      Interval        =   5
      Left            =   2760
      Top             =   3000
   End
   Begin VB.Timer tmrMoveRight 
      Enabled         =   0   'False
      Interval        =   5
      Left            =   2760
      Top             =   2520
   End
   Begin VB.Timer tmrGravity 
      Interval        =   5
      Left            =   3240
      Top             =   2520
   End
   Begin VB.Timer tmrUserText 
      Enabled         =   0   'False
      Left            =   3240
      Top             =   3000
   End
   Begin VB.Timer tmrWalk 
      Enabled         =   0   'False
      Interval        =   75
      Left            =   3720
      Top             =   2520
   End
   Begin VB.Timer tmrGame 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   3720
      Top             =   3000
   End
   Begin VB.Timer tmrExplode 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   4200
      Top             =   2520
   End
   Begin VB.Timer tmrEnemyGenerator 
      Interval        =   4000
      Left            =   4200
      Top             =   3000
   End
   Begin VB.Timer tmrEnemyMove 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   4680
      Top             =   3000
   End
   Begin VB.Timer tmrEnemyAttack 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   4680
      Top             =   2520
   End
   Begin VB.Timer tmrCharAttack 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   5160
      Top             =   2520
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
   Begin VB.Label Collision 
      Caption         =   "Col1"
      Height          =   375
      Index           =   3
      Left            =   8160
      TabIndex        =   12
      Top             =   4320
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.Label Collision 
      Caption         =   "Col1"
      Height          =   375
      Index           =   2
      Left            =   5280
      TabIndex        =   11
      Top             =   5400
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Level 1 b"
      Height          =   615
      Left            =   3360
      TabIndex        =   10
      Top             =   1440
      Width           =   1935
   End
   Begin VB.Label ItemObj 
      Caption         =   "healthbox.bmp"
      Height          =   615
      Index           =   1
      Left            =   4560
      TabIndex        =   9
      Top             =   7200
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
      TabIndex        =   8
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
      TabIndex        =   7
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
      TabIndex        =   6
      Top             =   9000
      Width           =   8055
   End
   Begin VB.Label ItemObj 
      Caption         =   "chest.bmp"
      Height          =   1095
      Index           =   0
      Left            =   10200
      TabIndex        =   5
      Top             =   6600
      Width           =   975
   End
   Begin VB.Label Collision 
      Caption         =   "Col1"
      Height          =   375
      Index           =   1
      Left            =   3000
      TabIndex        =   4
      Top             =   6240
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label Collision 
      Caption         =   "Col0"
      Height          =   375
      Index           =   0
      Left            =   840
      TabIndex        =   3
      Top             =   6960
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Image imgAxe1 
      Height          =   705
      Left            =   10440
      Picture         =   "frmLevel2d.frx":0000
      Top             =   480
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.Image imgAxe0 
      Height          =   705
      Left            =   10080
      Picture         =   "frmLevel2d.frx":0E28
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
      Picture         =   "frmLevel2d.frx":1B31
      Top             =   0
      Visible         =   0   'False
      Width           =   12000
   End
End
Attribute VB_Name = "frmLevel2d"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit

'-Private cResChar As cSpriteBitmaps
'-Private cChar As cSprite
'-Private cResEnemy1 As cSpriteBitmaps
'-Private cEnemy1 As cSprite
'-Dim cResItemObj() As cSpriteBitmaps
'-Dim cItemObj() As cSprite
'-Dim cResHealthObj() As cSpriteBitmaps
'-Dim cHealthObj() As cSprite

'Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private cStage As cBitmap
'Excluded Private cBack As cBitmap

' Get Key presses:
'-Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

' General for game:
'-Private m_bInGame As Boolean

Dim jumpnumber As Integer, jumpdir As Integer, jumpheight As Integer, jumpland As Integer, jumptop As Integer, charJumping As Boolean, charAttacking As Boolean
Dim Cheat_superjump As Boolean, Cheat_teabag As Boolean, Cheat_speed As Boolean, Cheat_swirl As Boolean
Dim gravityCheck As Integer, charOnPlatform As Boolean
Const groundtop As Integer = 544
Dim AttackTimer As Integer, Blerg As Integer, YouveBeenHit As Integer
Dim ItemBeingUsed As Boolean
Dim ScreenNeedsUpdate As Boolean, HasAttackedThisKeypress As Boolean
Dim charAnimPos As Integer, enemy1AnimPos As Integer
Dim enemyExists As Boolean, enemy1Walking As Integer

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

    'NEWFORM
    If (previousLevel = 2.3) Then
        Unload frmLevel1c
    ElseIf (previousLevel = 2.5) Then
        Unload frmLevel1e
    End If
    'If (frmLevel1b.Visible) Then
    '    frmLevel1b.Hide
    'End If
    Randomize
    '======================================================
    jumpnumber = 0
    'jumpheight = 48
    jumpheight = 80
    'platform1top = 312
    'platform1right = 144
    charJumping = False
    charAttacking = False
    gravityCheck = True
    Blerg = 0
    AttackTimer = 0
    charWalking = 0
    enemy1Walking = 0
    timeleft = 100
    YouveBeenHit = 0
    ItemBeingUsed = False
    ScreenNeedsUpdate = True
    enemyExists = False
    '======================================================

    Me.Width = 800 * Screen.TwipsPerPixelX
    Me.Height = 620 * Screen.TwipsPerPixelY
    '--Dim i As Integer
    
    ' Create sprites:
    CreateSpriteResource cResChar, App.Path & "\lumberjack.bmp", 21, 1, &HFF00FF
    CreateSprite cResChar, cChar
    
    CreateSpriteResource cResEnemy1, App.Path & "\Tree2.bmp", 2, 1, &HFF00FF
    CreateSprite cResEnemy1, cEnemy1
    
    'cChar.X = 0
    'cChar.Y = 0
    
    'CreateSpriteResource cResChar, App.Path & "\Tree1S.bmp", 1, 1, &HFF00FF
    'CreateSprite cResChar, cChar
    'cChar.X = 0
    'cChar.Y = 0

    If (ItemObj.UBound > 0) Then

        ReDim cResItemObj(0 To 0)
        ReDim cItemObj(0 To 0)
    
        Do While i <= ItemObj.UBound
            If Level2x1_IntroText = False Then
                Call showUserText("Level Two - Northrend", 3000)
                charHealth = 80
                enemy1Health = 100
                points = 0
                timeleft = 7314
                'NEWFORM
                Level2x1_IntroText = True
            End If
            'NEWFORM
            If Level2x4_IntroText = False Then
                Dim DeclareItemArray As Integer
                'NEWFORM
                ReDim Level2x4_Items(0 To i)
                'NEWFORM
                For DeclareItemArray = 0 To UBound(Level2x4_Items)
                    'NEWFORM
                    Level2x4_Items(DeclareItemArray).FileName = ItemObj(DeclareItemArray).Caption
                    Level2x4_Items(DeclareItemArray).ItemType = Mid(ItemObj(DeclareItemArray).Caption, 1, Len(ItemObj(DeclareItemArray).Caption) - 4)
                    'Level1x1_Items(DeclareItemArray).ItemType = ItemObj(DeclareItemArray).Caption
                    Level2x4_Items(DeclareItemArray).Used = False
                Next DeclareItemArray
            End If
            ReDim Preserve cResItemObj(LBound(cResItemObj) To UBound(cResItemObj) + 1)
            Set cResItemObj(i) = New cSpriteBitmaps
            ReDim Preserve cItemObj(LBound(cItemObj) To UBound(cItemObj) + 1)
            Set cItemObj(i) = New cSprite
            'If Object(i).Caption = "chest" Then
            '    CreateSpriteResource cResItemObj(i), App.Path & "\chest_closed.bmp", 1, 1, &HFF00FF
            'End If
            'NEWFORM
            If (Level2x4_Items(i).Used = False) Then
                CreateSpriteResource cResItemObj(i), App.Path & "\" & ItemObj(i).Caption, 2, 1, &HFF00FF
                CreateSprite cResItemObj(i), cItemObj(i)
                cItemObj(i).X = ItemObj(i).Left \ Screen.TwipsPerPixelX
                cItemObj(i).Y = ItemObj(i).Top \ Screen.TwipsPerPixelY
                cItemObj(i).Cell = 1
            'NEWFORM
            ElseIf (Level2x4_Items(i).Used = True) Then
                CreateSpriteResource cResItemObj(i), App.Path & "\" & ItemObj(i).Caption, 2, 1, &HFF00FF
                CreateSprite cResItemObj(i), cItemObj(i)
                cItemObj(i).X = ItemObj(i).Left \ Screen.TwipsPerPixelX
                cItemObj(i).Y = ItemObj(i).Top \ Screen.TwipsPerPixelY
                cItemObj(i).Cell = 2
            End If
            i = i + 1
        Loop
        'NEWFORM
        Level2x4_IntroText = True
        
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
    
    lblCharLife.Caption = "Player Life: " & charHealth
    lblEnemyLife.Caption = "Enemy Life: " & enemy1Health
    '-Money-lblEnemyLife.Caption = "Moneyz: " & money
    
    Me.Show
    ' ? Me.Refresh
    GameLoop
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'NEWFORM
    previousLevel = 2.4
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
    'cChar.X = (lW - cChar.Width) \ 2
    'cChar.Y = (lH - cChar.Height) \ 2
    'Dim tempbug As String
    'tempbug = previousLevel
    'MsgBox ("previous level: " & previousLevel)
    'NEWFORM
    If (previousLevel = 2.3) Then
        cChar.X = 30
        cChar.Y = (groundtop - cChar.Height)
        cChar.Cell = 1
    ElseIf (previousLevel = 2.5) Then
        cChar.X = (lW - cChar.Width - 180)
        cChar.Y = (groundtop - cChar.Height)
        cChar.Cell = 8
    End If
    
    If (bGameLoop) Then
        cStage.RenderBitmap lHDC, 0, 0
    End If
    m_bInGame = bGameLoop
    
    Do While bGameLoop
        
        ' ******************************************************
        ' 1) Firstly, we restore the stage bitmap to its original
        ' state:
        
        If (charHealth <= 0) Then
            MsgBox ("Pause")
            Unload Me
            Exit Sub
        End If
        
        If (enemyExists = True And enemy1Health <= 0) Then
            ScreenNeedsUpdate = True
            enemyExists = False
            tmrEnemyMove.Enabled = False
            tmrEnemyAttack.Enabled = False
            tmrCharAttack.Enabled = False
            tmrEnemyGenerator.Enabled = True
            'NEWFORM
            Call levelLoad("level2d.bmp")
            Call showUserText("You defeated Tree.", 2000)
        End If
        
        'If (keypressed = True) Then
            Do While i <= ItemObj.UBound
                cItemObj(i).RestoreBackground cStage.hDC
                i = i + 1
            Loop
            
            If Cheat_swirl = False Then
                cChar.RestoreBackground cStage.hDC
                If enemyExists = True Then
                    cEnemy1.RestoreBackground cStage.hDC
                End If
            End If
        'End If
        
        ' ******************************************************
        ' jake is the sickest man in the world
        ' i love jake so much
        ' jake is my god
        ' (At this point you could modify the background in cStage)
        If blah = 0 Then
            ScreenNeedsUpdate = True
            'NEWFORM
            levelLoad ("level2d.bmp")
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
            If (charHealth > 0) Then
                charHealth = charHealth - 1
                lblCharLife.Caption = "Player Life: " & charHealth
            End If
            AttackTimer = AttackTimer + 1
            If AttackTimer = 4 Then
                'NEWFORM
                Call levelLoad("level2d.bmp")
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
            ScreenNeedsUpdate = True
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
            ScreenNeedsUpdate = True
            'move right
            Call charMoveRight
        ElseIf (GetAsyncKeyState(vbKeyLeft) <> 0) Then
            ScreenNeedsUpdate = True
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
        ElseIf (GetAsyncKeyState(vbKeyA) <> 0) Then
            If (HasAttackedThisKeypress = False) Then
                ScreenNeedsUpdate = True
                charAttacking = True
                Call charAnimate
            ElseIf (HasAttackedThisKeypress = True) Then
                charAttacking = False
            End If
            'Call charAttack("melee")
        ElseIf (GetAsyncKeyState(vbKeySpace) <> 0) Then
            ScreenNeedsUpdate = True
            If (ItemBeingUsed = False) Then
                ItemBeingUsed = True
                Dim ObjCollidedWith As Integer
                ObjCollidedWith = charObjCollision()
                'MsgBox ("Obj Collided With: " & ObjCollidedWith)
                If ObjCollidedWith >= 0 Then
                    UseItem (ObjCollidedWith)
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
        ElseIf (GetAsyncKeyState(vbKeyF11) <> 0) Then
            ScreenNeedsUpdate = True
            txtDebug.Visible = True
            cmdDebug.Visible = True
        ElseIf (GetAsyncKeyState(vbKeyF12) <> 0) Then
            ScreenNeedsUpdate = True
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
            'Exit Sub
            'Call UnloadAllForms
            Exit Sub
        End If
        
        If (GetAsyncKeyState(vbKeyA) = 0) Then
            If (charAttacking = True) Then
                charAttacking = False
                Call charAnimate
                ScreenNeedsUpdate = True
            End If
            HasAttackedThisKeypress = False
        ElseIf (GetAsyncKeyState(vbKeyRight) = 0 And GetAsyncKeyState(vbKeyLeft) = 0) Then
            charWalking = 0
        End If
        
        If charJumping = True Then
            Call charJump
            ScreenNeedsUpdate = True
        End If
        
        If gravityCheck = True Then
            Call gravity
        End If
        
        If enemyExists = True And tmrEnemyMove.Enabled = False Then
            tmrEnemyMove.Enabled = True
            tmrEnemyAttack.Enabled = True
            tmrCharAttack.Enabled = True
        End If
        
        lblCharLife.Caption = "Player Life: " & charHealth
        lblEnemyLife.Caption = "Enemy Life: " & enemy1Health
        '-Money-lblEnemyLife.Caption = "Moneyz: " & money
        'lblTime.Caption = timeleft
        'lblTime.Visible = False
        'lblTime.Visible = True
        
        'Animation:
        
        DoEvents
        
        If (ScreenNeedsUpdate = True) Then
            
            'Objects:
            Dim h As Integer
            For h = 0 To ItemObj.UBound Step 1
                cItemObj(h).StoreBackground cStage.hDC, cItemObj(h).X, cItemObj(h).Y
            Next h
            
            'Enemies:
            
            cChar.StoreBackground cStage.hDC, cChar.X, cChar.Y
            If enemyExists = True Then
                cEnemy1.StoreBackground cStage.hDC, cEnemy1.X, cEnemy1.Y
            End If
            If Cheat_swirl = True Then
                'Excluded cChar.RestoreBackground cBack.hDC
            End If
            
            ' ******************************************************
            
            ' ******************************************************
            ' 3) Next we draw all the sprites onto the stage:
            
            Dim g As Integer
            For g = 0 To ItemObj.UBound Step 1
                cItemObj(g).TransparentDraw cStage.hDC, cItemObj(g).X, cItemObj(g).Y, cItemObj(g).Cell, False
            Next g
            
            cChar.TransparentDraw cStage.hDC, cChar.X, cChar.Y, cChar.Cell, False
            If enemyExists = True Then
                cEnemy1.TransparentDraw cStage.hDC, cEnemy1.X, cEnemy1.Y, cEnemy1.Cell, False
            End If
            
            
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
            
            ScreenNeedsUpdate = False
        End If
        
        'DoEvents
        
        'Sleep 1
        
    Loop
        
End Sub

Public Sub DrawScreen(state As Integer)
    If (state = 0) Then
        
    ElseIf (state = 1) Then
        
    End If
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
    
    cResObj.CreateFromFile App.Path & "\" & FileName
    cResObj.RenderBitmap Me.hDC, 0, 0
    
End Sub

'===========================

Public Sub gravity()
    If charCollision = False Then
        'charOnPlatform = False
        If cChar.Y < groundtop - cChar.Height Then
            cChar.Y = cChar.Y + 1
            ScreenNeedsUpdate = True
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

    If (charWalking = 2 Or charWalking = 0) Then
        charWalking = 1
        Call charAnimate
    Else
        charWalking = 1
    End If
    If (charAnimPos > 118 Or charAnimPos < 100) Then
        charAnimPos = 118
    End If
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
    
    If charCollision = False Or (charCollision = True And charOnPlatform = True) Then
        
        'MsgBox ("1")
        If lblTeabag.Visible = True Then
            lblTeabag.Left = lblTeabag.Left + 1
        End If
        If imgAxe0.Visible = True Then
            imgAxe0.Left = cChar.X + 480
        ElseIf imgAxe1.Visible = True Then
            imgAxe1.Left = cChar.X + 480
        End If
        charAnimPos = charAnimPos - 1
        If (charAnimPos = 100) Then
            Call charAnimate
            charAnimPos = 118
        End If
    
    Else
        'Collision move char back
        cChar.X = cChar.X - 1
        charWalking = 3
    End If

End Sub

Public Sub charMoveLeft()

    If (charWalking = 1 Or charWalking = 0) Then
        charWalking = 2
        Call charAnimate
    Else
        charWalking = 2
    End If
    If (charAnimPos > 218 Or charAnimPos < 200) Then
        charAnimPos = 218
    End If
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
        charAnimPos = charAnimPos - 1
        If (charAnimPos = 200) Then
            Call charAnimate
            charAnimPos = 218
        End If
        
    Else
        'Collision move char back
        cChar.X = cChar.X + 1
        charWalking = 3
    End If

End Sub

'Public Function CharAttack(AttackType As String)
'    If (AttackType = "melee") Then
'        Call charAnimate
'    ElseIf (AttackType = "throw") Then
'        'Uncoded
'    End If
'End Function

Private Sub tmrCharAttack_Timer()

    If (cEnemy1.X < cChar.X + cChar.Width And cEnemy1.X + cEnemy1.Width > cChar.X) Then
        If (charAttacking = True) Then
            enemy1Health = enemy1Health - 5
            'charHealth = charHealth - 1
            Call showUserText("You did 5 damage to Tree.", 2000)
            HasAttackedThisKeypress = True
        End If
    End If

End Sub

Public Sub charAnimate()
    'Standing Still
    If (charWalking = 0) Then
        If (charAttacking = True And cChar.Cell > 0 And cChar.Cell <= 6) Then
            cChar.Cell = 7
        ElseIf (charAttacking = False And cChar.Cell = 7) Then
            cChar.Cell = 1
        ElseIf (charAttacking = True And cChar.Cell >= 8 And cChar.Cell <= 13) Then
            cChar.Cell = 14
        ElseIf (charAttacking = False And cChar.Cell = 14) Then
            cChar.Cell = 8
        End If
    'Walking Right
    ElseIf (charWalking = 1) Then
        If (charAttacking = True) Then
            cChar.Cell = 7
        ElseIf (charAttacking = False And cChar.Cell = 7) Then
            cChar.Cell = 1
        ElseIf (cChar.Cell < 6) Then
            cChar.Cell = cChar.Cell + 1
        Else
            cChar.Cell = 1
        End If
    'Walking Left
    ElseIf (charWalking = 2) Then
        If (charAttacking = True) Then
            cChar.Cell = 14
        ElseIf (charAttacking = False And cChar.Cell = 14) Then
            cChar.Cell = 8
        ElseIf (cChar.Cell < 13 And cChar.Cell > 7) Then
            cChar.Cell = cChar.Cell + 1
        Else
            cChar.Cell = 8
        End If
    ElseIf (charWalking = 3) Then
        If (charAttacking = True And cChar.Cell <= 6) Then
            cChar.Cell = 7
        ElseIf (charAttacking = False And cChar.Cell = 7) Then
            cChar.Cell = 1
        ElseIf (charAttacking = True And cChar.Cell >= 8 And cChar.Cell <= 13) Then
            cChar.Cell = 14
        ElseIf (charAttacking = False And cChar.Cell = 14) Then
            cChar.Cell = 8
        ElseIf (cChar.Cell < 6) Then
            cChar.Cell = cChar.Cell + 1
        ElseIf (cChar.Cell = 6) Then
            cChar.Cell = 1
        ElseIf (cChar.Cell < 13 And cChar.Cell > 7) Then
            cChar.Cell = cChar.Cell + 1
        ElseIf (cChar.Cell = 13) Then
            cChar.Cell = 8
        End If
    End If
End Sub

Public Function charCollision() As Boolean

    'MsgBox ("False")
    charOnPlatform = False
    Dim i As Integer
    If cChar.X + cChar.Width > 650 Or cChar.X < 0 Then
        If cChar.X + cChar.Width > 650 And cChar.Y + cChar.Height > 300 Then
            'Call showUserText("Further levels currently under construction.", 2000)
            'NEWFORM
            If (EdnaLevel = False) Then
                previousLevel = 2.4
                Unload Me
                'NEWFORM
                frmLevel2e.Show
            End If
            charCollision = True
        ElseIf cChar.X + cChar.Width > 800 And cChar.Y + cChar.Height < 300 Then
            'NEWFORM
            If (EdnaLevel = True) Then
                previousLevel = 2.4
                Unload Me
                'NEWFORM
                frmLevel2s.Show
            End If
            charCollision = True
        ElseIf cChar.X < 0 Then
            'NEWFORM
            previousLevel = 2.4
            Unload Me
            'NEWFORM
            frmLevel2c.Show
        End If
    End If
    Do While i <= Collision.UBound
        'Left-Right Collision
        If (cChar.X * Screen.TwipsPerPixelX) < Collision(i).Left + Collision(i).Width And (cChar.X * Screen.TwipsPerPixelX) + (cChar.Width * Screen.TwipsPerPixelX) > Collision(i).Left Then
            'Top-Bottom Collision
            'DEBUG THIS
            'If (cChar.Y * Screen.TwipsPerPixelY) < Collision(i).Top + Collision(i).Height And (cChar.Y * Screen.TwipsPerPixelY) + (cChar.Height * Screen.TwipsPerPixelY) > Collision(i).Top Then
                If (cChar.Y * Screen.TwipsPerPixelY) + (cChar.Height * Screen.TwipsPerPixelY) = Collision(i).Top + 15 Then
                    charOnPlatform = True
                    charCollision = True
                    'MsgBox ("Please Work..")
                Else

                End If
                'charCollision = True
            'End If
        End If
        i = i + 1
    Loop
        
End Function

Public Function charObjCollision() As Integer
    
    charObjCollision = -1
    Dim i As Integer
    Do While i <= ItemObj.UBound
        'Left-Right Collision
        If (cChar.X * Screen.TwipsPerPixelX) < ItemObj(i).Left + ItemObj(i).Width And (cChar.X * Screen.TwipsPerPixelX) + (cChar.Width * Screen.TwipsPerPixelX) > ItemObj(i).Left Then
            'Top-Bottom Collision
            If (cChar.Y * Screen.TwipsPerPixelY) < ItemObj(i).Top + ItemObj(i).Height And (cChar.Y * Screen.TwipsPerPixelY) + (cChar.Height * Screen.TwipsPerPixelY) > ItemObj(i).Top Then
                'If (cChar.Y * Screen.TwipsPerPixelY) + (cChar.Height * Screen.TwipsPerPixelY) = Object(i).Top + 15 Then
                    charObjCollision = i
                'End If
            End If
        End If
        i = i + 1
    Loop
 
End Function

Public Function enemyMove(direction As String) As Integer

    If (direction = "right") Then
    
        'Change frame
        If (enemy1Walking = 2 Or enemy1Walking = 0) Then
            enemy1Walking = 1
            Call enemy1Animate
        Else
            enemy1Walking = 1
        End If
        If (enemy1AnimPos > 118 Or enemy1AnimPos < 100) Then
            enemy1AnimPos = 118
        End If
        
        'Collision
        If (cEnemy1.X + cEnemy1.Width < 800) Then
            cEnemy1.X = cEnemy1.X + 1
            enemy1Walking = 1
            
            enemy1AnimPos = enemy1AnimPos - 1
            If (enemy1AnimPos = 200) Then
                Call enemy1Animate
                enemy1AnimPos = 218
            End If
        Else
            enemyMove = 1
            enemy1Walking = 0
        End If
    ElseIf (direction = "left") Then
    
        'Change frame
        If (enemy1Walking = 1 Or enemy1Walking = 0) Then
            enemy1Walking = 2
            Call enemy1Animate
        Else
            enemy1Walking = 2
        End If
        
        'Collision
        If (enemy1AnimPos > 218 Or enemy1AnimPos < 200) Then
            enemy1AnimPos = 218
        End If
        
        If (cEnemy1.X > 0) Then
            cEnemy1.X = cEnemy1.X - 1
            enemy1Walking = 2
            
            enemy1AnimPos = enemy1AnimPos - 1
            If (enemy1AnimPos = 200) Then
                Call enemy1Animate
                enemy1AnimPos = 218
            End If
        Else
            enemyMove = 1
            enemy1Walking = 0
        End If
    End If
    enemyMove = 0

End Function

Public Function enemy1Animate()
    
    'Walking Right
    If (enemy1Walking = 1) Then
        cEnemy1.Cell = 2
    'Walking Left
    ElseIf (enemy1Walking = 2) Then
        cEnemy1.Cell = 1
    'Not Walking
    ElseIf (enemy1Walking = 3) Then
        'Code Here
    End If
    
End Function

Public Function enemyCollision() As Boolean
    
End Function

Public Function enemyObjCollision() As Integer

End Function

Public Function UseItem(ObjCollidedWith As Integer)
    Dim itemname As String
    'NEWFORM
    itemname = Level2x4_Items(ObjCollidedWith).ItemType
    'MsgBox ("itemname: " & itemname)
    Select Case itemname
        Case "chest"
            'MsgBox ("You Found a chest")
            'DEBUG
            'MsgBox ("Chest Collision")
            'NEWFORM
            If (Level2x4_Items(ObjCollidedWith).Used = True) Then
                'MsgBox "Chest Already Opened"
                'Call showUserText("You have already looted this chest.  Nothing remains.", 3000)
                If (tmrUserText.Enabled = False) Then
                    'Call showUserText("You peek in and see some of your past life's charred remains.", 4000)
                    Call showUserText("You have already looted this chest.  Nothing remains.", 3000)
                End If
                'Call showUserText("Greedy Bastard.", 2000)
            'NEWFORM
            ElseIf (Level2x4_Items(ObjCollidedWith).Used = False) Then
                cItemObj(ObjCollidedWith).Cell = 2
                'NEWFORM
                Call levelLoad("level2d.bmp")
                'EXPLODE
                'tmrExplode.Enabled = True
                'Call showUserText("The chest was empty.", 2000)
                'Call showUserText("HAPPY BIRTHDAY - Blow out your pipe bombs... Too late.", 3500)
                'tmrExplode.Enabled = True
                noodle = Int((100 + 1) * Rnd + 1)
                points = points + noodle
                'Call showUserText("You found " & noodle & " gold.", 2000)
                Call showUserText(noodle & " points.", 2000)
                'Call showUserText("The chest was empty.", 2000)
                'NEWFORM
                Level2x4_Items(ObjCollidedWith).Used = True
            End If
        Case "healthbox"
            'healthbox used
            'MsgBox ("You found a health box")
            'NEWFORM
            If (Level2x4_Items(ObjCollidedWith).Used = True) Then
                Call showUserText("This healthbox has already been utilized.", 2000)
                'If (Cheat_health = True) Then
                    'Level1x1_Items(ObjCollidedWith).Used = False
                    'cItemObj(ObjCollidedWith).Cell = 0
                'End If
            'NEWFORM
            ElseIf (Level2x4_Items(ObjCollidedWith).Used = False) Then
                cItemObj(ObjCollidedWith).Cell = 2
                'NEWFORM
                Call levelLoad("level2d.bmp")
                peanut = Int((100 + 1) * Rnd + 1)
                Dim actualrestore As Integer
                If (charHealth >= 100) Then
                    actualrestore = 0
                Else
                    If (peanut >= (100 - charHealth)) Then
                        actualrestore = 100 - charHealth
                    Else
                        actualrestore = peanut
                    End If
                End If
                charHealth = charHealth + actualrestore
                Call showUserText("Restoring " & actualrestore & " life.", 2000)
                'NEWFORM
                Level2x4_Items(ObjCollidedWith).Used = True
            End If
    End Select
End Function

Public Sub tmrEnemyGenerator_Timer()
    
    cEnemy1.X = Int(((800 - cEnemy1.Width) + 1) * Rnd + 1)
    cEnemy1.Y = (groundtop - cEnemy1.Height)
    cEnemy1.Cell = 1
    enemy1Health = 100
    enemyExists = True
    tmrEnemyGenerator.Enabled = False
    ScreenNeedsUpdate = True
    
End Sub

Private Sub tmrEnemyMove_Timer()
    
    If (cEnemy1.X > cChar.X + cChar.Width - 1) Then
        If (enemyMove("left") = 0) Then
            ScreenNeedsUpdate = True
        End If
    ElseIf (cEnemy1.X + cEnemy1.Width < cChar.X + 1) Then
        If (enemyMove("right") = 0) Then
            ScreenNeedsUpdate = True
        End If
    End If
            
End Sub

Private Sub tmrEnemyAttack_Timer()

    If (cEnemy1.X < cChar.X + cChar.Width And cEnemy1.X + cEnemy1.Width > cChar.X And cEnemy1.Y < cChar.Y + cChar.Height And cEnemy1.Y + cEnemy1.Height > cChar.Y) Then
        charHealth = charHealth - 2
        Call showUserText("You took 2 damage from Tree.", 2000)
    End If

End Sub

Public Sub tmrGame_Timer()
    
    'lblTime.Caption = timeleft
    timeleft = timeleft - 1
    Me.Caption = "Jim the Axe Man(VB is gay) - TIME: " & timeleft
    'Call showUserText("Enemy Life: " & timeleft, 300)
    
End Sub

Private Sub tmrWalk_Timer()

    'Walking Right
    If (charWalking = 1) Then
        If (cChar.Cell < 6) Then
            cChar.Cell = cChar.Cell + 1
        Else
            cChar.Cell = 1
        End If
    'Walking Left
    ElseIf (charWalking = 2) Then
        If (cChar.Cell < 13 And cChar.Cell > 7) Then
            cChar.Cell = cChar.Cell + 1
        Else
            cChar.Cell = 8
        End If
    ElseIf (charWalking = 3) Then
        If (cChar.Cell < 6) Then
            cChar.Cell = cChar.Cell + 1
        ElseIf (cChar.Cell = 6) Then
            cChar.Cell = 1
        ElseIf (cChar.Cell < 13 And cChar.Cell > 7) Then
            cChar.Cell = cChar.Cell + 1
        ElseIf (cChar.Cell = 13) Then
            cChar.Cell = 8
        End If
    End If

End Sub

Private Sub tmrExplode_Timer()

    If (cChar.Cell < 15) Then
        cChar.Cell = 15
        YouveBeenHit = 1
    ElseIf (cChar.Cell < 21) Then
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
    ItemBeingUsed = False
    
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
    ElseIf txtDebug.text = "health" Then
        charHealth = 100
    ElseIf txtDebug.text = "nme" Then
        tmrEnemyGenerator.Enabled = False
    Else
        MsgBox ("0x0000")
    End If
    txtDebug.text = ""
    txtDebug.Visible = False
    cmdDebug.Visible = False

End Function

'=================

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If (m_bInGame) Then
        GameLoop
    End If
    tmrGame.Enabled = False
    Unload Me ' frmLevel1a
End Sub
