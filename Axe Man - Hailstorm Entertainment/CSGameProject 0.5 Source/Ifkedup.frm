VERSION 5.00
Begin VB.Form Ifkedup 
   Caption         =   "IFkedUpBad"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Ifkedup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case (KeyCode)
        Case vbKeyRight
            Call CharMove(0, 1)
        Case vbKeyLeft
            Call CharMove(1, 1)
        Case vbKeyUp
            Call CharMove(2, 1)
        Case vbKeyDown
            Call CharMove(3, 1)
        Case vbKeyZ
            Call CharAttack(0, 1)
        Case vbKeyX
            Call CharAttack(1, 1)
        'Case vbKeyReturn
        '    Call DoDebug
        Case vbKeyF11
            txtDebug.Visible = True
            cmdDebug.Visible = True
        Case vbKeyEscape
            Unload Me
    End Select
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case (KeyCode)
        Case vbKeyRight
            Call CharMove(0, 0)
        Case vbKeyLeft
            Call CharMove(1, 0)
        Case vbKeyUp
            'Call CharMove(2, 0)
        Case vbKeyDown
            Call CharMove(3, 0)
        Case vbKeyZ
            Call CharAttack(0, 0)
        Case vbKeyX
            Call CharAttack(1, 0)
    End Select
End Sub

Public Function CharMove(Direction As Integer, KeyPos As Integer)
    
    Select Case (Direction)
        Case 0
            If KeyPos = 1 And cChar.X <= 11400 Then 'And (tmrMoveLeft.Enabled = False) Then
                'move right
                tmrMoveLeft.Enabled = False
                tmrMoveRight.Enabled = True
            ElseIf KeyPos = 0 And (tmrMoveRight.Enabled = True) Then
                tmrMoveRight.Enabled = False
            End If
            'imgCharacter.Left = imgCharacter.Left + 50
        Case 1
            If KeyPos = 1 And cChar.X >= 0 Then 'And (tmrMoveRight.Enabled = False) Then
                'move left
                tmrMoveRight.Enabled = False
                tmrMoveLeft.Enabled = True
            ElseIf KeyPos = 0 And (tmrMoveLeft.Enabled = True) Then
                tmrMoveLeft.Enabled = False
            End If
            'imgCharacter.Left = imgCharacter.Left - 50
        Case 2
            If jumpnumber = 0 And KeyPos = 1 Then
                'jump
                jumpdir = 1
                jumpnumber = 1
                tmrJump.Enabled = True
            End If
            'imgCharacter.Top = imgCharacter.Top - 15
        Case 3
            'If KeyPos = 1 Then
            '    'crouch
            '    If lblTeabag.Visible = False And Cheat_teabag = 1 Then
            '        lblTeabag.Left = cChar.X - 550
            '        lblTeabag.Top = cChar.Y - (lblTeabag.Height + 50)
            '        lblTeabag.Visible = True
            '    End If
            '    imgCharacter.Picture = LoadPicture("charactercrouch.bmp")
            'ElseIf KeyPos = 0 Then
            '    'un-crouch
            '    If Cheat_teabag = 1 Then
            '        lblTeabag.Visible = False
            '    End If
            '    imgCharacter.Picture = LoadPicture("character.bmp")
            'End If
    End Select
    
End Function

Private Sub tmrMoveLeft_Timer()
    cChar.StoreBackground cStage.hDC, cChar.X, cChar.Y
    
    If cChar.X >= 0 Then
        cChar.X = cChar.X - 1
    Else
        tmrMoveLeft.Enabled = False
    End If
    If lblTeabag.Visible = True Then
        lblTeabag.Left = lblTeabag.Left - 1
    End If
    If imgAxe0.Visible = True Then
        imgAxe0.Left = cChar.X + 480
    ElseIf imgAxe1.Visible = True Then
        imgAxe1.Left = cChar.X + 480
    End If
    
    cChar.TransparentDraw cStage.hDC, cChar.X, cChar.Y, cChar.Cell, False
    cChar.StageToScreen lHDC, cStage.hDC
    DoEvents
End Sub

Private Sub tmrMoveRight_Timer()
    cChar.StoreBackground cStage.hDC, cChar.X, cChar.Y
    
    If cChar.X <= 11400 Then
        cChar.X = cChar.X + 1
    Else
        tmrMoveRight.Enabled = False
    End If
    If lblTeabag.Visible = True Then
        lblTeabag.Left = lblTeabag.Left + 1
    End If
    If imgAxe0.Visible = True Then
        imgAxe0.Left = cChar.X + 480
     ElseIf imgAxe1.Visible = True Then
        imgAxe1.Left = cChar.X + 480
    End If
    
    cChar.TransparentDraw cStage.hDC, cChar.X, cChar.Y, cChar.Cell, False
    cChar.StageToScreen lHDC, cStage.hDC
    DoEvents
End Sub

Private Sub tmrJump_Timer()
    cChar.StoreBackground cStage.hDC, cChar.X, cChar.Y
    
    If jumpdir = 1 Then
        If cChar.Y > jumpheight Then
            cChar.Y = cChar.Y - 1
        ElseIf cChar.Y <= jumpheight Then
            jumpdir = 2
            tmrJump.Interval = 10
        End If
    ElseIf jumpdir = 2 Then
        If cChar.Y >= (groundtop - cChar.Height) Then
            cChar.Y = (groundtop - cChar.Height)
            tmrJump.Enabled = False
            tmrJump.Interval = 10
            jumpdir = 1
            jumpnumber = 0
        ElseIf cChar.Y < (groundtop - cChar.Height) Then
            cChar.Y = cChar.Y + 50
        End If
    End If
    'If imgAxe0.Visible = True Then
    '    imgAxe0.Top = imgCharacter.Top
    'ElseIf imgAxe1.Visible = True Then
    '    imgAxe1.Top = imgCharacter.Top
    'End If
    
    cChar.TransparentDraw cStage.hDC, cChar.X, cChar.Y, cChar.Cell, False
    cChar.StageToScreen lHDC, cStage.hDC
    DoEvents
End Sub

Private Sub cmdDebug_Click()
    Call DoCheat
End Sub

Public Function DoCheat()

    If txtDebug.text = "superjump" Then
        If Cheat_superjump = 0 Then
            Cheat_superjump = 1
            jumpheight = 1000
            MsgBox ("0x1001")
        ElseIf Cheat_superjump = 1 Then
            Cheat_superjump = 0
            jumpheight = 5160
            MsgBox ("0x0001")
        End If
    ElseIf txtDebug.text = "teabag" Then
        If Cheat_teabag = 0 Then
            Cheat_teabag = 1
            MsgBox ("0x1002")
        ElseIf Cheat_teabag = 1 Then
            Cheat_teabag = 0
            MsgBox ("0x0002")
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

'=================================================================================
'                                   BACKUP CODE
'=================================================================================

Public Function charCollisionBak() As Boolean
    Dim i As Integer
    If cChar.X <= 760 And cChar.X >= 0 Then
        Do While i <= Collision.UBound
            'MsgBox (i)
            'Left-Right Collision
            If (cChar.X * Screen.TwipsPerPixelX) > Collision(i).Left And (cChar.X * Screen.TwipsPerPixelX) < Collision(i).Left + Collision(i).Width Then
                'Top-Bottom Collision
                If (cChar.Y * Screen.TwipsPerPixelY) > Collision(i).Top And (cChar.Y * Screen.TwipsPerPixelY) < Collision(i).Top + Collision(i).Height Then
                    'MsgBox ("Collide")
                    charCollisionBak = True
                    
                Else
                    'MsgBox ("cCharx" & cChar.X * Screen.TwipsPerPixelX)
                    'MsgBox ("cCharY" & cChar.Y * Screen.TwipsPerPixelY)
                    'MsgBox ("cCharH" & cChar.Height * Screen.TwipsPerPixelY)
                    'MsgBox ("cCharW" & cChar.Width * Screen.TwipsPerPixelX)
                    'MsgBox ("ileft" & Collision(i).Left)
                    'MsgBox ("itop" & Collision(i).Top)
                    'MsgBox ("iT+h" & (Collision(i).Top + Collision(i).Height))
                    'MsgBox ("iL+w" & (Collision(i).Left + Collision(i).Width))
                End If
                'MsgBox ("Collide3")
            End If
            i = i + 1
            
        Loop
    Else
        charCollisionBak = True
    End If
    'MsgBox (i)
End Function

'=================================================================================

        If (GetAsyncKeyState(vbKeyUp) <> 0) And (jumpnumber = 0) Then
            'jump
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
            If cChar.X <= 760 Then
                If charCollision = False Then
                
                    If cChar.X <= 760 Then
                        cChar.X = cChar.X + 1
                    'Else
                        'tmrMoveRight.Enabled = False
                    End If
                    If lblTeabag.Visible = True Then
                        lblTeabag.Left = lblTeabag.Left + 1
                    End If
                    If imgAxe0.Visible = True Then
                        imgAxe0.Left = cChar.X + 480
                    ElseIf imgAxe1.Visible = True Then
                        imgAxe1.Left = cChar.X + 480
                    End If
                    
                Else
                    'MsgBox ("Collide Right")
                End If
            End If
        ElseIf (GetAsyncKeyState(vbKeyLeft) <> 0) Then
            'move left
            If cChar.X >= 0 Then
                If charCollision = False Then
                
                    cChar.X = cChar.X - 1
                    If lblTeabag.Visible = True Then
                        lblTeabag.Left = lblTeabag.Left - 1
                    End If
                    If imgAxe0.Visible = True Then
                        imgAxe0.Left = cChar.X + 480
                    ElseIf imgAxe1.Visible = True Then
                        imgAxe1.Left = cChar.X + 480
                    End If
                        
                Else
                    'MsgBox ("Collide Left")
                End If
            End If
        ElseIf (GetAsyncKeyState(vbKeyDown) <> 0) Then
            'crouch
            '--cChar.Y = cChar.Y + 1
            If charCollision = False Then
            
                cChar.Y = cChar.Y + 1
                If lblTeabag.Visible = True Then
                    lblTeabag.Top = lblTeabag.Top + 1
                End If
                If imgAxe0.Visible = True Then
                    imgAxe0.Top = cChar.Y + 480
                ElseIf imgAxe1.Visible = True Then
                    imgAxe1.Top = cChar.Y + 480
                End If
                    
            Else
                'MsgBox ("Collide Down")
            End If
        ElseIf (GetAsyncKeyState(vbKeyF11) <> 0) Then
            txtDebug.Visible = True
            cmdDebug.Visible = True
        ElseIf (GetAsyncKeyState(vbKeyEscape) <> 0) Then
        '    ' accelerate in current direction:
            Unload Me
        End If
        
        If charJumping = True Then
            Call charJump
        End If
        
        If gravityCheck = True Then
            Call gravity
        End If

'=================================================================================

Public Sub gravityBak()
    If cChar.X < platform1right Then
        If cChar.Y < platform1top - cChar.Height Then
            cChar.Y = cChar.Y + 1
        ElseIf cChar.Y > platform1top - cChar.Height And cChar.Y < (groundtop - cChar.Height) Then
            cChar.Y = cChar.Y + 1
        Else
            jumpnumber = 0
        End If
    ElseIf cChar.X >= platform1right Then
        If cChar.Y < groundtop - cChar.Height Then
            cChar.Y = cChar.Y + 1
        ElseIf cChar.Y > groundtop - cChar.Height Then
            cChar.Y = groundtop - cChar.Height
        Else
            jumpnumber = 0
        End If
    End If
End Sub

