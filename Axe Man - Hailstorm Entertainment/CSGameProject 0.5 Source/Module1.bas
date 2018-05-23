Attribute VB_Name = "Globals"
Public ArrowPos As Integer
Public EdnaLevel As Boolean

Public timeleft As Integer
Public charWalking As Integer
Public previousLevel As Double
Public currentLevel As Double
Public charHealth As Integer
Public enemy1Health As Integer
Public points As Integer

Type item
  FileName As String
  ItemType As String
  Used As Boolean
End Type

Public Level1x1_ItemDelcaration As Boolean
Public Level1x1_IntroText As Boolean 'True = Already Printed
Public Level1x1_Items() As item 'True = Opened
'Public Level1x2_IntroText As Boolean 'True = Already Printed
Public Level1x2_Items() As item 'True = Opened
'Public Level1x3_IntroText As Boolean 'True = Already Printed
Public Level1x3_Items() As item 'True = Opened
'Public Level1x4_IntroText As Boolean 'True = Already Printed
Public Level1x4_Items() As item 'True = Opened
'Public Level1x5_IntroText As Boolean 'True = Already Printed
Public Level1x5_Items() As item 'True = Opened
'Public Level1x6_IntroText As Boolean 'True = Already Printed
Public Level1x6_Items() As item 'True = Opened

Public Level2x1_ItemDelcaration As Boolean
Public Level2x1_IntroText As Boolean 'True = Already Printed
Public Level2x1_Items() As item 'True = Opened
'Public Level2x2_IntroText As Boolean 'True = Already Printed
Public Level2x2_Items() As item 'True = Opened
'Public Level2x3_IntroText As Boolean 'True = Already Printed
Public Level2x3_Items() As item 'True = Opened
'Public Level2x4_IntroText As Boolean 'True = Already Printed
Public Level2x4_Items() As item 'True = Opened
'Public Level2x5_IntroText As Boolean 'True = Already Printed
Public Level2x5_Items() As item 'True = Opened
Public Level2xS_IntroText As Boolean 'True = Already Printed
Public Level2xS_Items() As item 'True = Opened

Public Level3x1_ItemDelcaration As Boolean
Public Level3x1_IntroText As Boolean 'True = Already Printed
Public Level3x1_Items() As item 'True = Opened
Public Level3x2_Items() As item 'True = Opened
Public Level3x3_Items() As item 'True = Opened
Public Level3x4_Items() As item 'True = Opened
Public Level3x5_Items() As item 'True = Opened
Public Level3x6_Items() As item 'True = Opened
Public Level3x7_Items() As item 'True = Opened
Public Level3x8_Items() As item 'True = Opened

Public Level4x1_ItemDelcaration As Boolean
Public Level4x1_IntroText As Boolean 'True = Already Printed
Public Level4x1_Items() As item 'True = Opened
Public Level4x2_Items() As item 'True = Opened
Public Level4x3_Items() As item 'True = Opened
Public Level4x4_Items() As item 'True = Opened
Public Level4x5_Items() As item 'True = Opened
Public Level4x6_Items() As item 'True = Opened

Public cResChar As cSpriteBitmaps
Public cChar As cSprite
Public cResEnemy1 As cSpriteBitmaps
Public cEnemy1 As cSprite
'Public cResChestObj() As cSpriteBitmaps
'Public cChestObj() As cSprite
Public cResItemObj() As cSpriteBitmaps
Public cItemObj() As cSprite
'Public cResHealthObj() As cSpriteBitmaps
'Public cHealthObj() As cSprite

'Excluded Private cBack As cBitmap

' Get Key presses:
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

' General for game:
Public m_bInGame As Boolean

Public Sub UnloadAllForms()

    Dim oFrm As Form
    
    For Each oFrm In Forms
        Unload oFrm
    Next

End Sub
