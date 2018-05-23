VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cTile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ======================================================================
' Class:    cTile
' Filename: cTile.cls
' Author:   SP McMahon
' Date:     14 July 1998
'
' Tiles a sprite bitmap into a specified area
' ======================================================================

' Filename of bitmap to use in tiling:
Private m_sFileName As String
' m_cSprite & m_cSpriteBitmap Objects:
Private m_cBitmap As cBitmap
' Object to render on:
Private m_objRender As Object

Public Sub TileObject()
Dim lOH As Long
Dim lOW As Long
Dim lHDC As Long

    With m_objRender
        lOW = .ScaleWidth \ Screen.TwipsPerPixelX
        lOH = .ScaleHeight \ Screen.TwipsPerPixelY
        lHDC = .hDC
    End With
    TileDC lHDC, lOW, lOH
            
End Sub
Public Sub TileDC( _
        ByVal lHDC As Long, _
        ByVal lWidth As Long, _
        ByVal lHeight As Long _
    )
Dim lX As Long
Dim lY As Long
Dim dX As Long
Dim dY As Long

    dX = m_cBitmap.Width
    dY = m_cBitmap.Height
    lY = 0
    Do While lY <= lHeight
        lX = 0
        Do While lX <= lWidth
            m_cBitmap.RenderBitmap lHDC, lX, lY
            lX = lX + dX
        Loop
        lY = lY + dY
    Loop
    
End Sub
Property Let FileName( _
        ByVal sFileName As String _
    )
Dim hDCRender As Long
    If (sFileName <> m_sFileName) Then
        If (sFileName = "") Then
            ClearUp
        Else
            If (pbValid()) Then
                ' We create the sprite bitmap from the file if
                ' possible:
                If (m_cBitmap.CreateFromFile( _
                        sFileName)) Then
                    m_sFileName = sFileName
                End If
            End If
        End If
    End If
End Property
Property Let Picture( _
        ByRef oPic As StdPicture _
    )
    ClearUp
    ' We create the sprite bitmap from the file if
    ' possible:
    Set m_cBitmap = New cBitmap
    If (m_cBitmap.CreateFromPicture( _
            oPic)) Then
    End If
End Property
Private Function pbValid() As Boolean
    If (Not (m_objRender Is Nothing)) Then
        pbValid = True
    Else
        pErr 1, "Object not initialised"
    End If
End Function
Public Sub Initialise( _
        ByVal objRender As Object _
    )
    Set m_objRender = objRender
End Sub
Private Sub pErr( _
        ByVal lErrNum As Long, _
        ByVal sErrText As String _
    )
    ' err.raise vbObjectError + 1048 + lErrNum,App.ExeName & ".m_cTileile",sErrText
    Debug.Print "Error: " & sErrText
End Sub
Private Sub ClearUp()
    Set m_cBitmap = Nothing
    Set m_objRender = Nothing
    m_sFileName = ""
End Sub
Private Sub Class_Initialize()
    ' Create valid instances of the sprite and
    ' sprite bitmap Objects:
    Debug.Print "m_cTileile:Initialise"
    Set m_cBitmap = New cBitmap
End Sub

Private Sub Class_Terminate()
    ' clear up all the Object references:
    Debug.Print "m_cTileile:Terminate"
    ClearUp
End Sub
