VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cSpriteBitmaps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ======================================================================
' Class:    cSpriteBitmaps
' Filename: cSpriteBitmaps.cls
' Author:   SP McMahon
' Date:     14 July 1998
'
' A class to store a picture clip of sprite bitmaps and an associated
' mask to allow transparent drawing onto a DC.
' ======================================================================

' Memory DC & Bitmap to hold the sprite
Private m_hdcSprite As Long
Private m_hBmpSprite As Long
Private m_hBmpSpriteOld As Long

' Memory DC & Bitmap to hold the mask element of the sprite
Private m_hdcMask As Long
Private m_hBmpMask As Long
Private m_hBmpMaskOld As Long

' Memory DC to build up the sprite and draw it:
Private m_hDCBuild As Long
Private m_hBmpBuild As Long
Private m_hBmpBuildOld As Long

' Variables to store sprite frames, current cell and cell size:
Private m_lDX As Long
Private m_lDY As Long
Private m_lCX As Long
Private m_lCY As Long

Public Sub TransparentDraw( _
        ByVal hDCDest As Long, _
        ByVal lX As Long, _
        ByVal lY As Long, _
        ByVal lCell As Long, _
        Optional ByVal bUseBuildDC As Boolean = True _
    )
Dim lDX As Long
Dim lDY As Long
Dim lDC As Long
Dim loX As Long
Dim loY As Long

    ' Get the position of the sprite cell within the data:
    pGetXYForCell lCell, lDX, lDY
    
    ' Determine where to draw to:
    If (bUseBuildDC) Then
        ' Now we take a copy of the background we are drawing on
        ' into the work DC:
        lDC = m_hDCBuild
        BitBlt lDC, 0, 0, m_lDX, m_lDY, hDCDest, lX, lY, SRCCOPY
        loX = 0
        loY = 0
    Else
        lDC = hDCDest
        loX = lX
        loY = lY
    End If
    
    ' We OR the mask into place, i.e. wherever the mask is black
    ' there is no effem_cTile, but wherever it is white the background
    ' goes completely white:
    BitBlt lDC, loX, loY, m_lDX, m_lDY, m_hdcMask, lDX, lDY, SRCPAINT
    ' Now we AND the sprite into position: wherever the sprite is white,
    ' there is no effem_cTile, but over the whitened mask area the sprite
    ' shows as normal:
    BitBlt lDC, loX, loY, m_lDX, m_lDY, m_hdcSprite, lDX, lDY, SRCAND
    
    If (bUseBuildDC) Then
        ' Now we can blit the combination of the background with the
        ' transparently drawn area onto the Destination:
        BitBlt hDCDest, lX, lY, m_lDX, m_lDY, m_hDCBuild, 0, 0, SRCCOPY
    End If
    
End Sub

Public Function CreateFromPicture( _
        ByRef oPic As StdPicture, _
        ByVal lXImageCount As Long, _
        ByVal lYImageCount As Long, _
        Optional ByRef oMaskPic As StdPicture = Nothing, _
        Optional ByVal lTransColor As Long = &HFFFFFF _
    ) As Boolean
Dim tBM As Bitmap
Dim bContinue As Boolean
    
    ' First ensure that we've cleared up anything
    ' we already have:
    ClearUp
    
    ' First create the sprite bitmap:
    If (GDILoadPictureIntoDC( _
            False, _
            oPic, _
            m_hdcSprite, _
            m_hBmpSprite, _
            m_hBmpSpriteOld _
            )) Then
        ' Get the dimensions of the bitmap (which is
        ' a DIB):
        GetObjectAPI m_hBmpSprite, Len(tBM), tBM
        
        If (oMaskPic Is Nothing) Then
            ' Create a mask based on white areas of the
            ' original sprite:
            bContinue = pbCreateMaskFromSprite( _
                tBM.bmWidth, _
                tBM.bmHeight, _
                lTransColor _
                )
        Else
            ' We load the specified mask image:
            bContinue = GDILoadPictureIntoDC( _
                True, _
                oMaskPic, _
                m_hdcMask, _
                m_hBmpMask, _
                m_hBmpMaskOld)
        End If
        
        If (bContinue) Then
            ' Set the x and y count for image cells in the sprite:
            m_lCX = lXImageCount
            m_lCY = lYImageCount
            ' The width & Height of the sprite:
            m_lDX = tBM.bmWidth \ m_lCX
            m_lDY = tBM.bmHeight \ m_lCY
            ' Make a working DC to draw the sprite in:
            If (GDIMakeDCAndBitmap( _
                False, _
                m_hDCBuild, m_hBmpBuild, m_hBmpBuildOld, _
                m_lDX, m_lDY)) Then
                CreateFromPicture = True
            Else
                Debug.Print "Failed to create working DC/Bitmap"
            End If
        Else
            Debug.Print "Failed to create mask DC/Bitmap"
        End If
    Else
        Debug.Print "Failed to create sprite DC/Bitmap"
    End If
    
End Function

Public Function CreateFromFile( _
        ByVal sSpriteBitmap As String, _
        ByVal lXImageCount As Long, _
        ByVal lYImageCount As Long, _
        Optional ByVal sMaskBitmap As String = "", _
        Optional ByVal lTransColor As Long = &HFFFFFF _
    ) As Boolean
Dim tBM As Bitmap
Dim bContinue As Boolean

    ' First ensure that we've cleared up anything
    ' we already have:
    ClearUp
    
    ' First create the sprite bitmap:
    If (GDILoadBitmapIntoDC( _
            False, _
            sSpriteBitmap, _
            m_hdcSprite, _
            m_hBmpSprite, _
            m_hBmpSpriteOld _
            )) Then
        ' Get the dimensions of the bitmap (which is
        ' a DIB):
        GetObjectAPI m_hBmpSprite, Len(tBM), tBM
        
        If (sMaskBitmap = "") Then
            ' Create a mask based on white areas of the
            ' original sprite:
            bContinue = pbCreateMaskFromSprite( _
                tBM.bmWidth, _
                tBM.bmHeight, _
                lTransColor _
                )
        Else
            ' We load the specified mask image:
            bContinue = GDILoadBitmapIntoDC( _
                True, _
                sMaskBitmap, _
                m_hdcMask, _
                m_hBmpMask, _
                m_hBmpMaskOld)
        End If
        
        If (bContinue) Then
            ' Set the x and y count for image cells in the sprite:
            m_lCX = lXImageCount
            m_lCY = lYImageCount
            ' The width & Height of the sprite:
            m_lDX = tBM.bmWidth \ m_lCX
            m_lDY = tBM.bmHeight \ m_lCY
            ' Make a working DC to draw the sprite in:
            If (GDIMakeDCAndBitmap( _
                False, _
                m_hDCBuild, m_hBmpBuild, m_hBmpBuildOld, _
                m_lDX, m_lDY)) Then
                CreateFromFile = True
            Else
                Debug.Print "Failed to create working DC/Bitmap"
            End If
        Else
            Debug.Print "Failed to create mask DC/Bitmap"
        End If
    Else
        Debug.Print "Failed to create sprite DC/Bitmap"
    End If
    
End Function
Private Function pbCreateMaskFromSprite( _
        ByVal bmWidth As Long, _
        ByVal bmHeight As Long, _
        ByVal lTransColor As Long _
    ) As Boolean
Dim lOrigColor As Long
Dim bTransColorIsWhite As Boolean
    
    ' Now create a mask Object.  Here we want
    ' to just copy the bitmap into a monochrome
    ' DC, then invert it:
    If (GDIMakeDCAndBitmap( _
        True, _
        m_hdcMask, m_hBmpMask, m_hBmpMaskOld, _
        bmWidth, bmHeight)) Then
        
        If (lTransColor = &HFFFFFF) Then
            bTransColorIsWhite = True
        End If
        
        If (bTransColorIsWhite) Then
            ' Copy the Bitmap from the source, inverting it
            ' as we go:
            BitBlt m_hdcMask, 0, 0, bmWidth, bmHeight, m_hdcSprite, 0, 0, &HFF0062 ' Whiteness
            BitBlt m_hdcMask, 0, 0, bmWidth, bmHeight, m_hdcSprite, 0, 0, SRCINVERT
            pbCreateMaskFromSprite = True
        Else
            ' We need to make an non-inverse mask first.  This
            ' is ORed with the original sprite to make the area
            ' with transparent colour = white, then it the mask is
            ' inverted.
            Dim hdcTemp As Long, hBmpTemp As Long, hBmpTempOld As Long
            
            If (GDIMakeDCAndBitmap( _
                True, _
                hdcTemp, hBmpTemp, hBmpTempOld, _
                bmWidth, bmHeight)) Then
                
                ' Set the back colour for the sprite to the
                ' transparent colour (this means it is set to
                ' black when we bitblt to the mono DC):
                lOrigColor = SetBkColor(m_hdcSprite, lTransColor)
                ' Make non-inverse mask:
                BitBlt hdcTemp, 0, 0, bmWidth, bmHeight, m_hdcSprite, 0, 0, SRCCOPY
                ' Return back colour of sprite
                SetBkColor m_hdcSprite, lOrigColor
                ' Make sprite white in lTransColor region:
                BitBlt m_hdcSprite, 0, 0, bmWidth, bmHeight, hdcTemp, 0, 0, SRCPAINT
                ' Invert the mask:
                BitBlt m_hdcMask, 0, 0, bmWidth, bmHeight, m_hdcSprite, 0, 0, &HFF0062 ' Whiteness
                BitBlt m_hdcMask, 0, 0, bmWidth, bmHeight, hdcTemp, 0, 0, SRCINVERT
                
                ' Clear up temporary bitmap:
                GDIClearDCBitmap hdcTemp, hBmpTemp, hBmpTempOld
            
                pbCreateMaskFromSprite = True
            End If
        End If
    End If

End Function
Property Get CellWidth() As Long
    CellWidth = m_lDX
End Property
Property Get CellHeight() As Long
    CellHeight = m_lDY
End Property
Public Sub DirectBltSprite( _
        ByVal lHDC As Long, _
        ByVal lX As Long, _
        ByVal lY As Long, _
        ByVal lCell As Long _
    )
Dim lDX As Long
Dim lDY As Long
    pGetXYForCell lCell, lDX, lDY
    BitBlt lHDC, lX, lY, m_lDX, m_lDY, m_hdcSprite, lDX, lDY, SRCCOPY
End Sub
Private Sub pGetXYForCell( _
        ByVal lCell As Long, _
        ByRef lX As Long, _
        ByRef lY As Long _
    )
    lX = ((lCell - 1) Mod m_lCX) * m_lDX
    lY = ((lCell - 1) \ m_lCX) * m_lDY
End Sub

Private Sub ClearUp()
    ' Here we ensure all our Objects are destroyed:
    GDIClearDCBitmap m_hdcSprite, m_hBmpSprite, m_hBmpSpriteOld
    GDIClearDCBitmap m_hdcMask, m_hBmpMask, m_hBmpMaskOld
    GDIClearDCBitmap m_hDCBuild, m_hBmpBuild, m_hBmpBuildOld
End Sub

Private Sub Class_Initialize()
    Debug.Print "m_cSpriteBitmaps:Initialise"
End Sub

Private Sub Class_Terminate()
    Debug.Print "m_cSpriteBitmaps:Terminate"
    ClearUp
End Sub
