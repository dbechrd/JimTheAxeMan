VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cBitmap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ======================================================================
' Class:    cBitmap
' Filename: cBitmap.cls
' Author:   SP McMahon
' Date:     14 July 1998
'
' A class to store a bitmap.  Similar to cSpriteBitmap but doesn't
' use a mask.
' ======================================================================

' Memory DC & Bitmap to hold the bitmap
Private m_hdcBitmap As Long
Private m_hBmpBitmap As Long
Private m_hBmpBitmapOld As Long
' Width & Height of the bitmap
Private m_lDX As Long
Private m_lDY As Long

Public Function CreateAtSize( _
        ByVal lDX As Long, _
        ByVal lDY As Long _
    ) As Boolean

    ' First ensure that we've cleared up anything
    ' we already have:
    ClearUp
    
    ' Create a DC at size to hold the bitmap:
    If (GDIMakeDCAndBitmap( _
        False, _
        m_hdcBitmap, _
        m_hBmpBitmap, _
        m_hBmpBitmapOld, _
        lDX, lDY)) Then
        m_lDX = lDX
        m_lDY = lDY
        CreateAtSize = True
    End If
    
End Function
Public Function CreateFromFile( _
        ByVal sBitmapBitmap As String _
    ) As Boolean
Dim tBM As Bitmap

    ' First ensure that we've cleared up anything
    ' we already have:
    ClearUp
    
    ' Load the Bitmap bitmap:
    If (GDILoadBitmapIntoDC( _
        False, _
        sBitmapBitmap, _
        m_hdcBitmap, _
        m_hBmpBitmap, _
        m_hBmpBitmapOld)) Then
        ' Get the dimensions of the bitmap (which is
        ' a DIB):
        GetObjectAPI m_hBmpBitmap, Len(tBM), tBM
        ' Store dimensions:
        m_lDX = tBM.bmWidth
        m_lDY = tBM.bmHeight
        CreateFromFile = True
    End If
    
End Function
Public Function CreateFromPicture( _
        ByRef picThis As StdPicture _
    )
Dim tBM As Bitmap
    ClearUp
    If (GDILoadPictureIntoDC(False, picThis, m_hdcBitmap, m_hBmpBitmap, m_hBmpBitmapOld)) Then
        ' Get the dimensions of the bitmap (which is
        ' a DIB):
        GetObjectAPI m_hBmpBitmap, Len(tBM), tBM
        ' Store dimensions:
        m_lDX = tBM.bmWidth
        m_lDY = tBM.bmHeight
        CreateFromPicture = True
    End If
End Function
Property Get Width() As Long
    Width = m_lDX
End Property
Property Get Height() As Long
    Height = m_lDY
End Property
Property Get hDC() As Long
    hDC = m_hdcBitmap
End Property
Public Sub RenderBitmap( _
        ByVal lHDC As Long, _
        ByVal lX As Long, _
        ByVal lY As Long, _
        Optional ByVal OpCode As RasterOpConstants = vbSrcCopy _
    )
    BitBlt _
        lHDC, _
        lX, lY, m_lDX, m_lDY, _
        m_hdcBitmap, _
        0, 0, _
        OpCode
End Sub
Private Sub ClearUp()
    ' Here we ensure all our Objects are destroyed:
    GDIClearDCBitmap m_hdcBitmap, m_hBmpBitmap, m_hBmpBitmapOld
End Sub
Private Sub Class_Terminate()
    ClearUp
End Sub
