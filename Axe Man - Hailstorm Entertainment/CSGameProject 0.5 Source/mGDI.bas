Attribute VB_Name = "mGDI"
Option Explicit

' API Declares:

' This is most useful but Win32 only.  Particularly try the
' LOADMAP3DCOLORS for a quick way to sort out those
' embarassing gray backgrounds in your fixed bitmaps!
Declare Function LoadImage Lib "user32" Alias "LoadImageA" ( _
    ByVal hInst As Long, _
    ByVal lpsz As String, _
    ByVal un1 As Long, _
    ByVal n1 As Long, ByVal n2 As Long, _
    ByVal un2 As Long _
    ) As Long
Public Const IMAGE_BITMAP = 0
Public Const IMAGE_ICON = 1
Public Const IMAGE_CURSOR = 2
Public Const LR_COLOR = &H2
Public Const LR_COPYDELETEORG = &H8
Public Const LR_COPYFROMRESOURCE = &H4000
Public Const LR_COPYRETURNORG = &H4
Public Const LR_CREATEDIBSECTION = &H2000
Public Const LR_DEFAULTCOLOR = &H0
Public Const LR_DEFAULTSIZE = &H40
Public Const LR_LOADFROMFILE = &H10
Public Const LR_LOADMAP3DCOLORS = &H1000
Public Const LR_LOADTRANSPARENT = &H20
Public Const LR_MONOCHROME = &H1
Public Const LR_SHARED = &H8000

' Creates a memory DC
Declare Function CreateCompatibleDC Lib "gdi32" ( _
    ByVal hDC As Long _
    ) As Long
' Creates a bitmap in memory:
Declare Function CreateCompatibleBitmap Lib "gdi32" ( _
    ByVal hDC As Long, _
    ByVal nWidth As Long, ByVal nHeight As Long _
    ) As Long
' Places a GDI object into DC, returning the previous one:
Declare Function SelectObject Lib "gdi32" _
    (ByVal hDC As Long, ByVal hObject As Long _
    ) As Long
' Deletes a GDI object:
Declare Function DeleteObject Lib "gdi32" _
    (ByVal hObject As Long _
    ) As Long
' Copies Bitmaps from one DC to another, can also perform
' raster operations during the transfer:
Declare Function BitBlt Lib "gdi32" ( _
    ByVal hDestDC As Long, _
    ByVal X As Long, ByVal Y As Long, _
    ByVal nWidth As Long, ByVal nHeight As Long, _
    ByVal hSrcDC As Long, _
    ByVal xSrc As Long, ByVal ySrc As Long, _
    ByVal dwRop As Long _
    ) As Long
Public Const SRCCOPY = &HCC0020
Public Const SRCAND = &H8800C6
Public Const SRCPAINT = &HEE0086
Public Const SRCINVERT = &H660046

' Structure used to hold bitmap information about Bitmaps
' created using GDI in memory:
Type Bitmap
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type
' Get information relating to a GDI Object
Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" ( _
    ByVal hObject As Long, _
    ByVal nCount As Long, _
    lpObject As Any _
    ) As Long
' The traditional rectangle structure:
Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
' Fills a rectangle in a DC with a specified brush
Declare Function FillRect Lib "user32" ( _
    ByVal hDC As Long, _
    lpRect As RECT, _
    ByVal hBrush As Long _
    ) As Long
' Create a brush of a certain colour:
Declare Function CreateSolidBrush Lib "gdi32" ( _
    ByVal crColor As Long _
    ) As Long
Declare Function SetBkColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Declare Function GetDesktopWindow Lib "user32" () As Long
Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hDC As Long) As Long

Public Function GDIMakeDCAndBitmap( _
        ByVal bMono As Boolean, _
        ByRef hDC As Long, _
        ByRef hBMP As Long, _
        ByRef hBmpOld As Long, _
        ByVal lDX As Long, _
        ByVal lDY As Long _
    ) As Boolean
' **********************************************************
' GDI Helper function: Makes a bitmap of a specified size
' and creates a DC to hold it.
' **********************************************************
Dim lCDC As Long
Dim lhWnd As Long

    ' Initialise byref variables:
    hDC = 0: hBMP = 0: hBmpOld = 0
    ' Create the DC from the basis DC:
    If (bMono) Then
        lCDC = 0
    Else
        lhWnd = GetDesktopWindow()
        lCDC = GetDC(lhWnd)
    End If
    hDC = CreateCompatibleDC(lCDC)
    If (bMono) Then
        lCDC = hDC
    End If
    
    If (hDC <> 0) Then
        ' If we get one, then time to make the bitmap:
        hBMP = CreateCompatibleBitmap(lCDC, lDX, lDY)
        ' If we succeed in creating the bitmap:
        If (hBMP <> 0) Then
            ' Select the bitmap into the memory DC and
            ' store the bitmap that was there before (need
            ' to do this because you need to select this
            ' bitmap back into the DC before deleting
            ' the new Bitmap):
            hBmpOld = SelectObject(hDC, hBMP)
            ' Success:
            GDIMakeDCAndBitmap = True
        End If
    End If
    
    If Not (bMono) Then
        ReleaseDC lhWnd, lCDC
    End If

End Function
Public Function GDILoadBitmapIntoDC( _
        ByVal bMono As Boolean, _
        ByVal sFileName As String, _
        ByRef hDC As Long, _
        ByRef hBMP As Long, _
        ByRef hBmpOld As Long _
    ) As Boolean
' **********************************************************
' GDI Helper function: Loads a bitmap from file and selects
' it into a memory DC.
' **********************************************************
Dim hInst As Long
Dim hDCBasis As Long
Dim lhWnd As Long

    ' Initialise byref variables:
    hDC = 0: hBMP = 0: hBmpOld = 0
    
    ' Now load the sprite bitmap:
    hInst = App.hInstance
    ' This is the quick, direct way where we don't get
    ' any extra copies of the bitmaps, as compared to
    ' using the VB picture object:
    hBMP = LoadImage(hInst, sFileName, IMAGE_BITMAP, 0, 0, LR_LOADFROMFILE)
    If (hBMP <> 0) Then
        ' Create a DC to hold the sprite, and select
        ' the sprite into it:
        If (bMono) Then
            hDCBasis = 0
        Else
            lhWnd = GetDesktopWindow()
            hDCBasis = GetDC(lhWnd)
        End If
        hDC = CreateCompatibleDC(hDCBasis)
        If (hDC <> 0) Then
            ' If DC Is created, select the bitmap into it:
            hBmpOld = SelectObject(hDC, hBMP)
            GDILoadBitmapIntoDC = True
        End If
        If Not (bMono) Then
            ReleaseDC lhWnd, hDCBasis
        End If
    End If

End Function
Public Function GDILoadPictureIntoDC( _
        ByVal bMono As Boolean, _
        ByRef oPic As StdPicture, _
        ByRef hDC As Long, _
        ByRef hBMP As Long, _
        ByRef hBmpOld As Long _
    ) As Boolean
' **********************************************************
' GDI Helper function: Creates a memory DC containing a new
' copy of bitmap from a StdPicture.
' **********************************************************
Dim hInst As Long
Dim hDCBasis As Long
Dim lhWnd As Long
Dim hDCTemp As Long
Dim hBmpTemp As Long
Dim hBmpTempOld As Long

    ' Initialise byref variables:
    hDC = 0: hBMP = 0: hBmpOld = 0
        
    ' Create a DC to hold the sprite, and select
    ' the sprite into it:
    If (bMono) Then
        hDCBasis = 0
    Else
        lhWnd = GetDesktopWindow()
        hDCBasis = GetDC(lhWnd)
    End If
    hDCTemp = CreateCompatibleDC(hDCBasis)
    If (bMono) Then
        hDCBasis = hDCTemp
    End If
    
    If (hDCTemp <> 0) Then
        hBmpTempOld = SelectObject(hDCTemp, oPic.Handle)
    
        hDC = CreateCompatibleDC(hDCBasis)
        If (hDC <> 0) Then
            ' If we get one, then time to make the bitmap:
            Dim tBM As Bitmap
            GetObjectAPI oPic.Handle, Len(tBM), tBM
            
            hBMP = CreateCompatibleBitmap(hDCBasis, tBM.bmWidth, tBM.bmHeight)
            If (hBMP <> 0) Then
                hBmpOld = SelectObject(hDC, hBMP)
                
                BitBlt hDC, 0, 0, tBM.bmWidth, tBM.bmHeight, hDCTemp, 0, 0, SRCCOPY
                
                GDILoadPictureIntoDC = True
            End If
        End If
        
        SelectObject hDCTemp, hBmpTempOld
        DeleteObject hDCTemp
        
    End If
    If Not (bMono) Then
        ReleaseDC lhWnd, hDCBasis
    End If

End Function

Public Sub GDIClearDCBitmap( _
        ByRef hDC As Long, _
        ByRef hBMP As Long, _
        ByVal hBmpOld As Long _
    )
' **********************************************************
' GDI Helper function: Goes through the steps required
' to clear up a bitmap within a DC.
' **********************************************************
    ' If we have a valid DC:
    If (hDC <> 0) Then
        ' If there is a valid bitmap in it:
        If (hBMP <> 0) Then
            ' Select the original bitmap into the DC:
            SelectObject hDC, hBmpOld
            ' Now delete the unreferenced bitmap:
            DeleteObject hBMP
            ' Byref so set the value to invalid BMP:
            hBMP = 0
        End If
        ' Delete the memory DC:
        DeleteObject hDC
        ' Byref so set the value to invalid DC:
        hDC = 0
    End If
End Sub


