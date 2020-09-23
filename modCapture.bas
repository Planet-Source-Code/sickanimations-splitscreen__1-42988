Attribute VB_Name = "modCapture"
'The Code Below the line after the next line is
'by Sombody that goes by the name of "SniperElite" and has some posts on PSC
'\/ \/ \/ \/ \/ THIS IS NOT [s]Animation's (MY) CODE \/ \/ \/ \/ \/ \/

' This module contains several routines for capturing windows into a
' picture.  All the routines work on both 16 and 32 bit Windows
' platforms.
' The routines also have palette support.
'
' CreateBitmapPicture - Creates a picture object from a bitmap and
' palette
' CaptureWindow - Captures any window given a window handle
' CaptureActiveWindow - Captures the active window on the desktop
' CaptureForm - Captures the entire form
' CaptureClient - Captures the client area of a form
' CaptureScreen - Captures the entire screen
' PrintPictureToFitPage - prints any picture as big as possible on
' the page
' NOTES
' - No error trapping is included in these routines


Option Explicit
Option Base 0
Private Type PALETTEENTRY
    peRed As Byte
    peGreen As Byte
    peBlue As Byte
    peFlags As Byte
End Type

Private Type LOGPALETTE
    palVersion As Integer
    palNumEntries As Integer
    palPalEntry(255) As PALETTEENTRY
    ' Enough for 256 colors
End Type

Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(7) As Byte
End Type


Private Const RASTERCAPS As Long = 38
Private Const RC_PALETTE As Long = &H100
Private Const SIZEPALETTE As Long = 104
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

'Code reformatted by [s]Animations starts here
Private Declare Function CreateCompatibleDC Lib "GDI32" (ByVal hDC As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "GDI32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function GetDeviceCaps Lib "GDI32" (ByVal hDC As Long, ByVal iCapabilitiy As Long) As Long
Private Declare Function GetSystemPaletteEntries Lib "GDI32" (ByVal hDC As Long, ByVal wStartIndex As Long, _
ByVal wNumEntries As Long, lpPaletteEntries As PALETTEENTRY) As Long
Private Declare Function CreatePalette Lib "GDI32" (lpLogPalette As LOGPALETTE) As Long
Private Declare Function SelectObject Lib "GDI32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function BitBlt Lib "GDI32" (ByVal hDCDest As Long, ByVal XDest As Long, ByVal YDest As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hDCSrc As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal dwRop As Long) As Long
Private Declare Function DeleteDC Lib "GDI32" (ByVal hDC As Long) As Long
'\/ \/  [s]Animations changed from Private to Public
Public Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function SelectPalette Lib "GDI32" (ByVal hDC As Long, ByVal hPalette As Long, ByVal bForceBackground As Long) As Long
Private Declare Function RealizePalette Lib "GDI32" (ByVal hDC As Long) As Long
Private Declare Function GetWindowDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
'\/ \/  [s]Animations changed from Private to Public
Public Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (PicDesc As PicBmp, RefIID As GUID, ByVal fPictureOwnsHandle As Long, IPic As IPicture) As Long
'Code reformatted by [s]Animations ends here

Private Type PicBmp
    Size As Long
    Type As Long
    hBmp As Long
    hPal As Long
    Reserved As Long
End Type



Public Function CreateBitmapPicture(ByVal hBmp As Long, _
    ByVal hPal As Long) As Picture
    Dim r As Long

    Dim Pic As PicBmp
    ' IPicture requires a reference to "Standard OLE Types"
    Dim IPic As IPicture
    Dim IID_IDispatch As GUID
    ' Fill in with IDispatch Interface ID
    With IID_IDispatch
        .Data1 = &H20400
        .Data4(0) = &HC0
        .Data4(7) = &H46
    End With
    ' Fill Pic with necessary parts
    With Pic
        .Size = Len(Pic) ' Length of structure
        .Type = vbPicTypeBitmap ' Type of Picture (bitmap)
        .hBmp = hBmp ' Handle to bitmap
        .hPal = hPal ' Handle to palette (may be null)
    End With
    ' Create Picture object
    r = OleCreatePictureIndirect(Pic, IID_IDispatch, 1, IPic)
    ' Return the new Picture object
    Set CreateBitmapPicture = IPic
End Function

' CaptureWindow ' - Captures any portion of a window
' ' hWndSrc ' - Handle to the window to be captured
' ' Client ' - If True CaptureWindow captures from the client area of the
' window ' - If False CaptureWindow captures from the entire window
' ' LeftSrc, TopSrc, WidthSrc, HeightSrc ' - Specify the portion of the window to capture ' - Dimensions need to be specified in pixels
' ' Returns ' - Returns a Picture object containing a bitmap of the specified
' portion of the window that was captured
Public Function CaptureWindow(ByVal hWndSrc As Long, _
    ByVal Client As Boolean, ByVal LeftSrc As Long, _
    ByVal TopSrc As Long, ByVal WidthSrc As Long, _
    ByVal HeightSrc As Long) As Picture
    Dim hDCMemory As Long
    Dim hBmp As Long
    Dim hBmpPrev As Long
    Dim r As Long
    Dim hDCSrc As Long
    Dim hPal As Long
    Dim hPalPrev As Long
    Dim RasterCapsScrn As Long
    Dim HasPaletteScrn As Long
    Dim PaletteSizeScrn As Long
    Dim LogPal As LOGPALETTE ' Depending on the value of Client get the proper device context
    If Client Then
        hDCSrc = GetDC(hWndSrc) ' Get device context for client area
    Else
        hDCSrc = GetWindowDC(hWndSrc) ' Get device context for entire   window
    End If
    ' Create a memory device context for the copy process
    hDCMemory = CreateCompatibleDC(hDCSrc)
    ' Create a bitmap and place it in the memory DC
    hBmp = CreateCompatibleBitmap(hDCSrc, WidthSrc, HeightSrc)
    hBmpPrev = SelectObject(hDCMemory, hBmp)
    ' Get screen properties RasterCapsScrn = GetDeviceCaps(hDCSrc, RASTERCAPS) ' Raster 'capabilities
    HasPaletteScrn = RasterCapsScrn And RC_PALETTE ' Palette 'support
    PaletteSizeScrn = GetDeviceCaps(hDCSrc, SIZEPALETTE) ' Size of ' palette
    ' If the screen has a palette make a copy and realize it
    If HasPaletteScrn And (PaletteSizeScrn = 256) Then ' Create a copy of the system palette
        LogPal.palVersion = &H300
        LogPal.palNumEntries = 256
        r = GetSystemPaletteEntries(hDCSrc, 0, 256, _
            LogPal.palPalEntry(0))
        hPal = CreatePalette(LogPal) ' Select the new palette into the memory DC and realize it
        hPalPrev = SelectPalette(hDCMemory, hPal, 0)
        r = RealizePalette(hDCMemory)
    End If ' Copy the on-screen image into the memory DC
    r = BitBlt(hDCMemory, 0, 0, WidthSrc, HeightSrc, hDCSrc, _
        LeftSrc, TopSrc, vbSrcCopy)
    ' Remove the new copy of the on-screen image
    hBmp = SelectObject(hDCMemory, hBmpPrev)
    ' If the screen has a palette get back the palette that was
    ' selected in previously
    If HasPaletteScrn And (PaletteSizeScrn = 256) Then
        hPal = SelectPalette(hDCMemory, hPalPrev, 0)
    End If ' Release the device context resources back to the system
    r = DeleteDC(hDCMemory)
    r = ReleaseDC(hWndSrc, hDCSrc)
    ' Call CreateBitmapPicture to create a picture object from the
    ' bitmap and palette handles. Then return the resulting picture
    ' object.
    Set CaptureWindow = CreateBitmapPicture(hBmp, hPal)
End Function
