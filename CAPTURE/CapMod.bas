Attribute VB_Name = "CapMod"
Option Explicit
Public Type RECT
    Left                         As Long
    Top                          As Long
    Right                        As Long
    Bottom                       As Long
End Type
Public Type PALETTEENTRY
    peRed                        As Byte
    peGreen                      As Byte
    peBlue                       As Byte
    peFlags                      As Byte
End Type
Public Type LOGPALETTE
    palVersion                   As Integer
    palNumEntries                As Integer
    palPalEntry(255)             As PALETTEENTRY    ' Enough for 256 colors.
End Type
Public Type GUID
    Data1                        As Long
    Data2                        As Integer
    Data3                        As Integer
    Data4(7)                     As Byte
End Type
Private Const RASTERCAPS     As Long = 38
Private Const RC_PALETTE     As Long = &H100
Private Const SIZEPALETTE    As Long = 104
'Public Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Public Type PicBmp
    size                         As Long
Type                         As Long
    hBmp                         As Long
    hpal                         As Long
    Reserved                     As Long
End Type
Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (PicDesc As PicBmp, _
                                                                      RefIID As GUID, _
                                                                      ByVal fPictureOwnsHandle As Long, _
                                                                      IPic As IPicture) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, _
                                                             ByVal nWidth As Long, _
                                                             ByVal nHeight As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, _
                                                    ByVal iCapabilitiy As Long) As Long
Private Declare Function GetSystemPaletteEntries Lib "gdi32" (ByVal hdc As Long, _
                                                              ByVal wStartIndex As Long, _
                                                              ByVal wNumEntries As Long, _
                                                              lpPaletteEntries As PALETTEENTRY) As Long
Private Declare Function CreatePalette Lib "gdi32" (lpLogPalette As LOGPALETTE) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, _
                                                   ByVal hObject As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDCDest As Long, _
                                             ByVal XDest As Long, _
                                             ByVal YDest As Long, _
                                             ByVal nWidth As Long, _
                                             ByVal nHeight As Long, _
                                             ByVal hDCSrc As Long, _
                                             ByVal xSrc As Long, _
                                             ByVal ySrc As Long, _
                                             ByVal dwRop As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function SelectPalette Lib "gdi32" (ByVal hdc As Long, _
                                                    ByVal hPalette As Long, _
                                                    ByVal bForceBackground As Long) As Long
Private Declare Function RealizePalette Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, _
                                                 ByVal hdc As Long) As Long

Public Sub Capture(picBox As PictureBox)

    On Error GoTo Pesan
    Clipboard.Clear
    Set picBox.Picture = CaptureClient(picBox) 'CaptureScreen()
    Clipboard.SetData picBox.Image
Pesan:
    If Err.Number <> 0 Then
        MsgBox Err.Number & "=  " & Err.Description, vbOKOnly, "CAPTURE ERROR"
    End If

End Sub

Public Function CaptureClient(picBox As PictureBox) As Picture

    Set CaptureClient = CaptureWindow(picBox.hwnd, True, 0, 0, picBox.ScaleX(picBox.ScaleWidth, picBox.ScaleMode, vbPixels), picBox.ScaleY(picBox.ScaleHeight, picBox.ScaleMode, vbPixels))

End Function

Public Function CaptureWindow(ByVal hWndSrc As Long, _
                              ByVal Client As Boolean, _
                              ByVal LeftSrc As Long, _
                              ByVal TopSrc As Long, _
                              ByVal WidthSrc As Long, _
                              ByVal HeightSrc As Long) As Picture

Dim hDCMemory       As Long
Dim hBmp            As Long
Dim hBmpPrev        As Long
Dim hDCSrc          As Long
Dim hpal            As Long
Dim hPalPrev        As Long
Dim RasterCapsScrn  As Long
Dim HasPaletteScrn  As Long
Dim PaletteSizeScrn As Long
Dim LogPal          As LOGPALETTE

    If Client Then
        hDCSrc = GetDC(hWndSrc)
    Else 'CLIENT = FALSE/0
        hDCSrc = GetWindowDC(hWndSrc)
    End If
    hDCMemory = CreateCompatibleDC(hDCSrc)
    hBmp = CreateCompatibleBitmap(hDCSrc, WidthSrc, HeightSrc)
    hBmpPrev = SelectObject(hDCMemory, hBmp)
    RasterCapsScrn = GetDeviceCaps(hDCSrc, RASTERCAPS)
    HasPaletteScrn = RasterCapsScrn And RC_PALETTE
    PaletteSizeScrn = GetDeviceCaps(hDCSrc, SIZEPALETTE)
    If HasPaletteScrn And (PaletteSizeScrn = 256) Then
        With LogPal
            .palVersion = &H300
            .palNumEntries = 256
            GetSystemPaletteEntries hDCSrc, 0, 256, .palPalEntry(0)
        End With 'LOGPAL
        hpal = CreatePalette(LogPal)
        hPalPrev = SelectPalette(hDCMemory, hpal, 0)
        RealizePalette hDCMemory
    End If
    BitBlt hDCMemory, 0, 0, WidthSrc, HeightSrc, hDCSrc, LeftSrc, TopSrc, vbSrcCopy
    hBmp = SelectObject(hDCMemory, hBmpPrev)
    If HasPaletteScrn And (PaletteSizeScrn = 256) Then
        hpal = SelectPalette(hDCMemory, hPalPrev, 0)
    End If
    DeleteDC hDCMemory
    ReleaseDC hWndSrc, hDCSrc
    Set CaptureWindow = CreateBitmapPicture(hBmp, hpal)

End Function

Public Function CreateBitmapPicture(ByVal hBmp As Long, _
                                    ByVal hpal As Long) As Picture

Dim Pic           As PicBmp
Dim IPic          As IPicture
Dim IID_IDispatch As GUID

    With IID_IDispatch
        .Data1 = &H20400
        .Data4(0) = &HC0
        .Data4(7) = &H46
    End With 'IID_IDISPATCH
    With Pic
        .size = Len(Pic)
        .Type = vbPicTypeBitmap
        .hBmp = hBmp
        .hpal = hpal
    End With 'PIC
    OleCreatePictureIndirect Pic, IID_IDispatch, 1, IPic
    Set CreateBitmapPicture = IPic

End Function
