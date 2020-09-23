Attribute VB_Name = "modBitmap"
Option Explicit
    
Private Type BITMAP
  bmType As Long
  bmWidth As Long
  bmHeight As Long
  bmWidthBytes As Long
  bmPlanes As Integer
  bmBitsPixel As Integer
  bmBits As Long
End Type

Private Type BITMAPINFOHEADER
   biSize As Long
   biWidth As Long
   biHeight As Long
   biPlanes As Integer
   biBitCount As Integer
   biCompression As Long
   biSizeImage As Long
   biXPelsPerMeter As Long
   biYPelsPerMeter As Long
   biClrUsed As Long
   biClrImportant As Long
End Type

Public Type BITMAPINFO
  bmiHeader As BITMAPINFOHEADER
  bmiColors(0 To 255) As RGBQUAD 'bmiColors As RGBQUAD
End Type

Const BI_RGB = 0&
Public Const DIB_RGB_COLORS = 0&


Public Const PICTYPE_UNINITIALIZED = -1
Public Const PICTYPE_NONE = 0
Public Const PICTYPE_BITMAP = 1
Public Const PICTYPE_METAFILE = 2
Public Const PICTYPE_ICON = 3
Public Const PICTYPE_ENHMETAFILE = 4

Public Type GUID
    Data1           As Long
    Data2           As Integer
    Data3           As Integer
    Data4(0 To 7)   As Byte
End Type

Public Type PICTDESC
    cbSizeOfStruct  As Long
    PicType         As Long
    hBitmap         As Long
    hPal            As Long
End Type

Private Type IID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

Private Declare Sub OleCreatePictureIndirect Lib "olepro32.dll" (lpPictDesc As PICTDESC, riid As GUID, ByVal fPictureOwnsHandle As Long, ipic As IPicture)
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function SetDIBits Lib "gdi32" (ByVal hdc As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Function StretchDIBits Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal wSrcWidth As Long, ByVal wSrcHeight As Long, lpBits As Any, lpBitsInfo As BITMAPINFO, ByVal wUsage As Long, ByVal dwRop As Long) As Long

Public Sub gSetBitsrgb(pic As Picture, bits() As RGBCOLOR, Optional ByVal channels As Integer = 4, Optional ByVal bitcount As Integer = 32, Optional ByVal compression As Long = BI_RGB)
  Dim hdcNew As Long
  Dim oldhand As Long
  Dim BytesPerScanLine As Long
  Dim PadBytesPerScanLine As Long
  Dim PicInfo As BITMAP
  Dim DIBInfo As BITMAPINFO
  
  hdcNew = CreateCompatibleDC(0&)
  oldhand = SelectObject(hdcNew, pic.Handle)
  
  Call GetObject(pic.Handle, Len(PicInfo), PicInfo)
  With DIBInfo.bmiHeader
    .biSize = 40
    .biWidth = PicInfo.bmWidth
    .biHeight = -PicInfo.bmHeight
    .biPlanes = 1
    .biBitCount = bitcount
    .biCompression = compression
    BytesPerScanLine = ((((.biWidth * bitcount) + 31) \ 32) * channels)
    PadBytesPerScanLine = _
       BytesPerScanLine - (((.biWidth * bitcount) + 7) \ 8)
    .biSizeImage = BytesPerScanLine * Abs(.biHeight)
  End With
  SetDIBits hdcNew, pic.Handle, 0, PicInfo.bmHeight, bits(LBound(bits, 1), LBound(bits, 2)), DIBInfo, DIB_RGB_COLORS

  SelectObject hdcNew, oldhand
  DeleteDC hdcNew
End Sub

Public Sub gSetBits(pic As Picture, bits() As RGBQUAD, Optional ByVal channels As Integer = 4, Optional ByVal bitcount As Integer = 32, Optional ByVal compression As Long = BI_RGB)
  Dim hdcNew As Long
  Dim oldhand As Long
  Dim BytesPerScanLine As Long
  Dim PadBytesPerScanLine As Long
  Dim PicInfo As BITMAP
  Dim DIBInfo As BITMAPINFO
  
  hdcNew = CreateCompatibleDC(0&)
  oldhand = SelectObject(hdcNew, pic.Handle)
  
  Call GetObject(pic.Handle, Len(PicInfo), PicInfo)
  With DIBInfo.bmiHeader
    .biSize = 40
    .biWidth = PicInfo.bmWidth
    .biHeight = -PicInfo.bmHeight
    .biPlanes = 1
    .biBitCount = bitcount
    .biCompression = compression
    BytesPerScanLine = ((((.biWidth * .biBitCount) + 31) \ 32) * channels)
    PadBytesPerScanLine = _
       BytesPerScanLine - (((.biWidth * .biBitCount) + 7) \ 8)
    .biSizeImage = BytesPerScanLine * Abs(.biHeight)
  End With
  
  SetDIBits hdcNew, pic.Handle, 0, PicInfo.bmHeight, bits(LBound(bits, 1), LBound(bits, 2)), DIBInfo, DIB_RGB_COLORS

  SelectObject hdcNew, oldhand
  DeleteDC hdcNew
End Sub

Public Function gGetBits(pic As Picture) As RGBQUAD()
  Dim hdcNew As Long
  Dim oldhand As Long
  Dim BytesPerScanLine As Long
  Dim PadBytesPerScanLine As Long
  Dim PicInfo As BITMAP
  Dim DIBInfo As BITMAPINFO
  Dim ret As Long
  Dim b() As RGBQUAD
    
  hdcNew = CreateCompatibleDC(0&)
  oldhand = SelectObject(hdcNew, pic.Handle)
  Call GetObject(pic.Handle, Len(PicInfo), PicInfo)
  With DIBInfo.bmiHeader
    .biSize = 40
    .biWidth = PicInfo.bmWidth
    .biHeight = -PicInfo.bmHeight     'bottom up scan line is now inverted
    .biPlanes = 1
    .biBitCount = 32
    .biCompression = BI_RGB
    BytesPerScanLine = ((((.biWidth * .biBitCount) + 31) \ 32) * 4)
    PadBytesPerScanLine = _
       BytesPerScanLine - (((.biWidth * .biBitCount) + 7) \ 8)
    .biSizeImage = BytesPerScanLine * Abs(.biHeight)
  End With
  'redimension the array (RGBQUAD) ...
  ReDim b(0 To PicInfo.bmWidth - 1, 0 To PicInfo.bmHeight - 1) As RGBQUAD
  'get picture data ...
  ret = GetDIBits(hdcNew, pic.Handle, 0, PicInfo.bmHeight, b(0, 0), DIBInfo, DIB_RGB_COLORS)
  SelectObject hdcNew, oldhand
  DeleteDC hdcNew
  
  gGetBits = b
End Function

Public Function gGetBitsFlat(pic As Picture) As Long()
  Dim hdcNew As Long
  Dim oldhand As Long
  Dim BytesPerScanLine As Long
  Dim PadBytesPerScanLine As Long
  Dim PicInfo As BITMAP
  Dim DIBInfo As BITMAPINFO
  Dim ret As Long
  Dim b() As Long
    
  hdcNew = CreateCompatibleDC(0&)
  oldhand = SelectObject(hdcNew, pic.Handle)
  Call GetObject(pic.Handle, Len(PicInfo), PicInfo)
  With DIBInfo.bmiHeader
    .biSize = 40
    .biWidth = PicInfo.bmWidth
    .biHeight = -PicInfo.bmHeight     'bottom up scan line is now inverted
    .biPlanes = 1
    .biBitCount = 32
    .biCompression = BI_RGB
    BytesPerScanLine = ((((.biWidth * .biBitCount) + 31) \ 32) * 4)
    PadBytesPerScanLine = _
       BytesPerScanLine - (((.biWidth * .biBitCount) + 7) \ 8)
    .biSizeImage = BytesPerScanLine * Abs(.biHeight)
  End With
  'redimension the array (RGBQUAD) ...
  ReDim b(0 To PicInfo.bmWidth - 1, 0 To PicInfo.bmHeight - 1) As Long
  'get picture data ...
  ret = GetDIBits(hdcNew, pic.Handle, 0, PicInfo.bmHeight, b(0, 0), DIBInfo, DIB_RGB_COLORS)
  SelectObject hdcNew, oldhand
  DeleteDC hdcNew
  
  gGetBitsFlat = b
End Function

Public Sub gSetBitsFlat(ByVal hbm As Long, bits() As Byte, Optional ByVal channels As Integer = 4, Optional ByVal bpp As Integer = 32, Optional ByVal compression As Long = BI_RGB)
  Dim hdcNew As Long
  Dim oldhand As Long
  Dim BytesPerScanLine As Long
  Dim PadBytesPerScanLine As Long
  Dim PicInfo As BITMAP
  Dim DIBInfo As BITMAPINFO
  
  hdcNew = CreateCompatibleDC(0&)
  oldhand = SelectObject(hdcNew, hbm)
  
  Call GetObject(hbm, Len(PicInfo), PicInfo)
  With DIBInfo.bmiHeader
    .biSize = 40
    .biWidth = PicInfo.bmWidth
    .biHeight = -PicInfo.bmHeight
    .biPlanes = 1
    .biBitCount = bpp
    .biCompression = compression
    BytesPerScanLine = ((((.biWidth * .biBitCount) + 31) \ 32) * channels)
    PadBytesPerScanLine = _
       BytesPerScanLine - (((.biWidth * .biBitCount) + 7) \ 8)
    .biSizeImage = BytesPerScanLine * Abs(.biHeight)
  End With
  SetDIBits hdcNew, hbm, 0, PicInfo.bmHeight, bits(LBound(bits, 1)), DIBInfo, DIB_RGB_COLORS

  SelectObject hdcNew, oldhand
  DeleteDC hdcNew
End Sub

Public Sub gSetBitsHandle(ByVal hbm As Long, bits() As RGBQUAD)
  Dim hdcNew As Long
  Dim oldhand As Long
  Dim BytesPerScanLine As Long
  Dim PadBytesPerScanLine As Long
  Dim PicInfo As BITMAP
  Dim DIBInfo As BITMAPINFO
  
  hdcNew = CreateCompatibleDC(0&)
  oldhand = SelectObject(hdcNew, hbm)
  
  Call GetObject(hbm, Len(PicInfo), PicInfo)
  With DIBInfo.bmiHeader
    .biSize = 40
    .biWidth = PicInfo.bmWidth
    .biHeight = -PicInfo.bmHeight
    .biPlanes = 1
    .biBitCount = 32
    .biCompression = BI_RGB
    BytesPerScanLine = ((((.biWidth * .biBitCount) + 31) \ 32) * 4)
    PadBytesPerScanLine = _
       BytesPerScanLine - (((.biWidth * .biBitCount) + 7) \ 8)
    .biSizeImage = BytesPerScanLine * Abs(.biHeight)
  End With
  SetDIBits hdcNew, hbm, 0, PicInfo.bmHeight, bits(0, 0), DIBInfo, DIB_RGB_COLORS

  SelectObject hdcNew, oldhand
  DeleteDC hdcNew
End Sub

Private Function BitmapToPicture(ByVal hbmp As Long, Optional ByVal hPal As Long = 0, Optional PicType = PICTYPE_BITMAP) As Picture
'
'This routine will take a device dependant bitmap and encapsulate it into
'an OLE Picture object.  Much thanks to Bruce McKinney for personally helping
'with this.
'
    Dim Picture     As Picture
    Dim PICTDESC    As PICTDESC
    Dim IID         As GUID
    
    PICTDESC.cbSizeOfStruct = Len(PICTDESC)     '16 Bytes
    PICTDESC.PicType = PicType
    PICTDESC.hBitmap = hbmp
    PICTDESC.hPal = hPal
    
    
    'IPicture GUID {7BF80980-BF32-101A-8BBB-00AA00300CAB}
    IID.Data1 = &H7BF80980
    IID.Data2 = &HBF32
    IID.Data3 = &H101A
    IID.Data4(0) = &H8B
    IID.Data4(1) = &HBB
    IID.Data4(2) = &H0
    IID.Data4(3) = &HAA
    IID.Data4(4) = &H0
    IID.Data4(5) = &H30
    IID.Data4(6) = &HC
    IID.Data4(7) = &HAB
    
    ' Create picture from bitmap handle
    OleCreatePictureIndirect PICTDESC, IID, True, Picture
    
    ' Result will be valid Picture or Nothing--either way set it
    Set BitmapToPicture = Picture
End Function

Function ClonePicture(pic As Picture) As Picture
    ' make a true copy of the picture instead of just another reference pointer ...
    Dim img() As RGBQUAD
    Dim hbm As Long
    Dim hdc As Long
    Dim hbmold As Long
    
    ' make a memory dc for the original picture ...
    hdc = CreateCompatibleDC(0&)
    ' select the picture into it ...
    hbmold = SelectObject(hdc, pic.Handle)
    ' get the bits for the original picture ...
    img = gGetBits(pic)
    ' now create a compatible bitmap for our clone picture - this gives a
    ' color one instead of the otherwise mono one cos the picture is selected
    ' into the dc we are using as a compatibility model ...
    hbm = CreateCompatibleBitmap(hdc, UBound(img, 1) + 1, UBound(img, 2) + 1)
    ' now set the original picture bits into the copy ...
    gSetBitsHandle hbm, img
    ' and make an IPicture from it ...
    Set ClonePicture = BitmapToPicture(hbm)
    ' finally release the picture from its' dc and destroy the dc ...
    SelectObject hdc, hbmold
    DeleteDC hdc
End Function

Function CloneBits(bits() As RGBQUAD) As RGBQUAD()
    Dim ret() As RGBQUAD
    Dim x As Integer
    Dim y As Integer
    
    ReDim ret(LBound(bits, 1) To UBound(bits, 1), LBound(bits, 2) To UBound(bits, 2)) As RGBQUAD
    For x = LBound(bits, 1) To UBound(bits, 1)
        For y = LBound(bits, 2) To UBound(bits, 2)
            ret(x, y) = bits(x, y)
        Next y
    Next x
    
    CloneBits = ret
End Function

Function CloneBytes(byts() As Byte) As Byte()
    Dim ret() As Byte
    Dim x As Integer
    Dim y As Integer
    
    ReDim ret(LBound(byts, 1) To UBound(byts, 1), LBound(byts, 2) To UBound(byts, 2)) As Byte
    For x = LBound(byts, 1) To UBound(byts, 1)
        For y = LBound(byts, 2) To UBound(byts, 2)
            ret(x, y) = byts(x, y)
        Next y
    Next x
    
    CloneBytes = ret
End Function

