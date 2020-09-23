Attribute VB_Name = "Module1"
Option Explicit

Public Type RGBQUAD
        rgbBlue As Byte
        rgbGreen As Byte
        rgbRed As Byte
        rgbReserved As Byte
End Type

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

Public Type RGBCOLOR
    r As Byte
    g As Byte
    b As Byte
End Type

Public Declare Sub OleCreatePictureIndirect Lib "olepro32.dll" (lpPictDesc As PICTDESC, riid As GUID, ByVal fPictureOwnsHandle As Long, ipic As IPicture)
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long

