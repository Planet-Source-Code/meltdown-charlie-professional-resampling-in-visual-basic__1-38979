VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form fMain 
   Caption         =   "Resampling Sample Project "
   ClientHeight    =   6780
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   10770
   Icon            =   "resamp.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   452
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   718
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog filer 
      Left            =   4620
      Top             =   6300
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DefaultExt      =   "bmp"
      DialogTitle     =   "Select an image to load"
      Filter          =   "Bitmaps|*.bmp|GIF's|*.gif|JPegs|*.jpg;*.jpeg"
   End
   Begin VB.TextBox tHeight 
      Height          =   285
      Left            =   2085
      TabIndex        =   23
      Text            =   "403"
      Top             =   6435
      Width           =   585
   End
   Begin VB.TextBox tWidth 
      Height          =   285
      Left            =   720
      TabIndex        =   20
      Text            =   "583"
      Top             =   6435
      Width           =   570
   End
   Begin VB.Frame Frame1 
      Caption         =   "Method"
      Height          =   4740
      Left            =   105
      TabIndex        =   3
      Top             =   1560
      Width           =   1470
      Begin VB.OptionButton optMethod 
         Caption         =   "Triangle"
         Height          =   210
         Index           =   15
         Left            =   135
         TabIndex        =   19
         Top             =   4320
         Width           =   1230
      End
      Begin VB.OptionButton optMethod 
         Caption         =   "Bell"
         Height          =   210
         Index           =   14
         Left            =   135
         TabIndex        =   18
         Top             =   4065
         Value           =   -1  'True
         Width           =   1230
      End
      Begin VB.OptionButton optMethod 
         Caption         =   "Spline"
         Height          =   210
         Index           =   13
         Left            =   135
         TabIndex        =   17
         Top             =   3810
         Width           =   1230
      End
      Begin VB.OptionButton optMethod 
         Caption         =   "Catrom"
         Height          =   210
         Index           =   12
         Left            =   135
         TabIndex        =   16
         Top             =   3540
         Width           =   1230
      End
      Begin VB.OptionButton optMethod 
         Caption         =   "Cubic"
         Height          =   210
         Index           =   11
         Left            =   135
         TabIndex        =   15
         Top             =   3285
         Width           =   1230
      End
      Begin VB.OptionButton optMethod 
         Caption         =   "Hermite"
         Height          =   210
         Index           =   10
         Left            =   135
         TabIndex        =   14
         Top             =   3015
         Width           =   1230
      End
      Begin VB.OptionButton optMethod 
         Caption         =   "Lanczos"
         Height          =   210
         Index           =   9
         Left            =   135
         TabIndex        =   13
         Top             =   2775
         Width           =   1230
      End
      Begin VB.OptionButton optMethod 
         Caption         =   "Mitchell"
         Height          =   210
         Index           =   8
         Left            =   135
         TabIndex        =   12
         Top             =   2520
         Width           =   1230
      End
      Begin VB.OptionButton optMethod 
         Caption         =   "Hanning"
         Height          =   210
         Index           =   7
         Left            =   135
         TabIndex        =   11
         Top             =   2250
         Width           =   1230
      End
      Begin VB.OptionButton optMethod 
         Caption         =   "Quadratic"
         Height          =   210
         Index           =   6
         Left            =   135
         TabIndex        =   10
         Top             =   2010
         Width           =   1230
      End
      Begin VB.OptionButton optMethod 
         Caption         =   "Hamming"
         Height          =   210
         Index           =   5
         Left            =   135
         TabIndex        =   9
         Top             =   1710
         Width           =   1230
      End
      Begin VB.OptionButton optMethod 
         Caption         =   "Gaussian"
         Height          =   210
         Index           =   4
         Left            =   135
         TabIndex        =   8
         Top             =   1440
         Width           =   1230
      End
      Begin VB.OptionButton optMethod 
         Caption         =   "Blackman"
         Height          =   210
         Index           =   3
         Left            =   135
         TabIndex        =   7
         Top             =   1185
         Width           =   1230
      End
      Begin VB.OptionButton optMethod 
         Caption         =   "Bilinear"
         Height          =   210
         Index           =   2
         Left            =   135
         TabIndex        =   6
         Top             =   930
         Width           =   1230
      End
      Begin VB.OptionButton optMethod 
         Caption         =   "Box"
         Height          =   210
         Index           =   1
         Left            =   135
         TabIndex        =   5
         Top             =   630
         Width           =   1230
      End
      Begin VB.OptionButton optMethod 
         Caption         =   "Averaging"
         Height          =   210
         Index           =   0
         Left            =   135
         TabIndex        =   4
         Top             =   360
         Width           =   1230
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Resample"
      Height          =   315
      Left            =   3090
      TabIndex        =   1
      Top             =   6405
      Width           =   1380
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6075
      Left            =   1935
      ScaleHeight     =   405
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   585
      TabIndex        =   0
      Top             =   180
      Width           =   8775
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1200
      Left            =   165
      Picture         =   "resamp.frx":030A
      ScaleHeight     =   80
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   104
      TabIndex        =   2
      Top             =   240
      Width           =   1560
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Height"
      Height          =   195
      Left            =   1485
      TabIndex        =   22
      Top             =   6465
      Width           =   465
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Width"
      Height          =   195
      Left            =   120
      TabIndex        =   21
      Top             =   6450
      Width           =   420
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open..."
      End
      Begin VB.Menu mnuSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "&About"
   End
End
Attribute VB_Name = "fMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private curOption As Integer

Private Sub Command1_Click()
    Dim img() As RGBQUAD
    Dim ret() As RGBQUAD
    Dim ret2() As RGBQUAD
    Dim ret3() As RGBQUAD
    
    img = gGetBits(Picture2.Picture)
    
    If curOption = 0 Then
        ret = StdResize2(img, Picture2.ScaleWidth, Picture2.ScaleHeight, tWidth, tHeight)
    Else
    
        ret = AllocAndScale(img, Picture2.ScaleWidth, Picture2.ScaleHeight, tWidth, tHeight, curOption)
    End If
    ' set the result to picture 1
    Picture1 = LoadPicture("")
    Picture1.Width = tWidth
    Picture1.Height = tHeight
    Picture1.Picture = Picture1.Image
    Picture1.Refresh
    gSetBits Picture1.Picture, ret()
    Picture1.Refresh
    
    Erase img
    Erase ret
End Sub

Private Sub Form_Load()
    curOption = 14
End Sub

Private Sub mnuAbout_Click()
    fAbout.Show 1
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub mnuOpen_Click()
    ' open a file and display it in picture2 ...
On Error GoTo OpenError
    filer.InitDir = App.Path & "\"
    filer.ShowOpen
    Picture2 = LoadPicture(filer.FileName)
OpenError:
    
End Sub

Private Sub optMethod_Click(Index As Integer)
    curOption = Index
End Sub

Private Sub tWidth_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape, vbKeyBack, vbKeyDelete, vbKeyReturn
        Case Else
            KeyCode = 0
    End Select
End Sub

Private Sub tWidth_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("0") To Asc("9"), vbKeyEscape, vbKeyBack, vbKeyDelete, vbKeyReturn
        Case Else
            KeyAscii = 0
    End Select
End Sub
