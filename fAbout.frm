VERSION 5.00
Begin VB.Form fAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About "
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "fAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Height          =   420
      Left            =   2850
      TabIndex        =   2
      Top             =   1110
      Width           =   1485
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "VB-IP"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      Left            =   2940
      TabIndex        =   3
      Top             =   120
      Width           =   1560
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "www25.brinkster.com/mferris"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   300
      Left            =   555
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   2640
      Width           =   3450
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"fAbout.frx":030A
      Height          =   645
      Left            =   165
      TabIndex        =   0
      Top             =   1830
      Width           =   4320
   End
   Begin VB.Image Image1 
      Height          =   1500
      Left            =   120
      Picture         =   "fAbout.frx":03B0
      Top             =   165
      Width           =   2250
   End
End
Attribute VB_Name = "fAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ===================================================================================
' ===================================================================================
' This source code was provided to you by the Visual Basic Image Processing
' site - visit us on the web at www.intactinteractive.com
' ===================================================================================
' ===================================================================================
'
' This source code is provided as is with no warrany expressed or implied.
' You use this code at your own risk, Intact Interactive and Malcolm Ferris
' assume no responsibility for any loss or damage occuring as a result of
' the misuse of this code.
'
' This source code is copyrighted to Malcolm Ferris, and may be altered, and
' included, (altered or as is) in any work of your own without royalty or
' obligation. It is requested however that you retain all copyright notices on
' the source code and that you give appropriate mention in your program credits
' to the original author.
'
' Portions of this code were adapted from works derived from a variety of sources
' including :
'
'    The Windows 200 Graphics API - Black Book by Damon Chandler & Michael Fotsch
'    Digital Imaging  by Howard E. Burdick
'    www.vbaccelerator.com
'    Manuel Augusto Santos - source code posted on PSC for fast image access...
'
'
' ===================================================================================
' ===================================================================================
' Copyright M Ferris - 2000
' ===================================================================================
' ===================================================================================
Option Explicit

Const SW_SHOWMAXIMIZED = 3
Const SW_SHOWNORMAL = 1

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Label2.MouseIcon = LoadResPicture(101, vbResCursor)
End Sub

Private Sub Label2_Click()
    ' link to the site
    ShellExecute Me.hwnd, vbNullString, "http://www25.brinkster.com/mferris", vbNullString, vbNullString, SW_SHOWMAXIMIZED Or SW_SHOWNORMAL
End Sub
