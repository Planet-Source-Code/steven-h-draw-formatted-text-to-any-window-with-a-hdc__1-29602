VERSION 5.00
Begin VB.Form fMain 
   Caption         =   "cTextEx Demo"
   ClientHeight    =   4335
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8700
   LinkTopic       =   "Form1"
   ScaleHeight     =   4335
   ScaleWidth      =   8700
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cDskTop 
      Caption         =   "Desktop"
      Height          =   435
      Left            =   2130
      TabIndex        =   3
      Top             =   120
      Width           =   1215
   End
   Begin VB.PictureBox pTest 
      Height          =   1815
      Left            =   660
      ScaleHeight     =   1755
      ScaleWidth      =   7275
      TabIndex        =   2
      Top             =   960
      Width           =   7335
   End
   Begin VB.CommandButton cPBox 
      Caption         =   "PictureBox"
      Height          =   435
      Left            =   630
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cExit 
      Caption         =   "Exit"
      Height          =   435
      Left            =   7260
      TabIndex        =   0
      Top             =   3420
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   $"fMain.frx":0000
      ForeColor       =   &H000000FF&
      Height          =   1005
      Left            =   660
      TabIndex        =   4
      Top             =   3090
      Width           =   6405
   End
End
Attribute VB_Name = "fMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ===========================================================================
'
' Filename:    fMain
' Author:      Steven Higgan
' Date:        9 November 2001
'
' Requires:    cStdFontEx
'              cTextEx
'
' Platform: Built on Windows 2000 using Visual Basic 6 Enterprise
'
' Tested: NOT TESTED ON ANY OTHER PLATFORMS
'
' Pourpose: used to demonstrate the usage of the cStdFontEx and cTextEx
' classes
'
' Built on 9 Dec 2001 as part of research into my up coming Menu Object
' a full implementation of a Menu with no user controls, built compleatly
' from the ground up using the Windows API
'
' Other Peoples Code: None
'
' if you change this code, add more funcunalaty debug it on other windows
' platforms i would apreciate a copy of any changes made to the source code
'
' ===========================================================================
'
'   NOTE - if you have a slow computer the code in Form_Load may slow your
'   machine down if you have lots of fonts installed on your system,
'   Therefore it is recomended that you change the number of illiterations
'   BEFORE you email me complaining about "how slow your code is"
'
' ===========================================================================

Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long

Private m_aFonts() As String
Private m_lFACount As Long
  
Private TextEx As cTextEx
Private StdFontEx As cStdFontEx

Private Sub cDoIt_Click()
  
End Sub

Private Sub cDskTop_Click()

  Dim mRECT As RECT
  
  GetWindowRect GetDesktopWindow, mRECT
  
  With StdFontEx
'    Set .StdFont = lFont.Font
'    With .StdFont
'      .Name = RandomFontName
'      .Size = RandomFontSize
'    End With

    .Name = RandomFontName
    .Size = RandomFontSize
    .Bold = RandomBoolean
    .Italic = RandomBoolean
    .Strikethrough = RandomBoolean
    .UnderLine = RandomBoolean
    .Charset = 0
    .Colour = RandomFontColor
  End With
  
  With TextEx
    Set .StdFontEx = StdFontEx

    .AllignmentFlags = CENTER Or SINGLELINE Or VCENTER
    .hdc = GetWindowDC(GetDesktopWindow)

    .RectBottom = mRECT.Bottom
    .RectLeft = mRECT.Left
    .RectRight = mRECT.Right
    .RectTop = mRECT.Top

    .Text = "Some Text"
    
    .Draw
  End With
  
End Sub

Private Sub cExit_Click()

  Unload Me
  
End Sub

Private Sub cPBox_Click()

  Dim mRECT As RECT
  
  GetClientRect pTest.hwnd, mRECT
  
  With StdFontEx
'    Set .StdFont = lFont.Font
'    With .StdFont
'      .Name = RandomFontName
'      .Size = RandomFontSize
'    End With

    .Name = RandomFontName
    .Size = RandomFontSize
    .Bold = RandomBoolean
    .Italic = RandomBoolean
    .Strikethrough = RandomBoolean
    .UnderLine = RandomBoolean
    .Charset = 0
    .Colour = RandomFontColor
  End With
  
  With TextEx

    .RectBottom = mRECT.Bottom
    .RectLeft = mRECT.Left
    .RectRight = mRECT.Right
    .RectTop = mRECT.Top
    
    .Draw StdFontEx, , , , , pTest.hdc, "This Is Some Text", _
      CENTER Or SINGLELINE Or VCENTER
  End With
  
End Sub

Private Sub Form_Load()

  Dim I As Integer
  
  Set TextEx = New cTextEx
  Set StdFontEx = New cStdFontEx
  
  'NOTE - if you have a slow computer this code may slow your machine down
  'if you have lots of fonts installed on your system, Therefore it is
  'recomended that you change the number of illiterations BEFORE you
  'email me complaining about "how slow your code is"
  
  'Fill the Font Name array with some fonts
  For I = 0 To Screen.FontCount - 1
    ReDim Preserve m_aFonts(0 To m_lFACount)
    m_aFonts(I) = Screen.Fonts(I)
    m_lFACount = m_lFACount + 1
    Next I
  
End Sub

Private Sub Form_Unload(Cancel As Integer)

  Set TextEx = Nothing
  Set StdFontEx = Nothing
  
End Sub

Private Function RandomFontColor() As Long

  Randomize Timer

  RandomFontColor = RGB(Int((250 * Rnd) + 0), Int((250 * Rnd) + 0), Int((250 * Rnd) + 0))
  
End Function
Private Function RandomFontName() As String

  Randomize Timer
  
  RandomFontName = m_aFonts(Int((m_lFACount * Rnd) + 0))
  
End Function
Private Function RandomFontSize() As Integer

  Randomize Timer
  
  RandomFontSize = Int((100 * Rnd) + 1)
  
End Function
Private Function RandomBoolean() As Boolean

  Randomize Timer
  
  RandomBoolean = Int((1 * Rnd) + 0)
  
End Function
