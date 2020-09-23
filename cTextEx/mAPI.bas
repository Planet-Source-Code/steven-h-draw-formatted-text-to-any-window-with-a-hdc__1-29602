Attribute VB_Name = "mAPI"
' ===========================================================================
'
' Filename:    mAPI
' Author:      Steven Higgan
' Date:        9 November 2001
'
' Requires:    NOTHING
'
' Platform: Built on Windows 2000 using Visual Basic 6 Enterprise
'
' Tested: NOT TESTED ON ANY OTHER PLATFORMS
'
' Pourpose: used in cTextEx, Provides all the API declares used in cTextEx
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

Public Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
  Public Const BITSPIXEL = 12
  Public Const LOGPIXELSX = 88
  Public Const LOGPIXELSY = 90

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
Public Declare Function MulDiv Lib "kernel32" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long
Public Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Public Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
  Public Const NEWTRANSPARENT = 3

Public Declare Function GetTextColor Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
  Public Const CLR_INVALID = &HFFFF

Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long

Public Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
  Public Const DT_BOTTOM = &H8
  Public Const DT_CENTER = &H1
  Public Const DT_LEFT = &H0
  Public Const DT_CALCRECT = &H400
  Public Const DT_WORDBREAK = &H10
  Public Const DT_VCENTER = &H4
  Public Const DT_TOP = &H0
  Public Const DT_TABSTOP = &H80
  Public Const DT_SINGLELINE = &H20
  Public Const DT_RIGHT = &H2
  Public Const DT_NOCLIP = &H100
  Public Const DT_INTERNAL = &H1000
  Public Const DT_EXTERNALLEADING = &H200
  Public Const DT_EXPANDTABS = &H40
  Public Const DT_CHARSTREAM = 4
  Public Const DT_NOPREFIX = &H800

Public Const LF_FACESIZE = 32
Type LOGFONT
  lfHeight As Long
  lfWidth As Long
  lfEscapement As Long
  lfOrientation As Long
  lfWeight As Long
  lfItalic As Byte
  lfUnderline As Byte
  lfStrikeOut As Byte
  lfCharSet As Byte
  lfOutPrecision As Byte
  lfClipPrecision As Byte
  lfQuality As Byte
  lfPitchAndFamily As Byte
  lfFaceName(LF_FACESIZE) As Byte
End Type

Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type
