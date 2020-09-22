Attribute VB_Name = "m_imutils"
' I am the black rainbow and i'm an ape of god
' I've got a face thats fit for violence upon
' And i'm a teen distortion, survived abortion
' A rebel from the waist down

' I never really hated one true god
' Just the god of the people i hated

Public Declare Function BitBlt Lib "gdi32" _
  (ByVal hDCDest As Long, ByVal XDest As Long, ByVal YDest As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hDCSrc As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal dwRop As Long) As Long

Public Declare Function GetDesktopWindow Lib _
   "user32" () As Long

Public Declare Function GetWindowDC Lib _
   "user32" (ByVal hWnd As Long) As Long

Public Declare Function ReleaseDC Lib "user32" _
   (ByVal hWnd As Long, ByVal hdc As Long) As Long
'--end block--'

Public Declare Function GrayScale Lib "ImageUtils.dll" (ByVal strInputFile As String, ByVal strOutputFile As String, ByVal blnEnableOverWrite As Boolean) As Integer
Public Declare Function getDesktop Lib "ImageUtils.dll" (ByVal strFileName As String, ByVal blnEnableOverWrite As Boolean, ByVal nWidth As Integer, ByVal nHeight As Integer, ByVal blnJpeg As Boolean, ByVal JPGCompressQuality As Integer) As Integer
Public Declare Function ConvertBMPtoJPG Lib "ImageUtils.dll" (ByVal strInputFile As String, ByVal strOutputFile As String, ByVal blnEnableOverWrite As Boolean, ByVal JPGCompressQuality As Integer, ByVal blnKeepBMP As Boolean) As Integer
Public Declare Function ConvertJPGtoBMP Lib "ImageUtils.dll" (ByVal strInputFile As String, ByVal strOutputFile As String, ByVal blnEnableOverWrite As Boolean, ByVal blnKeepJPG As Boolean) As Integer
Public Declare Function RotateRight Lib "ImageUtils.dll" (ByVal strInputFile As String, ByVal strOutputFile As String, ByVal blnEnableOverWrite As Boolean) As Integer
Public Declare Function RotateLeft Lib "ImageUtils.dll" (ByVal strInputFile As String, ByVal strOutputFile As String, ByVal blnEnableOverWrite As Boolean) As Integer
