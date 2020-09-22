Attribute VB_Name = "mod_Gradient"
Option Explicit

Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Const CLR_INVALID = -1

Public Declare Function SetWindowRgn Lib "user32" _
    (ByVal hwnd As Long, ByVal hRgn As Long, _
    ByVal blnRedraw As Boolean) As Long

Public Declare Function CreateRoundRectRgn Lib "gdi32" _
    (ByVal RectX1 As Long, ByVal RectY1 As Long, ByVal RectX2 As Long, _
    ByVal RectY2 As Long, ByVal EllipseWidth As Long, _
    ByVal EllipseHeight As Long) As Long
    
Public Declare Function FillRect Lib "user32" ( _
   ByVal hDC As Long, lpRect As RECT, _
   ByVal hBrush As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" ( _
   ByVal hObject As Long) As Long

Public Declare Function CreateSolidBrush Lib "gdi32" ( _
   ByVal crColor As Long) As Long
   

Private Declare Function GradientFill Lib "msimg32" ( _
   ByVal hDC As Long, _
   pVertex As TRIVERTEX, _
   ByVal dwNumVertex As Long, _
   pMesh As GRADIENT_RECT, _
   ByVal dwNumMesh As Long, _
   ByVal dwMode As Long) As Long

Private Declare Function OleTranslateColor Lib "OLEPRO32.DLL" ( _
   ByVal OLE_COLOR As Long, _
   ByVal HPALETTE As Long, _
   pccolorref As Long) As Long
   
Public Type TRIVERTEX
   x As Long
   y As Long
   Red As Integer
   Green As Integer
   Blue As Integer
   alpha As Integer
End Type

Public Type GRADIENT_RECT
    UpperLeft As Long
    LowerRight As Long
End Type


Public Const LWA_COLORKEY = 1
Public Const LWA_ALPHA = 2
Public Const LWA_BOTH = 3
Public Const WS_EX_LAYERED = &H80000
Public Const GWL_EXSTYLE = -20
Public Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal Color As Long, ByVal x As Byte, ByVal alpha As Long) As Boolean
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long

Public Sub GradientFillRect( _
      ByVal lHDC As Long, _
      tR As RECT, _
      ByVal oStartColor As OLE_COLOR, _
      ByVal oEndColor As OLE_COLOR, _
      ByVal eDir As Long _
   )
Dim hBrush As Long
Dim lStartColor As Long
Dim lEndColor As Long
   
   ' Use GradientFill:
   lStartColor = TranslateColor(oStartColor)
   lEndColor = TranslateColor(oEndColor)

   Dim tTV(0 To 1) As TRIVERTEX
   Dim tGR As GRADIENT_RECT
   
   setTriVertexColor tTV(0), lStartColor
   tTV(0).x = tR.Left
   tTV(0).y = tR.Top
   setTriVertexColor tTV(1), lEndColor
   tTV(1).x = tR.Right
   tTV(1).y = tR.Bottom
   
   tGR.UpperLeft = 0
   tGR.LowerRight = 1
   
   GradientFill lHDC, tTV(0), 2, tGR, 1, eDir
      
   If (Err.Number <> 0) Then
      ' Fill with solid brush:
      hBrush = CreateSolidBrush(TranslateColor(oEndColor))
      FillRect lHDC, tR, hBrush
      DeleteObject hBrush
   End If
   
End Sub


Public Function TranslateColor( _
    ByVal oClr As OLE_COLOR, _
    Optional hPal As Long = 0 _
    ) As Long
    ' Convert Automation color to Windows color
    If OleTranslateColor(oClr, hPal, TranslateColor) Then
        TranslateColor = CLR_INVALID
    End If
End Function

Private Sub setTriVertexColor(tTV As TRIVERTEX, lColor As Long)
    Dim lRed As Long
    Dim lGreen As Long
    Dim lBlue As Long
   lRed = (lColor And &HFF&) * &H100&
   lGreen = (lColor And &HFF00&)
   lBlue = (lColor And &HFF0000) \ &H100&
   setTriVertexColorComponent tTV.Red, lRed
   setTriVertexColorComponent tTV.Green, lGreen
   setTriVertexColorComponent tTV.Blue, lBlue
End Sub

Private Sub setTriVertexColorComponent( _
   ByRef iColor As Integer, _
   ByVal lComponent As Long _
   )
   If (lComponent And &H8000&) = &H8000& Then
      iColor = (lComponent And &H7F00&)
      iColor = iColor Or &H8000
   Else
      iColor = lComponent
   End If
End Sub



