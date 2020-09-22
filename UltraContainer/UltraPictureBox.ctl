VERSION 5.00
Begin VB.UserControl UltraPictureBox 
   AutoRedraw      =   -1  'True
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ControlContainer=   -1  'True
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   Begin VB.PictureBox picMirror 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1785
      Left            =   150
      ScaleHeight     =   119
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   120
      TabIndex        =   3
      Top             =   30
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.PictureBox picGlass 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   2355
      Left            =   210
      ScaleHeight     =   2355
      ScaleWidth      =   2730
      TabIndex        =   2
      Top             =   960
      Visible         =   0   'False
      Width           =   2730
   End
   Begin VB.Timer tmrFade 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   3300
      Top             =   540
   End
   Begin VB.PictureBox picMySelf 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1785
      Left            =   2160
      ScaleHeight     =   119
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   165
      TabIndex        =   1
      Top             =   240
      Visible         =   0   'False
      Width           =   2475
   End
   Begin VB.PictureBox picContainer 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1785
      Left            =   2190
      ScaleHeight     =   119
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   165
      TabIndex        =   0
      Top             =   1530
      Visible         =   0   'False
      Width           =   2475
   End
End
Attribute VB_Name = "UltraPictureBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum GradientFillRectType
   GRADIENT_FILL_RECT_H = 0
   GRADIENT_FILL_RECT_V = 1
End Enum

Const PS_SOLID = 0

Dim oldRect As RECT
Dim currentDistance As Long
Dim minDistance As Long

Dim IsMouseIn As Boolean

Private Type POINTAPI
        x As Long
        y As Long
End Type


Const AC_SRC_OVER = &H0

Private Type BLENDFUNCTION
  BlendOp As Byte
  BlendFlags As Byte
  SourceConstantAlpha As Byte
  AlphaFormat As Byte
End Type

Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function AlphaBlend Lib "msimg32.dll" (ByVal hDC As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal hDC As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal BLENDFUNCT As Long) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32.dll" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function PrintWindow Lib "user32" (ByVal hwnd As Long, ByVal hdcBlt As Long, ByVal nFlags As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function RoundRect Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Event DblClick() 'MappingInfo=UserControl,UserControl,-1,DblClick
Attribute DblClick.VB_Description = "Occurs when the user presses and releases a mouse button and then presses and releases it again over an object."
Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyDown
Attribute KeyDown.VB_Description = "Occurs when the user presses a key while an object has the focus."
Event KeyPress(KeyAscii As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyPress
Attribute KeyPress.VB_Description = "Occurs when the user presses and releases an ANSI key."
Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyUp
Attribute KeyUp.VB_Description = "Occurs when the user releases a key while an object has the focus."
Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseDown
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus."
Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseMove
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseUp
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus."
Event OLECompleteDrag(Effect As Long) 'MappingInfo=UserControl,UserControl,-1,OLECompleteDrag
Attribute OLECompleteDrag.VB_Description = "Occurs at the OLE drag/drop source control after a manual or automatic drag/drop has been completed or canceled."
Event OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=UserControl,UserControl,-1,OLEDragDrop
Attribute OLEDragDrop.VB_Description = "Occurs when data is dropped onto the control via an OLE drag/drop operation, and OLEDropMode is set to manual."
Event OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer) 'MappingInfo=UserControl,UserControl,-1,OLEDragOver
Attribute OLEDragOver.VB_Description = "Occurs when the mouse is moved over the control during an OLE drag/drop operation, if its OLEDropMode property is set to manual."
Event OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean) 'MappingInfo=UserControl,UserControl,-1,OLEGiveFeedback
Attribute OLEGiveFeedback.VB_Description = "Occurs at the source control of an OLE drag/drop operation when the mouse cursor needs to be changed."
Event OLESetData(Data As DataObject, DataFormat As Integer) 'MappingInfo=UserControl,UserControl,-1,OLESetData
Attribute OLESetData.VB_Description = "Occurs at the OLE drag/drop source control when the drop target requests data that was not provided to the DataObject during the OLEDragStart event."
Event OLEStartDrag(Data As DataObject, AllowedEffects As Long) 'MappingInfo=UserControl,UserControl,-1,OLEStartDrag
Attribute OLEStartDrag.VB_Description = "Occurs when an OLE drag/drop operation is initiated either manually or automatically."
Event Paint() 'MappingInfo=UserControl,UserControl,-1,Paint
Attribute Paint.VB_Description = "Occurs when any part of a form or PictureBox control is moved, enlarged, or exposed."
Event Resize() 'MappingInfo=UserControl,UserControl,-1,Resize
Attribute Resize.VB_Description = "Occurs when a form is first displayed or the size of an object changes."
'Default Property Values:
Const m_def_FadeEnabled = 0
Const m_def_ShowMirror = 0
Const m_def_MirrorPercent = 40
Const m_def_TitleCaption = "Title"
Const m_def_TransparencyDistance = 120
Const m_def_TitleHeight = 24
Const m_def_ShowBorder = 0
Const m_def_BorderColor = 0
Const m_def_RoundShape = 0
Const m_def_ShowTitle = 0
Const m_def_TitleBackColorFrom = 0
Const m_def_TitleBackColorTo = 0
Const m_def_TitleFontName = "Arial"
Const m_def_TitleFontSize = 10
Const m_def_TitleFontBold = 0
Const m_def_TitleFontColor = 0
Const m_def_ShowBackgroundGradient = 0
Const m_def_BackgroundGradientFrom = 0
Const m_def_BackgroundGradientTo = 0
Const m_def_BackgroundGradientDirection = 0
'Property Variables:
Dim m_FadeEnabled As Boolean
Dim m_ShowMirror As Boolean
Dim m_MirrorPercent As Long
Dim m_TitleCaption As String
Dim m_TransparencyDistance As Long
Dim m_TitleHeight As Long
Dim m_ShowBorder As Boolean
Dim m_BorderColor As OLE_COLOR
Dim m_RoundShape As Boolean
Dim m_ShowTitle As Boolean
Dim m_TitleBackColorFrom As OLE_COLOR
Dim m_TitleBackColorTo As OLE_COLOR
Dim m_TitleFontName As String
Dim m_TitleFontSize As Integer
Dim m_TitleFontBold As Boolean
Dim m_TitleFontColor As OLE_COLOR
Dim m_ShowBackgroundGradient As Boolean
Dim m_BackgroundGradientFrom As OLE_COLOR
Dim m_BackgroundGradientTo As OLE_COLOR
Dim m_BackgroundGradientDirection As Integer

Private Sub DrawCustomProperties()
    Dim hRPen As Long
    Dim r As RECT
    Dim tmpFontName As String
    Dim tmpFontSize As Integer
    Dim tmpFontBold As Boolean
    Dim tmpFontColor As Long
    
    On Error Resume Next
    
    
    If m_RoundShape Then
        SetWindowRgn Me.hwnd, CreateRoundRectRgn(0, 0, _
                  ScaleX(ScaleWidth, ScaleMode, vbPixels), ScaleY(ScaleHeight, ScaleMode, vbPixels), _
                    9, 9), True
        
    Else
        SetWindowRgn Me.hwnd, CreateRoundRectRgn(0, 0, _
                  ScaleX(ScaleWidth, ScaleMode, vbPixels), ScaleY(ScaleHeight, ScaleMode, vbPixels), _
                    0, 0), True
        
    End If
      
    If m_ShowTitle Then
        r.Top = 0
        r.Left = 0
        r.Right = ScaleX(ScaleWidth, ScaleMode, vbPixels)
        r.Bottom = m_TitleHeight / 3 * 2
        
        GradientFillRect hDC, r, m_TitleBackColorFrom, m_TitleBackColorTo, GRADIENT_FILL_RECT_V
        
        r.Top = r.Bottom
        r.Bottom = m_TitleHeight
        
        GradientFillRect hDC, r, m_TitleBackColorTo, m_TitleBackColorFrom, GRADIENT_FILL_RECT_V
        
        tmpFontName = UserControl.FontName
        tmpFontSize = UserControl.FontSize
        tmpFontBold = UserControl.FontBold
        tmpFontColor = UserControl.ForeColor
        
        UserControl.FontName = m_TitleFontName
        UserControl.FontSize = m_TitleFontSize
        UserControl.FontBold = m_TitleFontBold
        UserControl.ForeColor = m_TitleFontColor
        
        r.Left = 4
        r.Top = (m_TitleHeight / 2) - (ScaleY(m_TitleFontSize, vbPoints, vbPixels) / 2) - 1
        r.Right = r.Left + ScaleX(UserControl.ScaleWidth, ScaleMode, vbPixels) - 6
        r.Bottom = m_TitleHeight
        
        
        DrawText hDC, m_TitleCaption, Len(m_TitleCaption), r, &H0
        
        UserControl.FontName = tmpFontName
        UserControl.FontSize = tmpFontSize
        UserControl.FontBold = tmpFontBold
        UserControl.ForeColor = tmpFontColor
    End If
    
    If m_ShowBackgroundGradient Then
        If m_ShowTitle Then
            r.Top = m_TitleHeight
        Else
            r.Top = 0
        End If
        
        r.Left = 0
        r.Right = ScaleX(UserControl.ScaleWidth, UserControl.ScaleMode, vbPixels)
        r.Bottom = ScaleY(UserControl.ScaleHeight, UserControl.ScaleMode, vbPixels)
        
        GradientFillRect hDC, r, m_BackgroundGradientFrom, m_BackgroundGradientTo, m_BackgroundGradientDirection
    End If
   
    If ShowBorder Then
        hRPen = CreatePen(PS_SOLID, 1, m_BorderColor)
        SelectObject hDC, hRPen
         
        If m_RoundShape Then
            RoundRect Me.hDC, 0, 0, ScaleX(ScaleWidth, ScaleMode, vbPixels) - 1, ScaleY(ScaleHeight, ScaleMode, vbPixels) - 1, 9, 9
        Else
            RoundRect Me.hDC, 0, 0, ScaleX(ScaleWidth, ScaleMode, vbPixels) - 1, ScaleY(ScaleHeight, ScaleMode, vbPixels) - 1, 0, 0
        End If
        
        DeleteObject hRPen
    End If
    
    Refresh
End Sub

Public Sub SetTransparency(mLevel As Long)
    On Error Resume Next
    
    Redraw mLevel
    
    Set Picture = picGlass.Image
    picGlass.Visible = False
End Sub

Private Sub Redraw(Optional mAlfa As Long)
    Dim objContainer As Object
    Dim x As Long
    Dim y As Long
    Dim BF As BLENDFUNCTION, lBF As Long
    Dim MyPos As RECT
    
    On Error Resume Next
    
       
    
    
    Set objContainer = Extender.Container
    
    picContainer.Width = ScaleWidth
    picContainer.Height = ScaleHeight
    
    x = ScaleX(Extender.Left, objContainer.ScaleMode, vbPixels)
    y = ScaleY(Extender.Top, objContainer.ScaleMode, vbPixels)
        
    BitBlt picContainer.hDC, 0, 0, ScaleX(ScaleWidth, ScaleMode, vbPixels), ScaleY(ScaleHeight, ScaleMode, vbPixels), objContainer.hDC, x, y, vbSrcCopy
    
    If mAlfa = 0 Then
        currentDistance = 0
        
    End If
    
    If mAlfa = 0 Then
        picGlass.Visible = False
        DrawCustomProperties
        picMySelf.Width = ScaleWidth
        picMySelf.Height = ScaleHeight
        
        PrintWindow hwnd, picMySelf.hDC, 0
        picMirror.Visible = False
        UpdateMirror
        Exit Sub
    End If
    
    GetWindowRect hwnd, MyPos
    
    If MyPos.Left <> oldRect.Left Or _
        MyPos.Top <> oldRect.Top Or _
        MyPos.Right <> oldRect.Right Or _
        MyPos.Bottom <> oldRect.Bottom Then
        
        oldRect = MyPos
        picGlass.Visible = False
    End If
    
    If Not picGlass.Visible Then
        DrawCustomProperties
    
        picMySelf.Width = ScaleWidth
        picMySelf.Height = ScaleHeight
        
        PrintWindow hwnd, picMySelf.hDC, 0
    End If
    
    
    With BF
        .BlendOp = AC_SRC_OVER
        .BlendFlags = 0
        .SourceConstantAlpha = mAlfa
        .AlphaFormat = 0
    End With

    RtlMoveMemory lBF, BF, 4
    
    picGlass.Picture = picMySelf.Image
    
    AlphaBlend picGlass.hDC, 0, 0, picMySelf.ScaleWidth, picMySelf.ScaleHeight, picContainer.hDC, 0, 0, picContainer.ScaleWidth, picContainer.ScaleHeight, lBF
    
    picGlass.Refresh
    picGlass.Visible = True
    picGlass.ZOrder 0
    
    UpdateMirror
    
    Refresh
    
    picGlass.Refresh
End Sub

Public Sub RedrawMirror()
    On Error Resume Next
    
    If Not IsMouseIn And m_FadeEnabled Then Exit Sub
    
    DrawCustomProperties
    
    picMySelf.Width = ScaleWidth
    picMySelf.Height = ScaleHeight
    
    DoEvents
    PrintWindow hwnd, picMySelf.hDC, 0
    
    UpdateMirror
    
    
End Sub

Private Sub UpdateMirror()
    Dim i As Long
    Dim BF As BLENDFUNCTION, lBF As Long
    Dim objContainer As Object
    Dim mAlpha As Long
    Dim y As Long
    
    On Error Resume Next
    
    picMirror.Visible = m_ShowMirror
    
    If Not m_ShowMirror Then Exit Sub
    
    picMirror.Move 0, CLng(ScaleHeight * ((100 - MirrorPercent) / 100)), ScaleWidth, ScaleHeight * (MirrorPercent / 100)
    picMirror.ZOrder 0
    
    Set objContainer = Extender.Container
        
    picMirror.Cls
    
    For i = 0 To picMySelf.ScaleHeight * (MirrorPercent / 100)
        With BF
            .BlendOp = AC_SRC_OVER
            .BlendFlags = 0
            
            mAlpha = (i / picMirror.ScaleHeight) * 255 + currentDistance
            
            If i > picMirror.ScaleHeight - 5 Then
                mAlpha = 255
            End If
            
            If mAlpha > 255 Then mAlpha = 255
            
            .SourceConstantAlpha = CByte(mAlpha)
            .AlphaFormat = 0
        End With
    
        RtlMoveMemory lBF, BF, 4
        
        BitBlt picMirror.hDC, 0, i, picContainer.ScaleWidth, 1, picMySelf.hDC, 0, CLng(picContainer.ScaleHeight * ((100 - MirrorPercent) / 100)) - i - 1, vbSrcCopy ' , picContainer.ScaleWidth, 1

        y = CLng(ScaleY(ScaleHeight, ScaleMode, vbPixels) * ((100 - MirrorPercent) / 100)) + i
        AlphaBlend picMirror.hDC, 0, i, picMirror.ScaleWidth, 1, picContainer.hDC, 0, y, picMirror.ScaleWidth, 1, lBF
    Next i
    
    picMirror.Refresh
End Sub

Private Sub picGlass_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Not IsMouseIn Then
        IsMouseIn = True
        Redraw 0
    End If
End Sub

Private Sub tmrFade_Timer()
    Dim MyPos As RECT
    Dim MousePos As POINTAPI
    Dim tmpPoint As POINTAPI
    Dim MiddlePoint As POINTAPI
    Dim tmpDistance As Long
    Dim WindowHwnd As Long
    Dim tmpHwnd As Long
    Dim ctl As Control
    
    On Error Resume Next
    
    tmrFade = False
    
    GetCursorPos MousePos
    
    If IsMouseIn Then
        WindowHwnd = WindowFromPoint(MousePos.x, MousePos.y)
        
        If WindowHwnd <> hwnd Then
            For Each ctl In ContainedControls
                tmpHwnd = 0
                tmpHwnd = ctl.hwnd
                
                If WindowHwnd = tmpHwnd Then
                    tmrFade = True
                    Exit Sub
                End If
                
            Next
            
            If WindowHwnd = picMirror.hwnd Then
                tmrFade = True
                Exit Sub
            End If
            
            IsMouseIn = False
            
        End If
        
    Else
        GetWindowRect hwnd, MyPos
        
        minDistance = 99999999
        
                
        
        minDistance = 99999999

        MiddlePoint = GetIntersectionPoint(APoint(MyPos.Left, MyPos.Top), APoint(MyPos.Right, MyPos.Bottom), APoint(MyPos.Left, MyPos.Bottom), APoint(MyPos.Right, MyPos.Top), MyPos)

        tmpPoint = GetIntersectionPoint(APoint(MyPos.Left, MyPos.Top), APoint(MyPos.Right, MyPos.Top), MiddlePoint, MousePos, MyPos)

        If tmpPoint.x <> 99999999 Then
            tmpDistance = PointDistance(MousePos, tmpPoint)

            If tmpDistance < minDistance Then
                minDistance = tmpDistance
            End If
        End If

        tmpPoint = GetIntersectionPoint(APoint(MyPos.Right, MyPos.Top), APoint(MyPos.Right, MyPos.Bottom), MiddlePoint, MousePos, MyPos)

        If tmpPoint.x <> 99999999 Then
            tmpDistance = PointDistance(MousePos, tmpPoint)
            If tmpDistance < minDistance Then
                minDistance = tmpDistance
            End If
        End If

        tmpPoint = GetIntersectionPoint(APoint(MyPos.Right, MyPos.Bottom), APoint(MyPos.Left, MyPos.Bottom), MiddlePoint, MousePos, MyPos)

        If tmpPoint.x <> 99999999 Then
            tmpDistance = PointDistance(MousePos, tmpPoint)

            If tmpDistance < minDistance Then
                minDistance = tmpDistance
            End If
        End If

        tmpPoint = GetIntersectionPoint(APoint(MyPos.Left, MyPos.Bottom), APoint(MyPos.Left, MyPos.Top), MiddlePoint, MousePos, MyPos)

        If tmpPoint.x <> 99999999 Then

            tmpDistance = PointDistance(MousePos, tmpPoint)

            If tmpDistance < minDistance Then
                minDistance = tmpDistance
            End If
        End If
        
        minDistance = Int(minDistance / 10) * 10 + 10
        
        If minDistance > (m_TransparencyDistance - 10) Then
            minDistance = m_TransparencyDistance
        End If
            
        If currentDistance <> minDistance Then
            currentDistance = currentDistance + 10 * (CLng(currentDistance > minDistance) + Abs(currentDistance < minDistance))
            Redraw currentDistance
        End If
                                
    End If
    
    tmrFade = True
End Sub

Private Function APoint(ByVal x As Long, ByVal y As Long) As POINTAPI
    APoint.x = x
    APoint.y = y
End Function

Private Function PointDistance(P1 As POINTAPI, P2 As POINTAPI) As Long
    On Error Resume Next
    
    PointDistance = Sqr((P1.x - P2.x) * (P1.x - P2.x) + (P1.y - P2.y) * (P1.y - P2.y))
End Function

Private Function GetIntersectionPoint(P1 As POINTAPI, _
                                        P2 As POINTAPI, _
                                        P3 As POINTAPI, _
                                        P4 As POINTAPI, _
                                        MyPos As RECT) As POINTAPI
    Dim u As Double
    Dim ret As POINTAPI
    
    On Error Resume Next
    
    u = ((P4.x - P3.x) * (P1.y - P3.y) - (P4.y - P3.y) * (P1.x - P3.x)) / ((P4.y - P3.y) * (P2.x - P1.x) - (P4.x - P3.x) * (P2.y - P1.y))
    
    ret.x = P1.x + u * (P2.x - P1.x)
    ret.y = P1.y + u * (P2.y - P1.y)
    
    If Not (ret.x >= MyPos.Left And ret.x <= MyPos.Right And _
        ret.y >= MyPos.Top And ret.y <= MyPos.Bottom) Then
        
        ret.x = 99999999
        ret.y = 99999999
    End If
    
    GetIntersectionPoint = ret
End Function

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    UserControl.BackStyle = PropBag.ReadProperty("BackStyle", 1)
    UserControl.ClipControls = PropBag.ReadProperty("ClipControls", True)
    UserControl.DrawMode = PropBag.ReadProperty("DrawMode", 13)
    UserControl.DrawStyle = PropBag.ReadProperty("DrawStyle", 0)
    UserControl.DrawWidth = PropBag.ReadProperty("DrawWidth", 1)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    Set UserControl.Font = PropBag.ReadProperty("Font", Ambient.Font)
    UserControl.FontBold = PropBag.ReadProperty("FontBold", 0)
    UserControl.FontItalic = PropBag.ReadProperty("FontItalic", 0)
    UserControl.FontName = PropBag.ReadProperty("FontName", "Arial")
    UserControl.FontSize = PropBag.ReadProperty("FontSize", 10)
    UserControl.FontStrikethru = PropBag.ReadProperty("FontStrikethru", 0)
    UserControl.FontTransparent = PropBag.ReadProperty("FontTransparent", True)
    UserControl.FontUnderline = PropBag.ReadProperty("FontUnderline", 0)
    UserControl.ForeColor = PropBag.ReadProperty("ForeColor", &H80000012)
    Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    UserControl.MousePointer = PropBag.ReadProperty("MousePointer", 0)
    Set Palette = PropBag.ReadProperty("Palette", Nothing)
    UserControl.PaletteMode = PropBag.ReadProperty("PaletteMode", 3)
    Set Picture = PropBag.ReadProperty("Picture", Nothing)
    UserControl.ScaleMode = PropBag.ReadProperty("ScaleMode", 1)
    UserControl.OLEDropMode = PropBag.ReadProperty("OLEDropMode", 0)
    m_ShowBorder = PropBag.ReadProperty("ShowBorder", m_def_ShowBorder)
    m_BorderColor = PropBag.ReadProperty("BorderColor", m_def_BorderColor)
    m_RoundShape = PropBag.ReadProperty("RoundShape", m_def_RoundShape)
    m_ShowTitle = PropBag.ReadProperty("ShowTitle", m_def_ShowTitle)
    m_TitleBackColorFrom = PropBag.ReadProperty("TitleBackColorFrom", m_def_TitleBackColorFrom)
    m_TitleBackColorTo = PropBag.ReadProperty("TitleBackColorTo", m_def_TitleBackColorTo)
    m_TitleFontName = PropBag.ReadProperty("TitleFontName", m_def_TitleFontName)
    m_TitleFontSize = PropBag.ReadProperty("TitleFontSize", m_def_TitleFontSize)
    m_TitleFontBold = PropBag.ReadProperty("TitleFontBold", m_def_TitleFontBold)
    m_TitleFontColor = PropBag.ReadProperty("TitleFontColor", m_def_TitleFontColor)
    m_ShowBackgroundGradient = PropBag.ReadProperty("ShowBackgroundGradient", m_def_ShowBackgroundGradient)
    m_BackgroundGradientFrom = PropBag.ReadProperty("BackgroundGradientFrom", m_def_BackgroundGradientFrom)
    m_BackgroundGradientTo = PropBag.ReadProperty("BackgroundGradientTo", m_def_BackgroundGradientTo)
    m_BackgroundGradientDirection = PropBag.ReadProperty("BackgroundGradientDirection", m_def_BackgroundGradientDirection)
    m_TitleHeight = PropBag.ReadProperty("TitleHeight", m_def_TitleHeight)
    m_TitleCaption = PropBag.ReadProperty("TitleCaption", m_def_TitleCaption)
    m_TransparencyDistance = PropBag.ReadProperty("TransparencyDistance", m_def_TransparencyDistance)
    m_ShowMirror = PropBag.ReadProperty("ShowMirror", m_def_ShowMirror)
    m_MirrorPercent = PropBag.ReadProperty("MirrorPercent", m_def_MirrorPercent)
    m_FadeEnabled = PropBag.ReadProperty("FadeEnabled", m_def_FadeEnabled)
    tmrFade = m_FadeEnabled And Ambient.UserMode
End Sub

Private Sub UserControl_Resize()
    RaiseEvent Resize
    On Error Resume Next
    
    picGlass.Move 0, 0, ScaleWidth, ScaleHeight
End Sub

Private Sub UserControl_Terminate()
    tmrFade = False
End Sub




'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackStyle
Public Property Get BackStyle() As Integer
Attribute BackStyle.VB_Description = "Indicates whether a Label or the background of a Shape is transparent or opaque."
    BackStyle = UserControl.BackStyle
End Property

Public Property Let BackStyle(ByVal New_BackStyle As Integer)
    UserControl.BackStyle() = New_BackStyle
    PropertyChanged "BackStyle"
End Property


'The Underscore following "Circle" is necessary because it
'is a Reserved Word in VBA.
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Circle
Public Sub Circle_(x As Single, y As Single, Radius As Single, Color As Long, StartPos As Single, EndPos As Single, Aspect As Single)
    UserControl.Circle (x, y), Radius, Color, StartPos, EndPos, Aspect
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,ClipControls
Public Property Get ClipControls() As Boolean
Attribute ClipControls.VB_Description = "Determines whether graphics methods in Paint events repaint an entire object or newly exposed areas."
    ClipControls = UserControl.ClipControls
End Property

Public Property Let ClipControls(ByVal New_ClipControls As Boolean)
    UserControl.ClipControls() = New_ClipControls
    PropertyChanged "ClipControls"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,DrawMode
Public Property Get DrawMode() As Integer
Attribute DrawMode.VB_Description = "Sets the appearance of output from graphics methods or of a Shape or Line control."
    DrawMode = UserControl.DrawMode
End Property

Public Property Let DrawMode(ByVal New_DrawMode As Integer)
    UserControl.DrawMode() = New_DrawMode
    PropertyChanged "DrawMode"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,DrawStyle
Public Property Get DrawStyle() As Integer
Attribute DrawStyle.VB_Description = "Determines the line style for output from graphics methods."
    DrawStyle = UserControl.DrawStyle
End Property

Public Property Let DrawStyle(ByVal New_DrawStyle As Integer)
    UserControl.DrawStyle() = New_DrawStyle
    PropertyChanged "DrawStyle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,DrawWidth
Public Property Get DrawWidth() As Integer
Attribute DrawWidth.VB_Description = "Returns/sets the line width for output from graphics methods."
    DrawWidth = UserControl.DrawWidth
End Property

Public Property Let DrawWidth(ByVal New_DrawWidth As Integer)
    UserControl.DrawWidth() = New_DrawWidth
    PropertyChanged "DrawWidth"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = UserControl.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set UserControl.Font = New_Font
    PropertyChanged "Font"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,FontBold
Public Property Get FontBold() As Boolean
Attribute FontBold.VB_Description = "Returns/sets bold font styles."
    FontBold = UserControl.FontBold
End Property

Public Property Let FontBold(ByVal New_FontBold As Boolean)
    UserControl.FontBold() = New_FontBold
    PropertyChanged "FontBold"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,FontItalic
Public Property Get FontItalic() As Boolean
Attribute FontItalic.VB_Description = "Returns/sets italic font styles."
    FontItalic = UserControl.FontItalic
End Property

Public Property Let FontItalic(ByVal New_FontItalic As Boolean)
    UserControl.FontItalic() = New_FontItalic
    PropertyChanged "FontItalic"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,FontName
Public Property Get FontName() As String
Attribute FontName.VB_Description = "Specifies the name of the font that appears in each row for the given level."
    FontName = UserControl.FontName
End Property

Public Property Let FontName(ByVal New_FontName As String)
    UserControl.FontName() = New_FontName
    PropertyChanged "FontName"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,FontSize
Public Property Get FontSize() As Single
Attribute FontSize.VB_Description = "Specifies the size (in points) of the font that appears in each row for the given level."
    FontSize = UserControl.FontSize
End Property

Public Property Let FontSize(ByVal New_FontSize As Single)
    UserControl.FontSize() = New_FontSize
    PropertyChanged "FontSize"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,FontStrikethru
Public Property Get FontStrikethru() As Boolean
Attribute FontStrikethru.VB_Description = "Returns/sets strikethrough font styles."
    FontStrikethru = UserControl.FontStrikethru
End Property

Public Property Let FontStrikethru(ByVal New_FontStrikethru As Boolean)
    UserControl.FontStrikethru() = New_FontStrikethru
    PropertyChanged "FontStrikethru"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,FontTransparent
Public Property Get FontTransparent() As Boolean
Attribute FontTransparent.VB_Description = "Returns/sets a value that determines whether background text/graphics on a Form, Printer or PictureBox are displayed."
    FontTransparent = UserControl.FontTransparent
End Property

Public Property Let FontTransparent(ByVal New_FontTransparent As Boolean)
    UserControl.FontTransparent() = New_FontTransparent
    PropertyChanged "FontTransparent"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,FontUnderline
Public Property Get FontUnderline() As Boolean
Attribute FontUnderline.VB_Description = "Returns/sets underline font styles."
    FontUnderline = UserControl.FontUnderline
End Property

Public Property Let FontUnderline(ByVal New_FontUnderline As Boolean)
    UserControl.FontUnderline() = New_FontUnderline
    PropertyChanged "FontUnderline"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ForeColor = UserControl.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    UserControl.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,hDC
Public Property Get hDC() As Long
Attribute hDC.VB_Description = "Returns a handle (from Microsoft Windows) to the object's device context."
    hDC = UserControl.hDC
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,HasDC
Public Property Get HasDC() As Boolean
Attribute HasDC.VB_Description = "Determines whether a unique display context is allocated for the control."
    HasDC = UserControl.HasDC
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,hWnd
Public Property Get hwnd() As Long
Attribute hwnd.VB_Description = "Returns a handle (from Microsoft Windows) to an object's window."
    
    On Error Resume Next
    
    hwnd = UserControl.hwnd
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Image
Public Property Get Image() As Picture
Attribute Image.VB_Description = "Returns a handle, provided by Microsoft Windows, to a persistent bitmap."
    Set Image = UserControl.Image
End Property

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Line
Public Sub Line(ByVal X1 As Single, ByVal Y1 As Single, ByVal X2 As Single, ByVal Y2 As Single, ByVal Color As Long)
Attribute Line.VB_Description = "Draws lines and rectangles on an object."
    UserControl.Line (X1, Y1)-(X2, Y2), Color
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseDown(Button, Shift, x, y)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,MouseIcon
Public Property Get MouseIcon() As Picture
Attribute MouseIcon.VB_Description = "Sets a custom mouse icon."
    Set MouseIcon = UserControl.MouseIcon
End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
    Set UserControl.MouseIcon = New_MouseIcon
    PropertyChanged "MouseIcon"
End Property

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseMove(Button, Shift, x, y)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,MousePointer
Public Property Get MousePointer() As Integer
Attribute MousePointer.VB_Description = "Returns/sets the type of mouse pointer displayed when over part of an object."
    MousePointer = UserControl.MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As Integer)
    UserControl.MousePointer() = New_MousePointer
    PropertyChanged "MousePointer"
End Property

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseUp(Button, Shift, x, y)
End Sub

Private Sub UserControl_OLECompleteDrag(Effect As Long)
    RaiseEvent OLECompleteDrag(Effect)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,OLEDrag
Public Sub OLEDrag()
Attribute OLEDrag.VB_Description = "Starts an OLE drag/drop event with the given control as the source."
    UserControl.OLEDrag
End Sub

Private Sub UserControl_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent OLEDragDrop(Data, Effect, Button, Shift, x, y)
End Sub

Private Sub UserControl_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
    RaiseEvent OLEDragOver(Data, Effect, Button, Shift, x, y, State)
End Sub

Private Sub UserControl_OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
    RaiseEvent OLEGiveFeedback(Effect, DefaultCursors)
End Sub

Private Sub UserControl_OLESetData(Data As DataObject, DataFormat As Integer)
    RaiseEvent OLESetData(Data, DataFormat)
End Sub

Private Sub UserControl_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
    RaiseEvent OLEStartDrag(Data, AllowedEffects)
End Sub

Private Sub UserControl_Paint()
    RaiseEvent Paint
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,PaintPicture
Public Sub PaintPicture(ByVal Picture As Picture, ByVal X1 As Single, ByVal Y1 As Single, ByVal Width1 As Variant, ByVal Height1 As Variant, ByVal X2 As Variant, ByVal Y2 As Variant, ByVal Width2 As Variant, ByVal Height2 As Variant, ByVal Opcode As Variant)
Attribute PaintPicture.VB_Description = "Draws the contents of a graphics file on a Form, PictureBox, or Printer object."
    UserControl.PaintPicture Picture, X1, Y1, Width1, Height1, X2, Y2, Width2, Height2, Opcode
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Palette
Public Property Get Palette() As Picture
Attribute Palette.VB_Description = "Returns/sets an image that contains the palette to use on an object when PaletteMode is set to Custom"
    Set Palette = UserControl.Palette
End Property

Public Property Set Palette(ByVal New_Palette As Picture)
    Set UserControl.Palette = New_Palette
    PropertyChanged "Palette"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,PaletteMode
Public Property Get PaletteMode() As Integer
Attribute PaletteMode.VB_Description = "Returns/sets a value that determines which palette to use for the controls on a object."
    PaletteMode = UserControl.PaletteMode
End Property

Public Property Let PaletteMode(ByVal New_PaletteMode As Integer)
    UserControl.PaletteMode() = New_PaletteMode
    PropertyChanged "PaletteMode"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Picture
Public Property Get Picture() As Picture
Attribute Picture.VB_Description = "Returns/sets a graphic to be displayed in a control."
    Set Picture = UserControl.Picture
End Property

Public Property Set Picture(ByVal New_Picture As Picture)
    Set UserControl.Picture = New_Picture
    PropertyChanged "Picture"
End Property

'The Underscore following "Point" is necessary because it
'is a Reserved Word in VBA.
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Point
Public Function Point(x As Single, y As Single) As Long
Attribute Point.VB_Description = "Returns, as an integer of type Long, the RGB color of the specified point on a Form or PictureBox object."
    Point = UserControl.Point(x, y)
End Function

'The Underscore following "PSet" is necessary because it
'is a Reserved Word in VBA.
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,PSet
Public Sub PSet_(x As Single, y As Single, Color As Long)
    UserControl.PSet Step(x, y), Color
End Sub

'The Underscore following "Scale" is necessary because it
'is a Reserved Word in VBA.
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Scale
Public Sub Scale_(Optional X1 As Variant, Optional Y1 As Variant, Optional X2 As Variant, Optional Y2 As Variant)
    UserControl.Scale (X1, Y1)-(X2, Y2)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,ScaleHeight
Public Property Get ScaleHeight() As Single
Attribute ScaleHeight.VB_Description = "Returns/sets the number of units for the vertical measurement of an object's interior."
    ScaleHeight = UserControl.ScaleHeight
End Property

Public Property Let ScaleHeight(ByVal New_ScaleHeight As Single)
    UserControl.ScaleHeight() = New_ScaleHeight
    PropertyChanged "ScaleHeight"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,ScaleLeft
Public Property Get ScaleLeft() As Single
Attribute ScaleLeft.VB_Description = "Returns/sets the horizontal coordinates for the left edges of an object."
    ScaleLeft = UserControl.ScaleLeft
End Property

Public Property Let ScaleLeft(ByVal New_ScaleLeft As Single)
    UserControl.ScaleLeft() = New_ScaleLeft
    PropertyChanged "ScaleLeft"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,ScaleMode
Public Property Get ScaleMode() As ScaleModeConstants
Attribute ScaleMode.VB_Description = "Returns/sets a value indicating measurement units for object coordinates when using graphics methods or positioning controls."
    ScaleMode = UserControl.ScaleMode
End Property

Public Property Let ScaleMode(ByVal New_ScaleMode As ScaleModeConstants)
    UserControl.ScaleMode() = New_ScaleMode
    PropertyChanged "ScaleMode"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,ScaleTop
Public Property Get ScaleTop() As Single
Attribute ScaleTop.VB_Description = "Returns/sets the vertical coordinates for the top edges of an object."
    ScaleTop = UserControl.ScaleTop
End Property

Public Property Let ScaleTop(ByVal New_ScaleTop As Single)
    UserControl.ScaleTop() = New_ScaleTop
    PropertyChanged "ScaleTop"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,ScaleWidth
Public Property Get ScaleWidth() As Single
Attribute ScaleWidth.VB_Description = "Returns/sets the number of units for the horizontal measurement of an object's interior."
    ScaleWidth = UserControl.ScaleWidth
End Property

Public Property Let ScaleWidth(ByVal New_ScaleWidth As Single)
    UserControl.ScaleWidth() = New_ScaleWidth
    PropertyChanged "ScaleWidth"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=5
Public Sub Update()
     Redraw 0
     
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,OLEDropMode
Public Property Get OLEDropMode() As Integer
Attribute OLEDropMode.VB_Description = "Returns/Sets whether this object can act as an OLE drop target."
    OLEDropMode = UserControl.OLEDropMode
End Property

Public Property Let OLEDropMode(ByVal New_OLEDropMode As Integer)
    UserControl.OLEDropMode() = New_OLEDropMode
    PropertyChanged "OLEDropMode"
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    Set UserControl.Font = Ambient.Font
    m_ShowBorder = m_def_ShowBorder
    m_BorderColor = m_def_BorderColor
    m_RoundShape = m_def_RoundShape
    m_ShowTitle = m_def_ShowTitle
    m_TitleBackColorFrom = m_def_TitleBackColorFrom
    m_TitleBackColorTo = m_def_TitleBackColorTo
    m_TitleFontName = m_def_TitleFontName
    m_TitleFontSize = m_def_TitleFontSize
    m_TitleFontBold = m_def_TitleFontBold
    m_TitleFontColor = m_def_TitleFontColor
    m_ShowBackgroundGradient = m_def_ShowBackgroundGradient
    m_BackgroundGradientFrom = m_def_BackgroundGradientFrom
    m_BackgroundGradientTo = m_def_BackgroundGradientTo
    m_BackgroundGradientDirection = m_def_BackgroundGradientDirection
    m_TitleHeight = m_def_TitleHeight
    m_TitleCaption = m_def_TitleCaption
    m_TransparencyDistance = m_def_TransparencyDistance
    m_ShowMirror = m_def_ShowMirror
    m_MirrorPercent = m_def_MirrorPercent
    m_FadeEnabled = m_def_FadeEnabled
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
    Call PropBag.WriteProperty("BackStyle", UserControl.BackStyle, 1)
    Call PropBag.WriteProperty("ClipControls", UserControl.ClipControls, True)
    Call PropBag.WriteProperty("DrawMode", UserControl.DrawMode, 13)
    Call PropBag.WriteProperty("DrawStyle", UserControl.DrawStyle, 0)
    Call PropBag.WriteProperty("DrawWidth", UserControl.DrawWidth, 1)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("Font", UserControl.Font, Ambient.Font)
    Call PropBag.WriteProperty("FontBold", UserControl.FontBold, 0)
    Call PropBag.WriteProperty("FontItalic", UserControl.FontItalic, 0)
    Call PropBag.WriteProperty("FontName", UserControl.FontName, "")
    Call PropBag.WriteProperty("FontSize", UserControl.FontSize, 0)
    Call PropBag.WriteProperty("FontStrikethru", UserControl.FontStrikethru, 0)
    Call PropBag.WriteProperty("FontTransparent", UserControl.FontTransparent, True)
    Call PropBag.WriteProperty("FontUnderline", UserControl.FontUnderline, 0)
    Call PropBag.WriteProperty("ForeColor", UserControl.ForeColor, &H80000012)
    Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
    Call PropBag.WriteProperty("MousePointer", UserControl.MousePointer, 0)
    Call PropBag.WriteProperty("Palette", Palette, Nothing)
    Call PropBag.WriteProperty("PaletteMode", UserControl.PaletteMode, 3)
    Call PropBag.WriteProperty("Picture", Picture, Nothing)
    Call PropBag.WriteProperty("ScaleMode", UserControl.ScaleMode, 1)
    Call PropBag.WriteProperty("OLEDropMode", UserControl.OLEDropMode, 0)
    Call PropBag.WriteProperty("ShowBorder", m_ShowBorder, m_def_ShowBorder)
    Call PropBag.WriteProperty("BorderColor", m_BorderColor, m_def_BorderColor)
    Call PropBag.WriteProperty("RoundShape", m_RoundShape, m_def_RoundShape)
    Call PropBag.WriteProperty("ShowTitle", m_ShowTitle, m_def_ShowTitle)
    Call PropBag.WriteProperty("TitleBackColorFrom", m_TitleBackColorFrom, m_def_TitleBackColorFrom)
    Call PropBag.WriteProperty("TitleBackColorTo", m_TitleBackColorTo, m_def_TitleBackColorTo)
    Call PropBag.WriteProperty("TitleFontName", m_TitleFontName, m_def_TitleFontName)
    Call PropBag.WriteProperty("TitleFontSize", m_TitleFontSize, m_def_TitleFontSize)
    Call PropBag.WriteProperty("TitleFontBold", m_TitleFontBold, m_def_TitleFontBold)
    Call PropBag.WriteProperty("TitleFontColor", m_TitleFontColor, m_def_TitleFontColor)
    Call PropBag.WriteProperty("ShowBackgroundGradient", m_ShowBackgroundGradient, m_def_ShowBackgroundGradient)
    Call PropBag.WriteProperty("BackgroundGradientFrom", m_BackgroundGradientFrom, m_def_BackgroundGradientFrom)
    Call PropBag.WriteProperty("BackgroundGradientTo", m_BackgroundGradientTo, m_def_BackgroundGradientTo)
    Call PropBag.WriteProperty("BackgroundGradientDirection", m_BackgroundGradientDirection, m_def_BackgroundGradientDirection)
    Call PropBag.WriteProperty("TitleHeight", m_TitleHeight, m_def_TitleHeight)
    Call PropBag.WriteProperty("TitleCaption", m_TitleCaption, m_def_TitleCaption)
    Call PropBag.WriteProperty("TransparencyDistance", m_TransparencyDistance, m_def_TransparencyDistance)
    Call PropBag.WriteProperty("ShowMirror", m_ShowMirror, m_def_ShowMirror)
    Call PropBag.WriteProperty("MirrorPercent", m_MirrorPercent, m_def_MirrorPercent)
    Call PropBag.WriteProperty("FadeEnabled", m_FadeEnabled, m_def_FadeEnabled)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get ShowBorder() As Boolean
    ShowBorder = m_ShowBorder
End Property

Public Property Let ShowBorder(ByVal New_ShowBorder As Boolean)
    m_ShowBorder = New_ShowBorder
    PropertyChanged "ShowBorder"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get BorderColor() As OLE_COLOR
    BorderColor = m_BorderColor
End Property

Public Property Let BorderColor(ByVal New_BorderColor As OLE_COLOR)
    m_BorderColor = New_BorderColor
    PropertyChanged "BorderColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get RoundShape() As Boolean
    RoundShape = m_RoundShape
End Property

Public Property Let RoundShape(ByVal New_RoundShape As Boolean)
    m_RoundShape = New_RoundShape
    PropertyChanged "RoundShape"
    
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get ShowTitle() As Boolean
    ShowTitle = m_ShowTitle
End Property

Public Property Let ShowTitle(ByVal New_ShowTitle As Boolean)
    m_ShowTitle = New_ShowTitle
    PropertyChanged "ShowTitle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get TitleBackColorFrom() As OLE_COLOR
    TitleBackColorFrom = m_TitleBackColorFrom
End Property

Public Property Let TitleBackColorFrom(ByVal New_TitleBackColorFrom As OLE_COLOR)
    m_TitleBackColorFrom = New_TitleBackColorFrom
    PropertyChanged "TitleBackColorFrom"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get TitleBackColorTo() As OLE_COLOR
    TitleBackColorTo = m_TitleBackColorTo
End Property

Public Property Let TitleBackColorTo(ByVal New_TitleBackColorTo As OLE_COLOR)
    m_TitleBackColorTo = New_TitleBackColorTo
    PropertyChanged "TitleBackColorTo"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,Arial
Public Property Get TitleFontName() As String
    TitleFontName = m_TitleFontName
End Property

Public Property Let TitleFontName(ByVal New_TitleFontName As String)
    m_TitleFontName = New_TitleFontName
    PropertyChanged "TitleFontName"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,10
Public Property Get TitleFontSize() As Integer
    TitleFontSize = m_TitleFontSize
End Property

Public Property Let TitleFontSize(ByVal New_TitleFontSize As Integer)
    m_TitleFontSize = New_TitleFontSize
    PropertyChanged "TitleFontSize"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get TitleFontBold() As Boolean
    TitleFontBold = m_TitleFontBold
End Property

Public Property Let TitleFontBold(ByVal New_TitleFontBold As Boolean)
    m_TitleFontBold = New_TitleFontBold
    PropertyChanged "TitleFontBold"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get TitleFontColor() As OLE_COLOR
    TitleFontColor = m_TitleFontColor
End Property

Public Property Let TitleFontColor(ByVal New_TitleFontColor As OLE_COLOR)
    m_TitleFontColor = New_TitleFontColor
    PropertyChanged "TitleFontColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get ShowBackgroundGradient() As Boolean
    ShowBackgroundGradient = m_ShowBackgroundGradient
End Property

Public Property Let ShowBackgroundGradient(ByVal New_ShowBackgroundGradient As Boolean)
    m_ShowBackgroundGradient = New_ShowBackgroundGradient
    PropertyChanged "ShowBackgroundGradient"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get BackgroundGradientFrom() As OLE_COLOR
    BackgroundGradientFrom = m_BackgroundGradientFrom
End Property

Public Property Let BackgroundGradientFrom(ByVal New_BackgroundGradientFrom As OLE_COLOR)
    m_BackgroundGradientFrom = New_BackgroundGradientFrom
    PropertyChanged "BackgroundGradientFrom"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get BackgroundGradientTo() As OLE_COLOR
    BackgroundGradientTo = m_BackgroundGradientTo
End Property

Public Property Let BackgroundGradientTo(ByVal New_BackgroundGradientTo As OLE_COLOR)
    m_BackgroundGradientTo = New_BackgroundGradientTo
    PropertyChanged "BackgroundGradientTo"
End Property



'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get BackgroundGradientDirection() As Integer
    BackgroundGradientDirection = m_BackgroundGradientDirection
End Property

Public Property Let BackgroundGradientDirection(ByVal New_BackgroundGradientDirection As Integer)
    m_BackgroundGradientDirection = New_BackgroundGradientDirection
    PropertyChanged "BackgroundGradientDirection"
    
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,18
Public Property Get TitleHeight() As Long
    TitleHeight = m_TitleHeight
End Property

Public Property Let TitleHeight(ByVal New_TitleHeight As Long)
    m_TitleHeight = New_TitleHeight
    PropertyChanged "TitleHeight"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,0
Public Property Get TitleCaption() As String
    TitleCaption = m_TitleCaption
End Property

Public Property Let TitleCaption(ByVal New_TitleCaption As String)
    m_TitleCaption = New_TitleCaption
    PropertyChanged "TitleCaption"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get TransparencyDistance() As Long
    TransparencyDistance = m_TransparencyDistance
End Property

Public Property Let TransparencyDistance(ByVal New_TransparencyDistance As Long)
    m_TransparencyDistance = New_TransparencyDistance
    PropertyChanged "TransparencyDistance"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get ShowMirror() As Boolean
    ShowMirror = m_ShowMirror
End Property

Public Property Let ShowMirror(ByVal New_ShowMirror As Boolean)
    m_ShowMirror = New_ShowMirror
    PropertyChanged "ShowMirror"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,40
Public Property Get MirrorPercent() As Long
    MirrorPercent = m_MirrorPercent
End Property

Public Property Let MirrorPercent(ByVal New_MirrorPercent As Long)
    m_MirrorPercent = New_MirrorPercent
    PropertyChanged "MirrorPercent"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get FadeEnabled() As Boolean
    FadeEnabled = m_FadeEnabled
End Property

Public Property Let FadeEnabled(ByVal New_FadeEnabled As Boolean)
    m_FadeEnabled = New_FadeEnabled
    tmrFade = m_FadeEnabled And Ambient.UserMode
    PropertyChanged "FadeEnabled"
End Property

