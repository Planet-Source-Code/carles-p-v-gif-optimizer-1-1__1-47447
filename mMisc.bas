Attribute VB_Name = "mMisc"
Option Explicit

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Type RECT2
    x1 As Long
    y1 As Long
    x2 As Long
    y2 As Long
End Type

Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function CreatePatternBrush Lib "gdi32" (ByVal hBitmap As Long) As Long
Private Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Integer) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT2, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As RECT2, ByVal hBrush As Long) As Long
Private Declare Function SetBrushOrgEx Lib "gdi32" (ByVal hDC As Long, ByVal nXOrg As Long, ByVal nYOrg As Long, lppt As POINTAPI) As Long
Private Declare Function DrawFocusRect Lib "user32" (ByVal hDC As Long, lpRect As RECT2) As Long

Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Const SWP_SHOWWINDOW    As Long = &H40
Private Const SWP_FRAMECHANGED  As Long = &H20
Private Const SWP_NOACTIVATE    As Long = &H10
Private Const SWP_NOMOVE        As Long = &H2
Private Const SWP_NOOWNERZORDER As Long = &H200
Private Const SWP_NOSIZE        As Long = &H1
Private Const SWP_NOZORDER      As Long = &H4

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Const GWL_STYLE         As Long = (-16)
Private Const WS_THICKFRAME     As Long = &H40000
Private Const WS_BORDER         As Long = &H800000
Private Const GWL_EXSTYLE       As Long = (-20)
Private Const WS_EX_WINDOWEDGE  As Long = &H100&
Private Const WS_EX_CLIENTEDGE  As Long = &H200&
Private Const WS_EX_STATICEDGE  As Long = &H20000

'//

Private m_hBrush As Long

'//

Public Sub InitializePatternBrush()

  Dim hBitmap        As Long
  Dim tBytes(1 To 8) As Integer
    
    '-- Brush pattern (8x8)
    tBytes(1) = &H55
    tBytes(2) = &HAA
    tBytes(3) = &H55
    tBytes(4) = &HAA
    tBytes(5) = &H55
    tBytes(6) = &HAA
    tBytes(7) = &H55
    tBytes(8) = &HAA
    
    '-- Create brush
    hBitmap = CreateBitmap(8, 8, 1, 1, tBytes(1))
    m_hBrush = CreatePatternBrush(hBitmap)
    DeleteObject hBitmap
End Sub

Public Sub DestroyPatternBrush()
    DeleteObject m_hBrush
End Sub

'//

Public Sub DrawRectangle(ByVal hDC As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long, ByVal Color As Long)

  Dim pPoint As POINTAPI
  Dim rRect  As RECT2
  Dim hBrush As Long
  
    If (Color > -1) Then
        '-- Solid color
        hBrush = CreateSolidBrush(Color)
        SetRect rRect, x1, y1, x2, y2
        FillRect hDC, rRect, hBrush
        DeleteObject hBrush
      Else
        '-- Pattern brush
        SetBrushOrgEx hDC, 0, 0, pPoint
        SetRect rRect, x1, y1, x2, y2
        FillRect hDC, rRect, m_hBrush
    End If
End Sub

Public Sub DrawFocus(ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal Width As Long, ByVal Height As Long)

  Dim lpFrameRect As RECT2
    
    '-- Draw simple dot-rectangle
    With lpFrameRect
        .x1 = x
        .y1 = y
        .x2 = x + Width
        .y2 = y + Height
    End With
    DrawFocusRect hDC, lpFrameRect
End Sub

Public Sub RemoveBorder(ByVal lhWnd As Long)
    pvSetWinStyle lhWnd, GWL_STYLE, 0, WS_BORDER Or WS_THICKFRAME
    pvSetWinStyle lhWnd, GWL_EXSTYLE, 0, WS_EX_STATICEDGE Or WS_EX_CLIENTEDGE Or WS_EX_WINDOWEDGE
End Sub

'//

Private Sub pvSetWinStyle(ByVal lhWnd As Long, ByVal lType As Long, ByVal lStyle As Long, ByVal lStyleNot As Long)

  Dim lS As Long
    
    lS = GetWindowLong(lhWnd, lType)
    lS = (lS And Not lStyleNot) Or lStyle
    SetWindowLong lhWnd, lType, lS
    SetWindowPos lhWnd, 0, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOOWNERZORDER Or SWP_NOZORDER Or SWP_FRAMECHANGED
End Sub
