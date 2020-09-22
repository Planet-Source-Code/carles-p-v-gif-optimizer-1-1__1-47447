VERSION 5.00
Begin VB.UserControl ucFramesList 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1530
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1200
   FillStyle       =   4  'Upward Diagonal
   KeyPreview      =   -1  'True
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   102
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   80
   Begin VB.TextBox txtDummy 
      Height          =   405
      Left            =   -195
      TabIndex        =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.VScrollBar ucBar 
      Height          =   1110
      Left            =   900
      Max             =   0
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   75
      Visible         =   0   'False
      Width           =   225
   End
End
Attribute VB_Name = "ucFramesList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'================================================
' User control:  ucFramesList.ctl (variation)
' Author:        Carles P.V.
' Dependencies:  cGIF.cls
'                cTile.cls
' Last revision: 2003.07.30
'================================================

Option Explicit

'-- API:

Private Type RECT2
    x1 As Long
    y1 As Long
    x2 As Long
    y2 As Long
End Type

Private Const SM_CXVSCROLL     As Long = &H2
Private Const DT_RIGHT         As Long = &H2
Private Const DT_WORDBREAK     As Long = &H10
Private Const DFC_BUTTON       As Long = &H4
Private Const DFCS_BUTTONCHECK As Long = &H0
Private Const DFCS_FLAT        As Long = &H4000
Private Const DFCS_CHECKED     As Long = &H400
Private Const PS_SOLID         As Long = &H0
Private Const BS_NULL          As Long = &H1

Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function OleTranslateColor Lib "olepro32" (ByVal OLE_COLOR As Long, ByVal hPalette As Long, ColorRef As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT2, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function PtInRect Lib "user32" (lpRect As RECT2, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As RECT2, ByVal hBrush As Long) As Long
Private Declare Function Rectangle Lib "gdi32" (ByVal hDC As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT2, ByVal wFormat As Long) As Long

'//

'-- Private Types:
Private Type tItem
    Text     As String
    ItemData As Long
End Type

'-- Default Property Values:
Private Const mdef_ThumbnailHeight As Integer = 75
Private Const mdef_ThumbnailPad    As Integer = 5

'-- Property Variables:
Private m_List()                  As tItem     ' List array of items
Private m_ThumbnailHeight         As Integer   ' Frame: Max. thumbnail height
Private m_ListIndex               As Integer   ' Current list index

'-- Private Variables:
Private m_lpoGIF                  As cGIF      ' Pointer to cGIF object
Private m_BackPattern             As New cTile ' Thumbnail background pattern
Private m_LastIndex               As Boolean   ' Last selected item
Private m_MouseDown               As Boolean   ' Mouse down flag
Private m_LastBar                 As Integer   ' Last scroll bar value
Private m_VisibleRows             As Integer   ' Visible rows
Private m_PerfectRowPad           As Boolean   ' Visible rows padding
Private m_ControlRect             As RECT2     ' User control rectangle (clearing)
Private m_RectExt()               As RECT2     ' Item rectangle
Private m_RectTxt()               As RECT2     ' Item text rectangle
Private m_RectFra()               As RECT2     ' Item frame rectangle
Private m_RectRowPad              As RECT2     ' Last row padding rectangle (background erasing)
Private m_ClrWindow               As Long      ' Back color [Normal]
Private m_ClrHighlight            As Long      ' Back color [Selected]
Private m_ClrWindowText           As Long      ' Font color [Normal]
Private m_ClrHighlightText        As Long      ' Font color [Selected]
Private m_ClrApplicationWorkspace As Long      ' Usercontrol background

'-- Event Declarations:
Public Event Click()
Public Event KeyDown(ByVal KeyCode As Integer)
Public Event MouseDown(ByVal Button As Integer, ByVal Item As Integer)
Public Event OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,OLEDragDrop



'========================================================================================
' UserControl
'========================================================================================

Private Sub UserControl_Initialize()

    '-- Initialize items array
    ReDim m_List(0)
    
    '-- Initialize position flags
    m_LastIndex = -1
    m_LastBar = -1
    
    '-- Set system default scroll bar width
    ucBar.Width = GetSystemMetrics(SM_CXVSCROLL)
    
    '-- Set background pattern
    m_BackPattern.SetPatternFromStdPicture LoadResPicture("PATTERN_08", vbResBitmap)
End Sub

Private Sub UserControl_Terminate()

    '-- Destroy GIF reference
    If (Not m_lpoGIF Is Nothing) Then
        Set m_lpoGIF = Nothing
    End If
    
    '-- Clear items array
    Erase m_List()
    
    '-- Destroy background pattern
    m_BackPattern.DestroyPattern
End Sub

'//

Private Sub UserControl_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent OLEDragDrop(Data, Effect, Button, Shift, X, Y)
End Sub

'//
    
Private Sub UserControl_Show()
    '-- Refresh control
    Refresh
End Sub

Private Sub UserControl_Resize()
    
  Dim sVisibleRows As Single
    
    '-- Check minimum height (one row)
    If (ScaleHeight < m_ThumbnailHeight + mdef_ThumbnailPad) Then
        Height = ((m_ThumbnailHeight + mdef_ThumbnailPad) + (Height \ Screen.TwipsPerPixelY - ScaleHeight)) * Screen.TwipsPerPixelY
    End If
    
    '-- Get visible rows
    sVisibleRows = ScaleHeight / (m_ThumbnailHeight + mdef_ThumbnailPad)
    
    '-- Perfect row adjustment [?]
    If (sVisibleRows = Int(sVisibleRows)) Then
        m_VisibleRows = sVisibleRows
        m_PerfectRowPad = -1
      Else
        m_VisibleRows = Int(sVisibleRows) + 1
        m_PerfectRowPad = 0
    End If

    '-- Calc. items rects.
    pvCalculateRects
    
    '-- Scroll bar
    ucBar.Visible = 0
    ucBar.Move ScaleWidth - ucBar.Width, 0, ucBar.Width, ScaleHeight
    pvReadjustBar
    
    '-- Refresh (erase background and refresh whole list)
    pvRectangle hDC, m_ControlRect, m_ClrApplicationWorkspace, m_ClrApplicationWorkspace
    Refresh
End Sub

'//

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Select Case KeyCode

      Case vbKeyUp       ' Row up
        If (m_ListIndex > 0) Then
            ListIndex = m_ListIndex - 1
        End If
        
      Case vbKeyDown     ' Row down
        If (m_ListIndex < UBound(m_List) - 1) Then
            ListIndex = m_ListIndex + 1
        End If

      Case vbKeyPageUp   ' Page up
        If (m_ListIndex > m_VisibleRows) Then
            ListIndex = m_ListIndex - m_VisibleRows - (Not m_PerfectRowPad)
          Else
            ListIndex = 0
        End If
       
      Case vbKeyPageDown ' Page down
        If (m_ListIndex < UBound(m_List) - m_VisibleRows - 1) Then
            ListIndex = m_ListIndex + m_VisibleRows + (Not m_PerfectRowPad)
          Else
            ListIndex = UBound(m_List) - 1
        End If

      Case vbKeyHome     ' Start
        ListIndex = 0

      Case vbKeyEnd      ' End
        ListIndex = UBound(m_List) - 1
    End Select
    
    RaiseEvent KeyDown(KeyCode)
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

  Dim nItm As Integer
  
    If (UBound(m_List) = 0) Then Exit Sub
    
    m_MouseDown = -1
    nItm = ucBar + (Y \ (m_ThumbnailHeight + mdef_ThumbnailPad))
    
    If (nItm > -1 And nItm < UBound(m_List)) Then
        '-- Select item
        If (Button = vbLeftButton) Then
            ListIndex = nItm
        End If
    End If
    
    RaiseEvent MouseDown(Button, nItm)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  
    If (UBound(m_List) = 0) Then Exit Sub
  
    If (m_MouseDown And Button = vbLeftButton) Then
        '-- Item selected
        RaiseEvent Click
    End If
    
    m_MouseDown = 0
End Sub

'========================================================================================
' Scroll bar
'========================================================================================

Private Sub ucBar_Change()
    If (m_LastBar <> ucBar) Then
        m_LastBar = ucBar
        Refresh
    End If
End Sub

Private Sub ucBar_Scroll()
    ucBar_Change
End Sub

'========================================================================================
' Methods
'========================================================================================

Public Sub SetGIFSource(ByVal oGIF As cGIF)
    '-- Set reference to source GIF object
    Set m_lpoGIF = oGIF
End Sub

Public Sub AddItem(ByVal Text, Optional ByVal ItemData As Long = 0)

    With m_List(UBound(m_List))
        .Text = CStr(Text)
        .ItemData = ItemData
    End With
    ReDim Preserve m_List(UBound(m_List) + 1)
    
    pvReadjustBar
End Sub

Public Sub Clear()

    '-- Reset/Hide scroll bar
    m_LastBar = 0: ucBar = 0
    m_LastBar = -1
    ucBar.Visible = 0
    '-- Clean control
    pvRectangle hDC, m_ControlRect, m_ClrApplicationWorkspace, m_ClrApplicationWorkspace
    '-- Reset items array / flags
    ReDim m_List(0)
    m_LastIndex = -1
    m_ListIndex = -1
End Sub

Public Sub Refresh()

    '-- Paint whole list
    If (Ambient.UserMode And Extender.Visible) Then
        pvDrawList
        UserControl.Refresh
    End If
End Sub

Public Sub RefreshItem(ByVal nIndex As Integer)

    '-- Paint single item
    If (Ambient.UserMode And Extender.Visible) Then
        If ((nIndex >= ucBar And nIndex <= ucBar + m_VisibleRows) And nIndex < UBound(m_List)) Then
            pvDrawItem nIndex
            UserControl.Refresh
        End If
    End If
End Sub

'========================================================================================
' Properties
'========================================================================================

Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property
Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    ucBar.Enabled = New_Enabled
End Property

Public Property Get ListCount() As Integer
Attribute ListCount.VB_MemberFlags = "400"
    ListCount = UBound(m_List)
End Property

Public Property Get ListIndex() As Integer
Attribute ListIndex.VB_MemberFlags = "400"
    ListIndex = m_ListIndex
End Property
Public Property Let ListIndex(ByVal New_ListIndex As Integer)
    
    '-- Check/Set
    If (New_ListIndex < 0 Or UBound(m_List) = 0) Then
        m_ListIndex = -1
      Else
        m_ListIndex = New_ListIndex
    End If
    m_LastIndex = m_ListIndex

    '-- Ensure visible current selected item
    If (m_ListIndex < ucBar And m_ListIndex > -1) Then
        ucBar = m_ListIndex
      ElseIf (m_ListIndex > ucBar + m_VisibleRows - 1 + (Not m_PerfectRowPad)) Then
        ucBar = m_ListIndex - m_VisibleRows + 1 - (Not m_PerfectRowPad)
      Else
        Refresh
    End If
    
    '-- Raise <Click> event [?]
    If (Not m_MouseDown) Then RaiseEvent Click
End Property

Public Property Get TopIndex() As Integer
Attribute TopIndex.VB_MemberFlags = "400"
    TopIndex = ucBar
End Property
Public Property Let TopIndex(ByVal New_TopIndex As Integer)

    '-- Check
    If (New_TopIndex > ucBar.Max) Then
        New_TopIndex = ucBar.Max
    End If
    If (New_TopIndex < 0) Then
        New_TopIndex = 0
    End If
    '-- Set and refresh
    m_LastBar = New_TopIndex
    ucBar = New_TopIndex
    Refresh
End Property

Public Property Get ThumbnailHeight() As Integer
    ThumbnailHeight = m_ThumbnailHeight
End Property

Public Property Let ThumbnailHeight(ByVal New_ThumbnailHeight As Integer)
    m_ThumbnailHeight = New_ThumbnailHeight
    '-- Clean control
    pvRectangle hDC, m_ControlRect, m_ClrApplicationWorkspace, m_ClrApplicationWorkspace
    '-- Refresh
    UserControl_Resize
End Property

'-- Item data...

Public Property Get ItemText(ByVal nIndex As Integer) As Variant
    ItemText = m_List(nIndex).Text
End Property
Public Property Let ItemText(ByVal nIndex As Integer, ByVal New_Text)
    m_List(nIndex).Text = CStr(New_Text)
End Property

Public Property Get ItemData(ByVal nIndex As Integer) As Long
    ItemData = m_List(nIndex).ItemData
End Property
Public Property Let ItemData(ByVal nIndex As Integer, ByVal New_ItemData As Long)
    m_List(nIndex).ItemData = New_ItemData
End Property

'//

Private Sub UserControl_InitProperties()
    '-- Default thumbnail height
    m_ThumbnailHeight = mdef_ThumbnailHeight
    '-- Set colors
    pvSetColors
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    '-- Set colors
    pvSetColors
    '-- Set font
    UserControl.Font = Ambient.Font
    '-- Read props.
    With PropBag
        UserControl.Enabled = .ReadProperty("Enabled", -1)
        ThumbnailHeight = .ReadProperty("ThumbnailHeight", mdef_ThumbnailHeight)
    End With
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        .WriteProperty "Enabled", UserControl.Enabled, -1
        .WriteProperty "ThumbnailHeight", m_ThumbnailHeight, mdef_ThumbnailHeight
    End With
End Sub

'========================================================================================
' Private
'========================================================================================

Private Sub pvDrawList()

  Dim nItm As Integer
    
    '-- Draw visible rows
    For nItm = ucBar To ucBar + m_VisibleRows - 1
        pvDrawItem nItm
    Next nItm
    
    '-- Clear pad row
    If (Not m_PerfectRowPad And ucBar = ucBar.Max) Then
        pvRectangle hDC, m_RectRowPad, m_ClrApplicationWorkspace, m_ClrApplicationWorkspace
    End If
End Sub

Private Sub pvDrawItem(ByVal nItm As Integer)
  
  Dim nRct As Integer
  Dim sF   As Single
  Dim xOff As Long
  Dim yOff As Long
    
    '-- Visible item [?]
    If ((nItm >= ucBar And nItm < ucBar + m_VisibleRows - (Not m_PerfectRowPad)) And nItm < UBound(m_List)) Then
        
        '-- Rect. number
        nRct = nItm - ucBar
        
        '-- Draw background
        If (m_ListIndex = nItm) Then
            pvRectangle hDC, m_RectExt(nRct), m_ClrHighlight, m_ClrApplicationWorkspace
            pvRectangle hDC, m_RectTxt(nRct), m_ClrWindow, m_ClrApplicationWorkspace
          Else
            pvRectangle hDC, m_RectExt(nRct), m_ClrWindow, m_ClrApplicationWorkspace
            pvRectangle hDC, m_RectTxt(nRct), m_ClrWindow, m_ClrApplicationWorkspace
        End If
        
        '-- Draw thumbnail
        If (Not m_lpoGIF Is Nothing) Then
            '-- Get fit factor and thumbnail offsets
            sF = pvGetBestFitFactor
            xOff = ((m_RectFra(nRct).x2 - m_RectFra(nRct).x1) - sF * m_lpoGIF.ScreenWidth) \ 2
            yOff = ((m_RectFra(nRct).y2 - m_RectFra(nRct).y1) - sF * m_lpoGIF.ScreenHeight) \ 2
            '-- Background pattern
            m_BackPattern.Tile hDC, xOff + m_RectFra(nRct).x1, yOff + m_RectFra(nRct).y1, sF * m_lpoGIF.ScreenWidth, sF * m_lpoGIF.ScreenHeight
            '-- Frame
            m_lpoGIF.FrameDraw hDC, nItm + 1, xOff + m_RectFra(nRct).x1, yOff + m_RectFra(nRct).y1, sF, -1
        End If
        
        '-- Draw text
        DrawText hDC, m_List(nItm).Text, Len(m_List(nItm).Text), m_RectTxt(nRct), DT_RIGHT Or DT_WORDBREAK
    End If
End Sub

Private Function pvGetBestFitFactor() As Single
    
  Dim sW As Single
  Dim sH As Single
   
    With m_RectFra(0)
        sW = (.x2 - .x1) / m_lpoGIF.ScreenWidth
        sH = (.y2 - .y1) / m_lpoGIF.ScreenHeight
    End With
    
    If (sW < sH) Then
        pvGetBestFitFactor = sW
      Else
        pvGetBestFitFactor = sH
    End If
    If (pvGetBestFitFactor > 1) Then
        pvGetBestFitFactor = 1
    End If
End Function

Private Sub pvRectangle(ByVal hDC As Long, m_Rect As RECT2, ByVal FillColor As Long, ByVal BorderColor As Long)

  Dim hBrush    As Long
  Dim hOldBrush As Long
  Dim hPen      As Long
  Dim hOldPen   As Long
    
    '-- Create Pen / Brush
    hPen = CreatePen(PS_SOLID, 1, BorderColor)
    hBrush = CreateSolidBrush(FillColor)
    
    '-- Select into given DC
    hOldPen = SelectObject(hDC, hPen)
    hOldBrush = SelectObject(hDC, hBrush)
    
    '-- Draw rectangle
    Rectangle hDC, m_Rect.x1, m_Rect.y1, m_Rect.x2, m_Rect.y2
    
    '-- Destroy used objects
    SelectObject hDC, hOldBrush
    DeleteObject hBrush
    SelectObject hDC, hOldPen
    DeleteObject hPen
End Sub

Private Sub pvReadjustBar()

    On Error Resume Next
    
    If (UBound(m_List) > m_VisibleRows + (Not m_PerfectRowPad)) Then
        If (Not ucBar.Visible) Then
            '-- Show scroll bar
            ucBar.Visible = -1
            ucBar.Refresh
            pvUpdateRectRight (ScaleWidth - ucBar.Width)
        End If
      Else
        '-- Hide scroll bar
        ucBar.Visible = 0
        ucBar.Refresh
        pvUpdateRectRight (ScaleWidth)
    End If

    '-- Update max value
    ucBar.LargeChange = m_VisibleRows
    ucBar.Max = (UBound(m_List) - m_VisibleRows) + -(Not m_PerfectRowPad)

    On Error GoTo 0
End Sub

Private Sub pvCalculateRects()
  
  Dim nRows    As Integer
  Dim nRct     As Integer
  Dim nItmH    As Integer
  Dim nTxtH    As Integer
  Dim nBarLeft As Integer
  
    nItmH = m_ThumbnailHeight + mdef_ThumbnailPad
    nTxtH = TextHeight("")
    nBarLeft = IIf(ucBar.Visible, ucBar.Left, ScaleWidth)
    
    '-- Main rect.
    SetRect m_ControlRect, 0, 0, ScaleWidth, ScaleHeight
    
    '-- Item rects.
    nRows = m_VisibleRows - 1
    ReDim m_RectExt(nRows)
    ReDim m_RectTxt(nRows)
    ReDim m_RectFra(nRows)
    
    For nRct = 0 To nRows
        SetRect m_RectTxt(nRct), 0, nRct * nItmH, 31, nRct * nItmH + nItmH + 1
        SetRect m_RectExt(nRct), 30, nRct * nItmH, nBarLeft, nRct * nItmH + nItmH + 1
        SetRect m_RectFra(nRct), 33, nRct * nItmH + 3, nBarLeft - 3, nRct * nItmH + nItmH - 2
    Next nRct
    
    '-- Pad rect.
    If (Not m_PerfectRowPad) Then
        With m_RectExt(nRows)
            SetRect m_RectRowPad, 0, .y1, .x2, ScaleHeight
        End With
      Else
        SetRect m_RectRowPad, 0, 0, 0, 0
    End If
End Sub

Private Sub pvUpdateRectRight(ByVal New_Right As Integer)

  Dim nRct As Integer
    
    '-- Rects. right offset
    For nRct = 0 To m_VisibleRows - 1
        m_RectExt(nRct).x2 = New_Right
        m_RectFra(nRct).x2 = New_Right - 3
    Next nRct
    m_RectRowPad.x2 = New_Right
End Sub

Private Sub pvSetColors()

    '-- Get long colors
    OleTranslateColor vbWindowBackground, 0, m_ClrWindow
    OleTranslateColor vbHighlight, 0, m_ClrHighlight
    OleTranslateColor vbWindowText, 0, m_ClrWindowText
    OleTranslateColor vbHighlightText, 0, m_ClrHighlightText
    OleTranslateColor vbApplicationWorkspace, 0, m_ClrApplicationWorkspace
End Sub

