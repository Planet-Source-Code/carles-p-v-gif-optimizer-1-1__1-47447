VERSION 5.00
Begin VB.Form fPreview 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Preview"
   ClientHeight    =   3600
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4800
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "fPreview.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer tmrDelay 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   90
      Top             =   105
   End
End
Attribute VB_Name = "fPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'================================================
' Form:          fPreview.frm
' Last revision: 2003.08.04
'================================================

Option Explicit

Private m_nFrames       As Integer
Private m_nFrame        As Integer
Private m_oFrameBuffDIB As New cDIB
Private m_oRestoringDIB As New cDIB
Private m_oBackground   As New cTile

Private m_xOffset       As Long
Private m_yOffset       As Long

'//

'========================================================================================
' Form main
'========================================================================================

Private Sub Form_Load()
    
    '-- Resize me for fit animation screen
    pvResizeMe
    
    With g_oGIF
    
        '-- Get number of frames
        m_nFrames = .FramesCount
        
        '-- Create DIBs
        m_oFrameBuffDIB.Create .ScreenWidth, .ScreenHeight, [24_bpp]
        m_oRestoringDIB.Create .ScreenWidth, .ScreenHeight, [24_bpp]
        '-- Create background pattern (solid color)
        m_oBackground.SetPatternFromSolidColor Me.BackColor
        m_oBackground.Tile m_oFrameBuffDIB.hDC, 0, 0, .ScreenWidth, .ScreenHeight
        
        '-- Start animation [?]
        If (m_nFrames > 1) Then
            m_nFrame = m_nFrames
            tmrDelay.Enabled = -1
          Else
            .FrameDraw m_oFrameBuffDIB.hDC, 1
        End If
    End With
End Sub

Private Sub Form_DblClick()
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    
    '-- [Esc] / [F5]
    If (KeyCode = vbKeyEscape Or KeyCode = vbKeyF5) Then
        Unload Me
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
  
    '-- Destroy objects
    Set m_oFrameBuffDIB = Nothing
    Set m_oRestoringDIB = Nothing
    Set m_oBackground = Nothing
End Sub

'========================================================================================
' Timing
'========================================================================================

Private Sub tmrDelay_Timer()
    
    '-- Next frame / First
    If (m_nFrame < m_nFrames) Then
        m_nFrame = m_nFrame + 1
      Else
        m_nFrame = 1
    End If
    pvFrame_Change
End Sub

'========================================================================================
' Frame rendering
'========================================================================================

Private Sub pvFrame_Change()
    
    With g_oGIF
    
        '-- Set current frame delay
        Select Case .FrameDelay(m_nFrame)
            Case Is < 0
                tmrDelay.Interval = 60000 ' Max.: 1 min.
            Case Is = 0
                tmrDelay.Interval = 100   ' Def.: 0.1 sec.
            Case Is < 5
                tmrDelay.Interval = 50    ' Min.: 0.05 sec.
            Case Else
                tmrDelay.Interval = .FrameDelay(m_nFrame) * 10
        End Select
        
        '-- Restore:
        If (m_nFrame = 1) Then
            m_oBackground.Tile m_oRestoringDIB.hDC, 0, 0, .ScreenWidth, .ScreenHeight
        End If
        m_oFrameBuffDIB.LoadBlt m_oRestoringDIB.hDC
        
        '-- Draw current frame:
        .FrameDraw m_oFrameBuffDIB.hDC, m_nFrame
        
        '-- Update restoring buffer:
        Select Case .FrameDisposalMethod(m_nFrame)
            Case [dmNotSpecified], [dmDoNotDispose]
                '-- Update from current
                m_oRestoringDIB.LoadBlt m_oFrameBuffDIB.hDC
            Case [dmRestoreToBackground]
                '-- Update from background
                m_oBackground.Tile m_oRestoringDIB.hDC, .FrameLeft(m_nFrame), .FrameTop(m_nFrame), .FrameDIBXOR(m_nFrame).Width, .FrameDIBXOR(m_nFrame).Height, 0
            Case [dmRestoreToPrevious]
                '-- Preserve buffer
        End Select
    End With
    
    '-- Paint frame
    Form_Paint
End Sub

Private Sub Form_Paint()
    '-- Paint to canvas
    m_oFrameBuffDIB.Stretch hDC, m_xOffset, m_yOffset, g_oGIF.ScreenWidth, g_oGIF.ScreenHeight
End Sub

'========================================================================================
' Private
'========================================================================================

Private Sub pvResizeMe()

  Dim ext_W As Long
  Dim ext_H As Long
  Const BDR As Long = 10
    
    With g_oGIF
    
        ext_W = (Me.Width \ Screen.TwipsPerPixelX) - Me.ScaleWidth + BDR
        ext_H = (Me.Height \ Screen.TwipsPerPixelY) - Me.ScaleHeight + BDR
        
        '-- Resize form [?]
        If (Me.ScaleWidth < .ScreenWidth + BDR) Then
            Me.Width = (.ScreenWidth + ext_W) * Screen.TwipsPerPixelX
        End If
        If (Me.ScaleHeight < .ScreenHeight + BDR) Then
            Me.Height = (.ScreenHeight + ext_H) * Screen.TwipsPerPixelY
        End If
        
        '-- Calc. screen offsets
        m_xOffset = (Me.ScaleWidth - .ScreenWidth) \ 2
        m_yOffset = (Me.ScaleHeight - .ScreenHeight) \ 2
    End With
End Sub
