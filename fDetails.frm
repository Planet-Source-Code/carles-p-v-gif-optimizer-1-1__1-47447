VERSION 5.00
Begin VB.Form fDetails 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "GIF details"
   ClientHeight    =   6630
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5955
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
   Icon            =   "fDetails.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   442
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   397
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin GO.ucProgress ucDummyProgress 
      Height          =   180
      Left            =   2895
      Top             =   6120
      Visible         =   0   'False
      Width           =   585
      _ExtentX        =   1032
      _ExtentY        =   318
   End
   Begin VB.CommandButton cmdCheckGIF 
      Caption         =   "&Check GIF"
      Height          =   375
      Left            =   3525
      TabIndex        =   23
      Top             =   6120
      Width           =   1050
   End
   Begin VB.CheckBox chkScrollMode 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Caption         =   "S"
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   2010
      TabIndex        =   17
      TabStop         =   0   'False
      ToolTipText     =   "Enable/Disable [Scroll] mode"
      Top             =   5340
      Width           =   360
   End
   Begin VB.CheckBox chkFitMode 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Caption         =   "F"
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   2430
      TabIndex        =   18
      TabStop         =   0   'False
      ToolTipText     =   "Enable/Disable [Best Fit] mode"
      Top             =   5340
      Value           =   1  'Checked
      Width           =   360
   End
   Begin VB.Timer tmrFlashTrns 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   5250
      Top             =   3285
   End
   Begin VB.PictureBox picPalette 
      AutoRedraw      =   -1  'True
      BackColor       =   &H8000000C&
      ClipControls    =   0   'False
      ForeColor       =   &H00E0E0E0&
      Height          =   1995
      Left            =   3180
      MousePointer    =   99  'Custom
      ScaleHeight     =   129
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   129
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   3285
      Width           =   1995
      Begin VB.Shape shpTrnsIdx 
         BorderColor     =   &H00FFC0FF&
         Height          =   135
         Left            =   0
         Top             =   0
         Width           =   135
      End
      Begin VB.Shape shpSelect 
         BorderColor     =   &H00FFFFFF&
         Height          =   135
         Left            =   0
         Top             =   0
         Width           =   135
      End
   End
   Begin GO.ucCanvas ucCanvas 
      Height          =   1995
      Left            =   195
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   3285
      Width           =   2610
      _ExtentX        =   4604
      _ExtentY        =   3519
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   4710
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   6120
      Width           =   1050
   End
   Begin VB.ListBox lstInfo 
      Appearance      =   0  'Flat
      Height          =   2370
      Left            =   225
      TabIndex        =   8
      Top             =   510
      Visible         =   0   'False
      Width           =   5505
   End
   Begin VB.Label lblPositionV 
      Height          =   195
      Left            =   900
      TabIndex        =   14
      Top             =   5340
      Width           =   1095
   End
   Begin VB.Label lblSizeV 
      Height          =   195
      Left            =   900
      TabIndex        =   16
      Top             =   5580
      Width           =   1395
   End
   Begin VB.Label lblPosition 
      Caption         =   "Position:"
      Height          =   195
      Left            =   210
      TabIndex        =   13
      Top             =   5340
      Width           =   795
   End
   Begin VB.Label lblSize 
      Caption         =   "Size:"
      Height          =   195
      Left            =   210
      TabIndex        =   15
      Top             =   5580
      Width           =   555
   End
   Begin VB.Line lnSep 
      X1              =   15
      X2              =   382
      Y1              =   33
      Y2              =   33
   End
   Begin VB.Label lblRGBV 
      Height          =   195
      Left            =   3735
      TabIndex        =   22
      Top             =   5580
      Width           =   1125
   End
   Begin VB.Label lblIndexV 
      Height          =   195
      Left            =   3735
      TabIndex        =   20
      Top             =   5340
      Width           =   495
   End
   Begin VB.Label lblRGB 
      Caption         =   "R,G,B:"
      Height          =   195
      Left            =   3180
      TabIndex        =   21
      Top             =   5580
      Width           =   630
   End
   Begin VB.Label lblIndex 
      Caption         =   "Index:"
      Height          =   195
      Left            =   3180
      TabIndex        =   19
      Top             =   5340
      Width           =   585
   End
   Begin VB.Label lblPalette 
      Caption         =   "Palette"
      Height          =   240
      Left            =   3195
      TabIndex        =   11
      Top             =   3030
      Width           =   1245
   End
   Begin VB.Label lblImage 
      Caption         =   "Image"
      Height          =   240
      Left            =   210
      TabIndex        =   9
      Top             =   3030
      Width           =   810
   End
   Begin VB.Label lblHeader 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Interl."
      Height          =   195
      Index           =   6
      Left            =   4740
      TabIndex        =   7
      Top             =   270
      Width           =   450
   End
   Begin VB.Label lblHeader 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Delay"
      Height          =   195
      Index           =   5
      Left            =   4185
      TabIndex        =   6
      Top             =   270
      Width           =   405
   End
   Begin VB.Label lblHeader 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Disposal mode"
      Height          =   195
      Index           =   4
      Left            =   2925
      TabIndex        =   5
      Top             =   270
      Width           =   1020
   End
   Begin VB.Label lblHeader 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Idx."
      Height          =   195
      Index           =   3
      Left            =   2400
      TabIndex        =   4
      Top             =   270
      Width           =   300
   End
   Begin VB.Label lblHeader 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Transp."
      Height          =   195
      Index           =   2
      Left            =   1590
      TabIndex        =   3
      Top             =   270
      Width           =   555
   End
   Begin VB.Label lblHeader 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Palette"
      Height          =   195
      Index           =   1
      Left            =   930
      TabIndex        =   2
      Top             =   270
      Width           =   510
   End
   Begin VB.Label lblHeader 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Frame"
      Height          =   195
      Index           =   0
      Left            =   315
      TabIndex        =   1
      Top             =   270
      Width           =   450
   End
   Begin VB.Label lblBorder 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   2655
      Left            =   195
      TabIndex        =   0
      Top             =   225
      Width           =   5565
   End
End
Attribute VB_Name = "fDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'================================================
' Form:          fInfo.frm
' Last revision: 2003.08.04
'================================================

Option Explicit

'-- API:

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const LB_SETTABSTOPS As Long = &H192

'//

Private m_nFrame As Integer

'//

'========================================================================================
' Form main
'========================================================================================

Private Sub Form_Load()

  Dim lTAB(6) As Long
    
    DoEvents
  
    '-- Set listbox column TABs
    lTAB(0) = -25
    lTAB(1) = -55
    lTAB(2) = -85
    lTAB(3) = -110
    lTAB(4) = 120
    lTAB(5) = -195
    lTAB(6) = -220
    SendMessage lstInfo.hWnd, LB_SETTABSTOPS, 7, lTAB(0)
    
    '-- Remove listbox control border
    mMisc.RemoveBorder lstInfo.hWnd

    '-- Fill list
    pvFillList
    
    '-- Set canvas modes: 'Best Fit' / 'User Mode'
    ucCanvas.FitMode = -1
    ucCanvas.WorkMode = [cnvUserMode]
    
    '--  Load palette grid canvas cursors
    picPalette.MouseIcon = LoadResPicture("CURSOR_PICKER", vbResCursor)
    Set ucCanvas.UserIcon = LoadResPicture("CURSOR_PICKER", vbResCursor)
    
    '-- Select current frame
    lstInfo.ListIndex = g_nFrame - 1
End Sub

Private Sub Form_Paint()

    '-- Some decorative lines
    Me.Line (0, 395)-(ScaleWidth, 395), vb3DShadow
    Me.Line (0, 396)-(ScaleWidth, 396), vb3DHighlight
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    
    '-- [F1]
    If (KeyCode = vbKeyF1) Then
        Unload Me
    End If
End Sub

'========================================================================================
' Item selected
'========================================================================================

Private Sub lstInfo_Click()
    
    '-- Current frame
    m_nFrame = lstInfo.ListIndex + 1
    
    With g_oGIF
    
        '-- Show frame info
        lblPositionV = .FrameLeft(m_nFrame) & "," & .FrameTop(m_nFrame)
        With .FrameDIBXOR(m_nFrame)
            lblSizeV = .Width & "x" & .Height
        End With
        
        '-- Create canvas
        With .FrameDIBXOR(m_nFrame)
            ucCanvas.DIB.Create .Width, .Height, [32_bpp]
        End With
        ucCanvas.Resize
        '-- Erase canvas background (transparent color) and draw frame
        If (.FrameUseTransparentColor(m_nFrame)) Then
            ucCanvas.DIB.Cls .LocalPaletteRGBEntry(m_nFrame, .FrameTransparentColorIndex(m_nFrame))
        End If
        g_oGIF.FrameDraw ucCanvas.DIB.hDC, m_nFrame, -.FrameLeft(m_nFrame), -.FrameTop(m_nFrame)
        ucCanvas.Repaint
        
        '-- Paint palette grid and 'select' first entry
        Call pvPaintPalette
        Call picPalette_MouseDown(vbLeftButton, 0, 0, 0)
        
        '-- Activate/Deactivate transparent index cursor shape
        tmrFlashTrns.Enabled = g_oGIF.FrameUseTransparentColor(m_nFrame)
        shpTrnsIdx.Move (g_oGIF.FrameTransparentColorIndex(m_nFrame) Mod 16) * 8, (g_oGIF.FrameTransparentColorIndex(m_nFrame) \ 16) * 8
        shpTrnsIdx.Visible = tmrFlashTrns.Enabled
    End With
End Sub

'========================================================================================
' Canvas control
'========================================================================================

Private Sub chkScrollMode_Click()
        
    '-- Scroll mode
    With ucCanvas
        .WorkMode = 1 - .WorkMode
    End With
End Sub

Private Sub chkFitMode_Click()

    '-- Fit mode
    With ucCanvas
        .FitMode = Not .FitMode
        .Resize
        .Repaint
    End With
End Sub

Private Sub ucCanvas_MouseDown(Button As Integer, Shift As Integer, x As Long, y As Long)
    ucCanvas_MouseMove vbLeftButton, Shift, x, y
End Sub

Private Sub ucCanvas_MouseMove(Button As Integer, Shift As Integer, x As Long, y As Long)
'-- Picking color from canvas

  Dim aIdx As Byte

    If (Button = vbLeftButton And ucCanvas.WorkMode = [cnvUserMode]) Then
    
        '-- Check bounds
        With ucCanvas.DIB
            If (x < 0) Then x = 0
            If (y < 0) Then y = 0
            If (x > .Width - 1) Then x = .Width - 1
            If (y > .Height - 1) Then y = .Height - 1
        End With
        
        '-- Get palette index from point
        aIdx = mGIFRemaper.PaletteIndex08(g_oGIF.FrameDIBXOR(m_nFrame), x, y)
    
        '-- Force 'picPalette_MouseMove' for labels update
        picPalette_MouseMove vbLeftButton, Shift, (aIdx Mod 16) * 8, (aIdx \ 16) * 8
    End If
End Sub

'========================================================================================
' Palette control
'========================================================================================

Private Sub picPalette_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
      
    '-- Force 'picPalette_MouseMove' for labels update
    picPalette_MouseMove Button, Shift, x, y
End Sub


Private Sub picPalette_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'-- Picking color from palette grid

  Dim aIdx As Byte
  Dim nFrm As Integer
  Dim lClr As Long

    If (Button = vbLeftButton) Then
        
        '-- Check bounds
        With picPalette
            If (x < 0) Then x = 0
            If (y < 0) Then y = 0
            If (x > .ScaleWidth - 2) Then x = .ScaleWidth - 2
            If (y > .ScaleHeight - 2) Then y = .ScaleHeight - 2
        End With

        '-- Get palette index and color
        aIdx = x \ 8 + (y \ 8) * 16
        lClr = g_oGIF.LocalPaletteRGBEntry(m_nFrame, aIdx): g_lColor = lClr
        
        '-- Locate cursor
        shpSelect.Move (aIdx Mod 16) * 8, (aIdx \ 16) * 8
        
        '-- Show info
        lblIndexV = aIdx
        lblRGBV = (lClr And &HFF&) & "," & (lClr And &HFF00&) \ 256 & "," & (lClr And &HFF0000) \ 65536
    End If
End Sub

Private Sub tmrFlashTrns_Timer()
    '-- Flash transparent index shape cursor
    shpTrnsIdx.BorderStyle = 1 - shpTrnsIdx.BorderStyle
End Sub

'========================================================================================
' Check GIF / Exit
'========================================================================================

Private Sub cmdCheckGIF_Click()
    
    Screen.MousePointer = vbHourglass
    
    '-- Check GIF...
    mGIFOptimizer.CheckGIF g_oGIF, ucDummyProgress
    
    '-- Fill list and select first item
    pvFillList
    lstInfo.ListIndex = 0
    
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdOK_Click()
    Unload Me
End Sub

'========================================================================================
' Private
'========================================================================================

Private Sub pvFillList()

  Dim nFrm As Integer
    
    With g_oGIF
        
        '-- Hide list
        lstInfo.Visible = 0
        
        '-- Clear list
        lstInfo.Clear
        
        '-- Fill list
        For nFrm = 1 To .FramesCount
        
            lstInfo.AddItem _
                vbTab & nFrm & _
                vbTab & IIf(.LocalPaletteUsed(nFrm), "L-", "G-") & .LocalPaletteEntries(nFrm) & _
                vbTab & IIf(.FrameUseTransparentColor(nFrm), "Yes", "No") & _
                vbTab & IIf(.FrameUseTransparentColor(nFrm), .FrameTransparentColorIndex(nFrm), "-") & _
                vbTab & Choose(.FrameDisposalMethod(nFrm) + 1, "Not especified", "Do not dispose", "To background", "To previous") & _
                vbTab & .FrameDelay(nFrm) & _
                vbTab & IIf(.FrameInterlaced(nFrm), "Yes", "No")
        Next nFrm
        lstInfo.Visible = -1
    End With
End Sub

Private Sub pvPaintPalette()
'-- Paint palette grid
  
  Dim i As Long, j As Long
  Dim lIdx As Long
  Dim lClr As Long
    
    '-- Paint palette grid
    For i = 1 To 121 Step 8
        For j = 1 To 121 Step 8
            With g_oGIF
                If (lIdx < .LocalPaletteEntries(m_nFrame)) Then
                    lClr = .LocalPaletteRGBEntry(m_nFrame, lIdx)
                  Else
                    lClr = -1
                End If
                mMisc.DrawRectangle picPalette.hDC, j, i, j + 7, i + 7, lClr
                lIdx = lIdx + 1
            End With
        Next j
    Next i
    picPalette.Refresh
End Sub
