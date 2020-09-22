VERSION 5.00
Begin VB.Form fOptimize 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Optimize GIF"
   ClientHeight    =   4920
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
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
   Icon            =   "fOptimize.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   328
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdRestore 
      Caption         =   "&Restore"
      Height          =   375
      Left            =   3480
      TabIndex        =   14
      Top             =   1500
      Width           =   1050
   End
   Begin VB.CheckBox chkMoveColorsFromLocalsToGlobal 
      Caption         =   "&Move colors from Locals to Global"
      Height          =   360
      Left            =   375
      TabIndex        =   6
      Top             =   1635
      Width           =   2760
   End
   Begin VB.CheckBox chkRemoveUnusedPaletteEntries 
      Caption         =   "Remove &unused palette entries"
      Height          =   360
      Left            =   375
      TabIndex        =   5
      Top             =   1305
      Width           =   2760
   End
   Begin VB.CheckBox chkCropTransparentImages 
      Caption         =   "&Crop transparent images"
      Height          =   360
      Left            =   375
      TabIndex        =   10
      Top             =   2985
      Width           =   2760
   End
   Begin VB.CheckBox chkRemoveRedundantPixels 
      Caption         =   "Remove &redundant pixels"
      Height          =   360
      Left            =   375
      TabIndex        =   7
      Top             =   1965
      Width           =   2760
   End
   Begin VB.CheckBox chkRemoveComments 
      Caption         =   "Remove co&mments"
      Enabled         =   0   'False
      Height          =   360
      Left            =   375
      TabIndex        =   12
      Top             =   3645
      Value           =   1  'Checked
      Width           =   2760
   End
   Begin VB.OptionButton optRemoveRedundantPixelsMethod 
      Caption         =   "&Minimum bounding rectangle"
      Enabled         =   0   'False
      Height          =   360
      Index           =   0
      Left            =   615
      TabIndex        =   8
      Top             =   2310
      Value           =   -1  'True
      Width           =   2505
   End
   Begin VB.OptionButton optRemoveRedundantPixelsMethod 
      Caption         =   "Frame &differencing"
      Enabled         =   0   'False
      Height          =   360
      Index           =   1
      Left            =   615
      TabIndex        =   9
      Top             =   2640
      Width           =   2475
   End
   Begin GO.ucProgress ucProgress 
      Height          =   165
      Left            =   135
      Top             =   4575
      Width           =   4395
      _ExtentX        =   7752
      _ExtentY        =   291
      BorderStyle     =   1
   End
   Begin VB.CommandButton cmdOptimize 
      Caption         =   "&Optimize"
      Default         =   -1  'True
      Height          =   375
      Left            =   3480
      TabIndex        =   13
      Top             =   1050
      Width           =   1050
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   3480
      TabIndex        =   15
      Top             =   3345
      Width           =   1050
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3480
      TabIndex        =   16
      Top             =   3795
      Width           =   1050
   End
   Begin VB.CheckBox chkDisableInterlacing 
      Caption         =   "Disable &interlacing"
      Height          =   360
      Left            =   375
      TabIndex        =   11
      Top             =   3315
      Width           =   1800
   End
   Begin VB.Label lblOptions 
      AutoSize        =   -1  'True
      Caption         =   " Options "
      Height          =   195
      Left            =   255
      TabIndex        =   4
      Top             =   960
      Width           =   645
   End
   Begin VB.Label lblFileSizeOptimizedV 
      Alignment       =   1  'Right Justify
      Height          =   270
      Left            =   1200
      TabIndex        =   3
      Top             =   510
      Width           =   1215
   End
   Begin VB.Label lblFileSizeOriginalV 
      Alignment       =   1  'Right Justify
      Height          =   270
      Left            =   1200
      TabIndex        =   1
      Top             =   210
      Width           =   1215
   End
   Begin VB.Label lblProgressInfo 
      Height          =   255
      Left            =   150
      TabIndex        =   17
      Top             =   4320
      Width           =   3450
   End
   Begin VB.Label lblFileSizeOptimized 
      Caption         =   "Optimized:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   210
      TabIndex        =   2
      Top             =   510
      Width           =   900
   End
   Begin VB.Label lblFileSizeOriginal 
      Caption         =   "Original:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   210
      TabIndex        =   0
      Top             =   210
      Width           =   975
   End
End
Attribute VB_Name = "fOptimize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'================================================
' Form:          fOptimize.frm
' Last revision: 2003.03.28
'================================================

Option Explicit

Private m_oTmpGIF    As New cGIF ' Temp. GIF
Private m_stInitGIF  As String   ' Initial GIF file path
Private m_stOptmGIF  As String   ' Optimized GIF file path
Private m_bOptimized As Boolean  ' Optimized flag

'//

'========================================================================================
' Form main
'========================================================================================

Private Sub Form_Activate()

  Dim sAppID   As String
  Dim sTmpPath As String

    Screen.MousePointer = vbHourglass
    ucProgress = ucProgress.Max
    DoEvents
    
    '-- Show info
    lblProgressInfo.Caption = "Calculating current GIF file size"
    lblProgressInfo.Refresh
   
    '-- Get temp. GIF file paths
    sTmpPath = IIf(Environ$("tmp") <> "", Environ$("tmp"), Environ$("temp"))
    sTmpPath = sTmpPath & IIf(Right$(sTmpPath, 1) <> "\", "\", "")
    sAppID = CStr(App.hInstance)
    m_stInitGIF = sTmpPath & sAppID & "i.tmp"
    m_stOptmGIF = sTmpPath & sAppID & "o.tmp"
    
    '-- Save...
    g_oGIF.Save m_stInitGIF
    
    '-- Show original and optimized initial file sizes
    lblFileSizeOriginalV.Caption = Format(FileLen(m_stInitGIF), "#,#0 bytes")
    lblFileSizeOptimizedV.Caption = lblFileSizeOriginalV.Caption
    lblProgressInfo.Caption = ""
    
    '-- Deactivate necessary options
    chkMoveColorsFromLocalsToGlobal.Enabled = (g_oGIF.FramesCount > 1)
    chkRemoveRedundantPixels.Enabled = (g_oGIF.FramesCount > 1)
    
    '-- Get settings
    With g_tGO
        chkRemoveUnusedPaletteEntries = -.OGRemoveUnusedPaletteEntries
        chkMoveColorsFromLocalsToGlobal = -.OGMoveColorsFromLocalsToGlobal
        chkRemoveRedundantPixels = -.OGRemoveRedundantPixels
        optRemoveRedundantPixelsMethod(.OGRemoveRedundantPixelsMethod) = -1
        chkCropTransparentImages = -.OGCropTransparentImages
        chkDisableInterlacing = -.OGDisableInterlacing
    End With
    
    ucProgress = 0
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Paint()

    '-- Options frame
    Me.Line (10, 71)-(215, 277), vb3DHighlight, B
    Me.Line (9, 70)-(214, 276), vb3DShadow, B
End Sub

'========================================================================================
' 'Remove redundant pixels' methods
'========================================================================================

Private Sub chkRemoveRedundantPixels_Click()
    optRemoveRedundantPixelsMethod(0).Enabled = -chkRemoveRedundantPixels And -chkRemoveRedundantPixels.Enabled
    optRemoveRedundantPixelsMethod(1).Enabled = -chkRemoveRedundantPixels And -chkRemoveRedundantPixels.Enabled
End Sub

'========================================================================================
' Optimize!
'========================================================================================

Private Sub cmdOptimize_Click()
     
  Dim nFrm As Integer
    
    '== Reset 'optimized' flag
    m_bOptimized = 0
    
    '== Check GIF:
    If ((Me.chkRemoveUnusedPaletteEntries) Or _
        (Me.chkMoveColorsFromLocalsToGlobal And Me.chkMoveColorsFromLocalsToGlobal.Enabled) Or _
        (Me.chkRemoveRedundantPixels And Me.chkRemoveRedundantPixels.Enabled) Or _
        (Me.chkCropTransparentImages) Or _
        (Me.chkDisableInterlacing)) Then
        
        '-- Pre-optimization checking
        lblProgressInfo.Caption = "Checking GIF"
        lblProgressInfo.Refresh
        mGIFOptimizer.CheckGIF g_oGIF, ucProgress
    End If
    
    '== Remove unused palette entries:
    If (Me.chkRemoveUnusedPaletteEntries) Then
        m_bOptimized = -1
        
        '1. Global palette:
        lblProgressInfo.Caption = "Optimizing Global palette"
        lblProgressInfo.Refresh
        
        '-- Optimize...
        mGIFOptimizer.OptimizeGlobalPalette g_oGIF, ucProgress
        
        '2. Local palette/s:
        lblProgressInfo.Caption = "Optimizing Local palettes"
        lblProgressInfo.Refresh
        
        '-- Get number of frames using Local palette
        ucProgress.Max = 0
        For nFrm = 1 To g_oGIF.FramesCount
            If (g_oGIF.LocalPaletteUsed(nFrm)) Then ucProgress.Max = ucProgress.Max + 1
        Next nFrm
        '-- Optimize...
        For nFrm = 1 To g_oGIF.FramesCount
            ucProgress = nFrm
            mGIFOptimizer.OptimizeLocalPalette g_oGIF, nFrm
        Next nFrm
        ucProgress = 0
    End If
    
    '== Move colors from Locals to Global
    If (Me.chkMoveColorsFromLocalsToGlobal And Me.chkMoveColorsFromLocalsToGlobal.Enabled) Then
        m_bOptimized = -1
        
        lblProgressInfo.Caption = "Moving colors from Local/s to Global"
        lblProgressInfo.Refresh
        
        '-- Optimize...
        mGIFOptimizer.RemoveLocalPalettes g_oGIF, ucProgress
    End If
    
    '== Remove redundant pixels
    If (Me.chkRemoveRedundantPixels And Me.chkRemoveRedundantPixels.Enabled) Then
        m_bOptimized = -1
        
        lblProgressInfo.Caption = "Removing redundant pixels"
        lblProgressInfo.Refresh
        
        '-- Optimize...
        mGIFOptimizer.OptimizeFrames g_oGIF, IIf(optRemoveRedundantPixelsMethod(0), 0, 1), ucProgress
    End If
    
    '== Crop transparent images
    If (Me.chkCropTransparentImages) Then
        m_bOptimized = -1

        lblProgressInfo.Caption = "Cropping transparent images"
        lblProgressInfo.Refresh

        '-- Optimize...
        mGIFOptimizer.CropTransparentImages g_oGIF, ucProgress
    End If
    
    '== Disable interlacing
    If (Me.chkDisableInterlacing) Then
        m_bOptimized = -1
        
        '-- Optimize...
        For nFrm = 1 To g_oGIF.FramesCount
            g_oGIF.FrameInterlaced(nFrm) = 0
        Next nFrm
    End If
    
    If (m_bOptimized) Then
    
        '-- Post-optimization checking
        lblProgressInfo.Caption = "Checking GIF"
        lblProgressInfo.Refresh
        mGIFOptimizer.CheckGIF g_oGIF, ucProgress
        
        '-- Remask frames
        lblProgressInfo.Caption = "Remasking frames"
        lblProgressInfo.Refresh
        mGIFOptimizer.RemaskFrames g_oGIF, ucProgress
        
        '-- Calculating optimized file size
        ucProgress = ucProgress.Max
        lblProgressInfo.Caption = "Calculating optimized file size"
        lblProgressInfo.Refresh
        g_oGIF.Save m_stOptmGIF
        ucProgress = 0
        
        '-- Update labels
        lblFileSizeOptimizedV.Caption = Format(FileLen(m_stOptmGIF), "#,#0 bytes")
        lblProgressInfo.Caption = ""
        
        '-- Destroy temp. optimized GIF
        On Error Resume Next
        Kill m_stOptmGIF
        
      Else
        '-- No optimization option checked
        lblFileSizeOptimizedV.Caption = lblFileSizeOriginalV.Caption
        lblProgressInfo.Caption = ""
    End If
End Sub

'========================================================================================
' Restore GIF
'========================================================================================

Private Sub cmdRestore_Click()
    
    ucProgress = ucProgress.Max
    
    '-- Show info
    lblProgressInfo.Caption = "Restoring initial GIF"
    lblProgressInfo.Refresh
    
    '-- Need to restore [?]
    If (m_bOptimized) Then
        
        '-- Load previous
        g_oGIF.LoadFromFile m_stInitGIF
        '-- Show current file sizes
        lblFileSizeOriginalV.Caption = Format(FileLen(m_stInitGIF), "#,#0 bytes")
        lblFileSizeOptimizedV.Caption = lblFileSizeOriginalV.Caption
        '-- Disable 'optimized' flag
        m_bOptimized = 0
    End If
    lblProgressInfo = ""
    
    ucProgress = 0
End Sub

'========================================================================================
' Accept / Cancel
'========================================================================================

Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub cmdCancel_Click()

    '-- GIF has been optimized
    If (m_bOptimized) Then
    
        '-- Restore to previous...
        ucProgress = ucProgress.Max
        lblProgressInfo.Caption = "Restoring initial GIF"
        lblProgressInfo.Refresh
        g_oGIF.LoadFromFile m_stInitGIF
    End If
    
    Unload Me
End Sub

'========================================================================================
' Unloading
'========================================================================================

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    '-- Close button pressed
    If (UnloadMode = vbFormControlMenu) Then
        cmdCancel_Click
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    '-- Store settings
    With g_tGO
        .OGRemoveUnusedPaletteEntries = -chkRemoveUnusedPaletteEntries
        .OGMoveColorsFromLocalsToGlobal = -chkMoveColorsFromLocalsToGlobal
        .OGRemoveRedundantPixels = -chkRemoveRedundantPixels
        .OGRemoveRedundantPixelsMethod = IIf(optRemoveRedundantPixelsMethod(0), 0, 1)
        .OGCropTransparentImages = -chkCropTransparentImages
        .OGDisableInterlacing = -chkDisableInterlacing
    End With
    
    '-- Destroy temp. GIF
    m_oTmpGIF.Destroy
    '-- Kill temp. file
    On Error Resume Next
    Kill m_stInitGIF
    
    '-- Yes, necessary
    Set fOptimize = Nothing
End Sub
