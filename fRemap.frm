VERSION 5.00
Begin VB.Form fRemap 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Remap GIF"
   ClientHeight    =   4200
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
   Icon            =   "fRemap.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   280
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdGIFdetails 
      Caption         =   "..."
      Height          =   300
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   11
      TabStop         =   0   'False
      ToolTipText     =   "Choose color from frame palette..."
      Top             =   2940
      Width           =   315
   End
   Begin VB.TextBox txtPaletteEntriesCustomEntries 
      Height          =   315
      Left            =   645
      MaxLength       =   3
      TabIndex        =   8
      Top             =   2370
      Width           =   660
   End
   Begin VB.CheckBox chkPreserveColor 
      Caption         =   "Preserve color"
      Height          =   210
      Left            =   390
      TabIndex        =   9
      Top             =   2985
      Width           =   1380
   End
   Begin VB.ComboBox cbPaletteEntriesFixedBPP 
      Height          =   315
      Left            =   630
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   1635
      Width           =   2265
   End
   Begin VB.OptionButton optPaletteEntriesMode 
      Caption         =   "&Custom [8-256]"
      CausesValidation=   0   'False
      Height          =   345
      Index           =   1
      Left            =   375
      TabIndex        =   7
      Top             =   2040
      Width           =   1905
   End
   Begin VB.OptionButton optPaletteEntriesMode 
      Caption         =   "&Fixed by color depth"
      CausesValidation=   0   'False
      Height          =   345
      Index           =   0
      Left            =   375
      TabIndex        =   5
      Top             =   1305
      Value           =   -1  'True
      Width           =   1905
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3480
      TabIndex        =   16
      Top             =   3075
      Width           =   1050
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   3480
      TabIndex        =   15
      Top             =   2625
      Width           =   1050
   End
   Begin VB.CommandButton cmdReduce 
      Caption         =   "Re&duce"
      Default         =   -1  'True
      Height          =   375
      Left            =   3480
      TabIndex        =   12
      Top             =   1050
      Width           =   1050
   End
   Begin VB.CommandButton cmdRestore 
      Caption         =   "&Restore"
      Height          =   375
      Left            =   3480
      TabIndex        =   13
      Top             =   1500
      Width           =   1050
   End
   Begin GO.ucProgress ucProgress 
      Height          =   165
      Left            =   135
      Top             =   3855
      Width           =   4395
      _ExtentX        =   7752
      _ExtentY        =   291
      BorderStyle     =   1
   End
   Begin VB.CommandButton cmdPreview 
      Caption         =   "Pre&view"
      Height          =   375
      Left            =   3480
      TabIndex        =   14
      Top             =   1965
      Width           =   1050
   End
   Begin VB.Label lblPreservedColor 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Left            =   1815
      MousePointer    =   99  'Custom
      TabIndex        =   10
      Top             =   2940
      Width           =   300
   End
   Begin VB.Label lblProgressInfo 
      Height          =   255
      Left            =   150
      TabIndex        =   17
      Top             =   3600
      Width           =   3450
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
   Begin VB.Label lblFileSizeOriginalV 
      Alignment       =   1  'Right Justify
      Height          =   270
      Left            =   1200
      TabIndex        =   1
      Top             =   210
      Width           =   1215
   End
   Begin VB.Label lblFileSizeOptimizedV 
      Alignment       =   1  'Right Justify
      Height          =   270
      Left            =   1200
      TabIndex        =   3
      Top             =   510
      Width           =   1215
   End
   Begin VB.Label lblPaletteEntries 
      AutoSize        =   -1  'True
      Caption         =   " Palette entries "
      Height          =   195
      Left            =   255
      TabIndex        =   4
      Top             =   960
      Width           =   1140
   End
End
Attribute VB_Name = "fRemap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'================================================
' Form:          fRemap.frm
' Last revision: 2003.08.07
'================================================

Option Explicit

'-- API:

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

'//

Private m_oTmpGIF    As New cGIF ' Temp. GIF
Private m_stInitGIF  As String   ' Initial GIF file path
Private m_stOptmGIF  As String   ' Optimized GIF file path
Private m_bOptimized As Boolean  ' Optimized flag
Private m_bActivated As Boolean

'//

'========================================================================================
' Form main
'========================================================================================

Private Sub Form_Activate()

  Dim sAppID   As String
  Dim sTmpPath As String
  Dim nItm     As Integer
    
    If (Not m_bActivated) Then
    
        m_bActivated = -1
   
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
        
        '-- Fill combo
        For nItm = 3 To 8
            cbPaletteEntriesFixedBPP.AddItem nItm & "-bpp (max. " & 2 ^ nItm & " colors)"
        Next nItm
        cbPaletteEntriesFixedBPP.SetFocus
        
        '-- Load color-picker cursor
        lblPreservedColor.MouseIcon = LoadResPicture("CURSOR_PICKER", vbResCursor)
        
        '-- A little effect
        SendMessage cmdGIFdetails.hWnd, &HF4&, &H0&, 0&
        
        '-- Get settings
        With g_tGO
            optPaletteEntriesMode(.RGPaletteEntriesMode) = -1
            cbPaletteEntriesFixedBPP.ListIndex = .RGPaletteEntriesFixedBPP
            txtPaletteEntriesCustomEntries.Text = .RGPaletteEntriesCustomEntries
            chkPreserveColor = -.RGPreserveColor
            lblPreservedColor.BackColor = .RGPreservedColor
        End With
        
        ucProgress = 0
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub Form_Paint()

    '-- Options frame
    Me.Line (10, 71)-(215, 229), vb3DHighlight, B
    Me.Line (9, 70)-(214, 228), vb3DShadow, B
End Sub

'========================================================================================
' Options
'========================================================================================

Private Sub txtPaletteEntriesCustomEntries_GotFocus()

    With txtPaletteEntriesCustomEntries
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub

Private Sub txtPaletteEntriesCustomEntries_KeyPress(KeyAscii As Integer)
    If ((KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8) Then KeyAscii = 0
End Sub

Private Sub txtPaletteEntriesCustomEntries_Change()
    
    With txtPaletteEntriesCustomEntries
        If (Val(.Text) > 256) Then
            .Text = "256"
        End If
        If (Val(.Text) = 0) Then
            .Text = "0"
            .SelStart = 0: .SelLength = 1
          Else
            .SelStart = .MaxLength
        End If
    End With
End Sub

Private Sub txtPaletteEntriesCustomEntries_KeyDown(KeyCode As Integer, Shift As Integer)

    With txtPaletteEntriesCustomEntries
        Select Case KeyCode
            Case vbKeyUp
                If (Val(.Text) < 256) Then .Text = Val(.Text) + 1
                KeyCode = 0
            Case vbKeyDown
                If (Val(.Text) > 0) Then .Text = Val(.Text) - 1
                KeyCode = 0
        End Select
    End With
End Sub

Private Sub lblPreservedColor_Click()
    
  Dim lClr As Long
    
    '-- Select color...
    lClr = mDialogColor.SelectColor(Me.hWnd, lblPreservedColor.BackColor, -1)
    
    If (lClr <> -1) Then
        lblPreservedColor.BackColor = lClr
    End If
End Sub

Private Sub cmdGIFdetails_Click()

    '-- Show GIF details form
    fDetails.Show vbModal, Me
    lblPreservedColor.BackColor = g_lColor
End Sub

'========================================================================================
' Reduce!
'========================================================================================

Private Sub cmdReduce_Click()
  
  Dim nFrm     As Integer
  Dim nEntries As Integer
  
    '-- Turn on 'optimized' flag
    m_bOptimized = -1
    
    '-- Get number of entries
    Select Case True
        Case optPaletteEntriesMode(0)
            nEntries = 2 ^ (cbPaletteEntriesFixedBPP.ListIndex + 3)
        Case optPaletteEntriesMode(1)
            If (Val(txtPaletteEntriesCustomEntries.Text) < 8) Then
                txtPaletteEntriesCustomEntries.Text = "8"
            End If
            nEntries = Val(txtPaletteEntriesCustomEntries.Text)
    End Select
  
    '-- Pre-optimization checking
    lblProgressInfo.Caption = "Checking GIF"
    lblProgressInfo.Refresh
    mGIFOptimizer.CheckGIF g_oGIF, ucProgress
    
    '-- Pre-optimize palettes
        
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
        
    '-- Now, we can start remaping
    lblProgressInfo.Caption = "Remaping frames"
    lblProgressInfo.Refresh
    Set mGIFRemaper.Palette = mGIFRemaper.GIFOptimalPalette(g_oGIF, nEntries, IIf(chkPreserveColor, lblPreservedColor.BackColor, -1))
    mGIFRemaper.RemapGIF g_oGIF, ucProgress
     
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
End Sub

'========================================================================================
' Preview
'========================================================================================

Private Sub cmdPreview_Click()
    '-- Preview animation
    fPreview.Show vbModal, Me
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
        .RGPaletteEntriesMode = IIf(optPaletteEntriesMode(0), 0, 1)
        .RGPaletteEntriesFixedBPP = cbPaletteEntriesFixedBPP.ListIndex
        .RGPaletteEntriesCustomEntries = Val(txtPaletteEntriesCustomEntries.Text)
        .RGPreserveColor = -chkPreserveColor
        .RGPreservedColor = lblPreservedColor.BackColor
    End With
        
    '-- Destroy temp. GIF
    m_oTmpGIF.Destroy
    '-- Kill temp. file
    On Error Resume Next
    Kill m_stInitGIF
    
    '-- Yes, necessary
    Set fRemap = Nothing
End Sub
