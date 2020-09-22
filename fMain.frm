VERSION 5.00
Begin VB.Form fMain 
   Caption         =   "GO v1.1"
   ClientHeight    =   6420
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   8760
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
   Icon            =   "fMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   428
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   584
   StartUpPosition =   2  'CenterScreen
   Begin GO.ucInfo ucInfo 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      Top             =   6150
      Width           =   8760
      _ExtentX        =   15452
      _ExtentY        =   476
   End
   Begin GO.ucFramesList ucFramesList 
      Height          =   5745
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   15
      Width           =   2940
      _ExtentX        =   5186
      _ExtentY        =   10134
      ThumbnailHeight =   100
   End
   Begin GO.ucCanvas ucCanvas 
      Height          =   5745
      Left            =   2985
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   15
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   10134
   End
   Begin VB.Menu mnuFileTop 
      Caption         =   "&File"
      Begin VB.Menu mnuFile 
         Caption         =   "&Open GIF..."
         Index           =   0
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFile 
         Caption         =   "&Save GIF..."
         Index           =   1
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFile 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuFile 
         Caption         =   "E&xit"
         Index           =   3
      End
   End
   Begin VB.Menu mnuViewTop 
      Caption         =   "&View"
      Begin VB.Menu mnuView 
         Caption         =   "GIF &details"
         Index           =   0
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuView 
         Caption         =   "&Preview animation"
         Index           =   1
         Shortcut        =   {F5}
      End
   End
   Begin VB.Menu mnuToolsTop 
      Caption         =   "&Tools"
      Begin VB.Menu mnuTools 
         Caption         =   "&Optimize GIF..."
         Index           =   0
         Shortcut        =   {F6}
      End
      Begin VB.Menu mnuTools 
         Caption         =   "Re&map GIF..."
         Index           =   1
         Shortcut        =   {F7}
      End
      Begin VB.Menu mnuTools 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuTools 
         Caption         =   "&Restore frames"
         Index           =   3
         Shortcut        =   {F8}
      End
   End
   Begin VB.Menu mnuHelpTop 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelp 
         Caption         =   "&About"
         Index           =   0
      End
   End
End
Attribute VB_Name = "fMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'================================================
' Project:       GO [1.1 rev. 09] - GIF optimizer
' Author:        Carles P.V.
' Last revision: 2003.12.04
'================================================
' Commercial use not permitted!
' Email author/s please.
'================================================
'
' Work based on:
'
'  - TGIFImage v.2.2 by Anders Melander:  http://www.torry.net/gif.htm
'  - VB Gif Library by Vlad Vissoultchev: http://www.pscode.com/vb/scripts/ShowCode.asp?txtCodeId=44216&lngWId=1
'  - RVTVBIMG by Ron van Tilburg :        http://www.pscode.com/vb/scripts/ShowCode.asp?txtCodeId=14210&lngWId=1
'
'========================================================================================
'
' Some notes about optimization routines:
'
'  - Pre-optimization checking. This routine previously scans GIF for
'    incoherent/incorrect  settings. If  this  checking is  not done,
'    some GIFs/frames could be  not optimized. This  routine has been
'    specialy improved.
'  - In palette  optimizations, entries  are not sorted by frequency.
'  - About 'Move colors from  Local/s to Global' optimization option:
'    A bug has  been fixed from  original code. Repeated  colors were
'    not  taken  in  account. When  one  of  these  entries  is  used
'    as transparent one, this can turn to transparent not transparent
'    pixels.
'  - Added 'Minimum Bounding Rectangle' method: When palette is full,
'    it is the only  method that allows  to reduce GIF  file size. By
'    the other  hand, this  method can, sometimes, be  more efficient
'    than 'Frame Differencing'.
'  - Added 'Disable interlacing' option.
'  - Current GIF class does not support comments, so current GIF file
'    size does not include them.
'
' Advices:
'
'  - Some GIFs can  reach the optimum file size  after a second pass.
'  - Don't  apply different  optimization  methods (Remove  redundant
'    pixels) during same optimization process. Restore first.
'
' Last notes (next versions):
'
'  - 'Remap GIF' and 'Frame restoring' are  not available yet. 'Remap
'    GIF' will be a last GIF optimization feature, but, in this case,
'    we will have loss of final quality.
'
' More info:
'
'  Take a look at: http://www.webreference.com/dev/gifanim/
'
'========================================================================================
'
' LOG:
'
'  - 2003.08.05: Improved 'Scan for redundant disposal modes'.
'  - 2003.08.08: Improved 'CropTransparentImages' function.
'                Only crops image if necessary (initial and
'                final rectangles are compared).
'  - 2003.08.11: Added 'Remap GIF' optimization.
'                This is a destructive optimization. Frames final
'                image quality could be affected. With this new
'                tool, we can eliminate those local palettes
'                that were not possible to remove using standard
'                not-destructive methods or, simply, reduce the
'                current global palette depth.
'  - 2003.08.11: Fixed 'Scan for redundant disposal modes'.
'  - 2003.08.11: Improved 'Scan for redundant disposal modes'.
'  - 2003.08.13: Minor code changes in 'mGIFOptimizer' module.
'  - 2003.08.14: Improved 'Scan for redundant disposal modes' (M.B.R.).
'                Improved 'Crop transparent images'.
'  - 2003.08.31: Fixed 'Scan for redundant disposal modes' (Yes, again).
'  - 2003.09.06: Fixed 'Unxpected block skiping'.
'  - 2003.09.14: Fixed 'Scan for redundant disposal modes' (Last one!).
'  - 2003.09.18: Improved 'Crop transparent images'.



Option Explicit

'========================================================================================
' Form main
'========================================================================================

Private Sub Form_Load()
    
    '-- App. title/version
    Me.Caption = "GO v" & App.Major & "." & App.Minor
    
    '-- Set 'frames listbox' source GIF object
    ucFramesList.SetGIFSource g_oGIF
    
    '-- Set background pattern (canvas)
    g_oPattern.SetPatternFromStdPicture LoadResPicture("PATTERN_08", vbResBitmap)
    
    '-- Initialize pattern brush (empty palette entry)
    mMisc.InitializePatternBrush
End Sub

Private Sub Form_Resize()
    
    '-- Resize listbox and canvas controls
    On Error Resume Next
    ucFramesList.Height = Me.ScaleHeight - ucInfo.Height - ucFramesList.Top - 2
    ucCanvas.Move ucCanvas.Left, ucCanvas.Top, Me.ScaleWidth - ucCanvas.Left, Me.ScaleHeight - ucInfo.Height - ucCanvas.Top - 2
    On Error GoTo 0
End Sub

Private Sub Form_Unload(Cancel As Integer)

    '-- Destroy graphic objects/ref.
    ucFramesList.SetGIFSource Nothing
    ucCanvas.DIB.Destroy
    g_oGIF.Destroy
    
    '-- Is next line necessary [?]
    Set fMain = Nothing
End Sub

'========================================================================================
' Item selected
'========================================================================================

Private Sub ucFramesList_Click()
    pvRefreshCanvas
End Sub

'========================================================================================
' Zoom
'========================================================================================

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    
    '-- Canvas zoom support
    With ucCanvas
        If (KeyCode = vbKeyAdd And .Zoom < 10) Then
            .Zoom = .Zoom + 1: .Resize: .Repaint
        End If
        If (KeyCode = vbKeySubtract And .Zoom > 1) Then
            .Zoom = .Zoom - 1: .Resize: .Repaint
        End If
    End With
End Sub

'========================================================================================
' Menus
'========================================================================================

Private Sub mnuFile_Click(Index As Integer)
    
  Dim sTmpFilename As String
  
    Select Case Index
    
        Case 0 '-- Open GIF...
        
            '-- Show open file dialog
            sTmpFilename = mDialogFile.GetFileName(g_sFilename, "GIF files (*.gif)|*.GIF", , "Open GIF", -1)
            
            If (Len(sTmpFilename)) Then
                g_sFilename = sTmpFilename
                
                '-- Load GIF...
                pvLoadGIF g_sFilename
            End If
            
        Case 1 '-- Save GIF...
            
            If (g_oGIF.FramesCount) Then
            
                '-- Show save file dialog
                sTmpFilename = mDialogFile.GetFileName(g_sFilename, "GIF files (*.gif)|*.GIF", , "Save GIF", 0)
                
                If (Len(sTmpFilename)) Then
                
                    '-- Save GIF...
                    pvSaveGIF sTmpFilename
                End If
            End If
            
        Case 3 '-- Exit
            Unload Me
    End Select
End Sub

Private Sub mnuView_Click(Index As Integer)
    
    Select Case Index
    
        Case 0 '-- GIF details
            If (g_oGIF.FramesCount) Then
                fDetails.Show vbModal, Me
            End If
        
        Case 1 '-- Preview animation
            If (g_oGIF.FramesCount) Then
                fPreview.Show vbModal, Me
            End If
    End Select
End Sub

Private Sub mnuTools_Click(Index As Integer)

    Select Case Index
    
        Case 0 '-- Optimize GIF...
            If (g_oGIF.FramesCount) Then
                fOptimize.Show vbModal, Me
                pvInitialize
            End If
            
        Case 1 '-- Reduce color depth...
            If (g_oGIF.FramesCount) Then
                fRemap.Show vbModal, Me
                pvInitialize
            End If
            
        Case 3 '-- Restore frames
            If (g_oGIF.FramesCount) Then
                MsgBox "Sorry, not implemented.", vbInformation, "Restore frames"
            End If
    End Select
End Sub

Private Sub mnuHelp_Click(Index As Integer)

    Select Case Index
        
        Case 0 '-- About
        
            MsgBox vbCrLf & _
                   "GO v" & App.Major & "." & App.Minor & " (rev. " & App.Revision & ")" & vbCrLf & _
                   "GIF file size optimizer" & vbCrLf & _
                   "Carles P.V. - © 2003" & vbCrLf & vbCrLf & _
                   "- GIF optimization based on TGIFImage by Anders Melander" & vbCrLf & _
                   "- GIF decoding based on work by Vlad Vissoultchev" & vbCrLf & _
                   "- GIF encoding based on work by Ron van Tilburg", , "About"
    End Select
End Sub

'========================================================================================
' Private
'========================================================================================

Private Sub pvLoadGIF(ByVal Filename As String)

    DoEvents
    Screen.MousePointer = vbHourglass
    
    '-- Load from path
    If (g_oGIF.LoadFromFile(Filename)) Then
        Screen.MousePointer = vbDefault
        pvInitialize
      Else
        Screen.MousePointer = vbDefault
        MsgBox "Unexpected error loading GIF file.", vbExclamation, "Open GIF"
        pvCleanUp
    End If
End Sub

Private Sub pvSaveGIF(ByVal Filename As String)

    DoEvents
    Screen.MousePointer = vbHourglass
    
    '-- Save to path
    If (g_oGIF.Save(Filename) = 0) Then
        Screen.MousePointer = vbDefault
        MsgBox "Unexpected error saving GIF file.", vbExclamation, "Save GIF"
      Else
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub ucframeslist_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    
  Const vbDropFilesFromExplorer = &H7
   
    '-- Load file from passed path collection (only first)
    If (Effect = vbDropFilesFromExplorer) Then
        Call pvLoadGIF(Data.Files(1))
    End If
End Sub

'//

Private Sub pvInitialize()

  Dim nFrm As Integer
  
    '-- Show animation info
    pvShowAnimationInfo

    '-- Create canvas DIB (GIF screen size)
    ucCanvas.DIB.Create g_oGIF.ScreenWidth, g_oGIF.ScreenHeight, [32_bpp]
    ucCanvas.Resize
    
    '-- Fill frames list and select first frame
    ucFramesList.Clear
    For nFrm = 1 To g_oGIF.FramesCount
        ucFramesList.AddItem Format(nFrm, "# ")
    Next nFrm
    ucFramesList.ListIndex = 0
End Sub

Private Sub pvCleanUp()
    
    '-- Destroy current GIF object
    g_oGIF.Destroy
    
    '-- Clear frames listbox
    ucFramesList.Clear
    ucFramesList.Refresh
    
    '-- Clear canvas
    ucCanvas.DIB.Destroy
    ucCanvas.Resize
    ucCanvas.Repaint

    '-- Clear statusbar
    ucInfo.TextFile = ""
    ucInfo.TextInfo = ""
    ucInfo.Refresh
End Sub

'//

Public Sub pvShowAnimationInfo()
    
    With g_oGIF
        '-- Show animation props.
        With ucInfo
            .TextFile = g_sFilename
            .TextInfo = g_oGIF.ScreenWidth & "x" & g_oGIF.ScreenHeight & " [" & g_oGIF.FramesCount & " frames]"
            .Refresh
        End With
    End With
End Sub

Private Sub pvRefreshCanvas()
    
    With g_oGIF
    
        '-- Current selected frame
        g_nFrame = ucFramesList.ListIndex + 1
        
        '-- Paint background pattern (~ transparent layer)
        g_oPattern.Tile ucCanvas.DIB.hDC, 0, 0, .ScreenWidth, .ScreenHeight
        
        '-- Paint current frame
        .FrameDraw ucCanvas.DIB.hDC, g_nFrame
        
        '-- Draw a simple focus rectangle around current selected frame
        If (.FramesCount > 1) Then
            mMisc.DrawFocus ucCanvas.DIB.hDC, .FrameLeft(g_nFrame), .FrameTop(g_nFrame), .FrameDIBXOR(g_nFrame).Width, .FrameDIBXOR(g_nFrame).Height
        End If
    End With
    
    '-- Refreah canvas
    ucCanvas.Repaint
End Sub
