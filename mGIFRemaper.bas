Attribute VB_Name = "mGIFRemaper"
'================================================
' Module:        mGIFRemaper.bas
' Author:        Carles P.V.
' Dependencies:  cGIF.cls
'                cDIB.cls
'                cPal8bpp.cls
' Last revision: 2003.08.11
'================================================

Option Explicit

'-- API:

Private Type SAFEARRAYBOUND
    cElements As Long
    lLbound   As Long
End Type

Private Type SAFEARRAY2D
    cDims      As Integer
    fFeatures  As Integer
    cbElements As Long
    cLocks     As Long
    pvData     As Long
    Bounds(1)  As SAFEARRAYBOUND
End Type

Private Declare Function VarPtrArray Lib "msvbvm50" Alias "VarPtr" (Ptr() As Any) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDst As Any, lpSrc As Any, ByVal Length As Long)

'//

'-- Private property variables:

Private m_oPalette As New cPal8bpp

'-- Private variables:

Private m_aInvIdxLUT(255) As Byte
Private m_nTrnsIdx        As Integer
Private m_aNewTrnsIdx     As Byte

'//

'========================================================================================
' Properties
'========================================================================================

Public Property Set Palette(oPalette As cPal8bpp)
    Set m_oPalette = oPalette
End Property

Public Property Get Palette() As cPal8bpp
    Set Palette = m_oPalette
End Property

'========================================================================================
' Methods
'========================================================================================

Public Function RemapGIF(oGIF As cGIF, oProgress As ucProgress) As Boolean

  Dim nFrm    As Integer
  Dim oOldPal As New cPal8bpp
    
    With oGIF
    
        oProgress.Max = .FramesCount
    
        For nFrm = 1 To .FramesCount
            
            oProgress = nFrm
            
            '-- Get current image palette
            oOldPal.Initialize .LocalPaletteEntries(nFrm)
            CopyMemory ByVal oOldPal.lpPalette, ByVal .lpLocalPalette(nFrm), 1024
            
            '-- Remap current image
            pvRemap .FrameDIBXOR(nFrm), oOldPal, m_oPalette, IIf(.FrameUseTransparentColor(nFrm), .FrameTransparentColorIndex(nFrm), -1)
            
            '-- Update transparent index
            If (.FrameUseTransparentColor(nFrm)) Then
                .FrameTransparentColorIndex(nFrm) = m_oPalette.Entries - 1
            End If
            
            '-- Remove Local palette
            .LocalPaletteUsed(nFrm) = 0
            '-- Update frame palette copy
            CopyMemory ByVal .lpLocalPalette(nFrm), ByVal m_oPalette.lpPalette, 1024
            .LocalPaletteEntries(nFrm) = m_oPalette.Entries
        Next nFrm
        
        '-- New Global
        CopyMemory ByVal .lpGlobalPalette, ByVal m_oPalette.lpPalette, 1024
        .GlobalPaletteEntries = m_oPalette.Entries
        .GlobalPaletteExists = -1
    End With
    
    oProgress = 0
End Function

Public Function GIFOptimalPalette(oGIF As cGIF, ByVal PaletteEntries As Integer, Optional ByVal PreservedColor As Long = -1) As cPal8bpp
  
  Dim oPal        As New cPal8bpp
  Dim nFrm        As Integer
  Dim lPal(1023)  As Long
  Dim lPalEnt     As Long
  Dim lColor()    As Long
  Dim lColorEnt   As Long
  Dim aOneTrnsIdx As Byte
  Dim nTrnsIdx    As Integer
  Dim oDummyDIB   As New cDIB

    With oGIF
    
        For nFrm = 1 To .FramesCount
            
            '-- Get temp. frame palette
            CopyMemory lPal(0), ByVal .lpLocalPalette(nFrm), 1024
            '-- Tranparent entry/flag
            If (.FrameUseTransparentColor(nFrm)) Then
                nTrnsIdx = .FrameTransparentColorIndex(nFrm)
                aOneTrnsIdx = 1
              Else
                nTrnsIdx = -1
            End If
            
            '-- Store space
            ReDim Preserve lColor(lColorEnt + .LocalPaletteEntries(nFrm))
            
            '-- Get not transparent entries
            For lPalEnt = 0 To .LocalPaletteEntries(nFrm) - 1
                If (lPalEnt <> nTrnsIdx) Then
                    lColor(lColorEnt) = lPal(lPalEnt)
                    lColorEnt = lColorEnt + 1
                End If
            Next lPalEnt
            ReDim Preserve lColor(lColorEnt - 1)
        Next nFrm
    End With
        
    '-- Create dummy DIB
    oDummyDIB.Create 4 * lColorEnt, 1, [32_bpp]
    
    '-- Copy all palette entries to dummy DIB
    CopyMemory ByVal oDummyDIB.lpBits, lColor(0), 4 * lColorEnt
    
    '-- Finaly, extract optimal palette and pass it (*)
    With oPal
        '-- Get palette
        .CreateOptimal oDummyDIB, PaletteEntries - -(PreservedColor > -1) - aOneTrnsIdx, 8
        '-- Preserved color [?]
        If (PreservedColor > -1) Then
            .Entries = .Entries + 1
            .rgbR(.Entries - 1) = (PreservedColor And &HFF&)
            .rgbG(.Entries - 1) = (PreservedColor And &HFF00&) \ 256
            .rgbB(.Entries - 1) = (PreservedColor And &HFF0000) \ 65536
            .BuildLogicalPalette
        End If
        '-- Transparent entry [?]
        .Entries = .Entries + aOneTrnsIdx
    End With
    Set GIFOptimalPalette = oPal
    
' (*) In case desired entries is 8, resulting final palette entries could be 9 or 10.
End Function

Public Function PaletteIndex08(oDIB08 As cDIB, ByVal x As Long, ByVal y As Long) As Byte
    
  Dim aBits() As Byte
  Dim tSA     As SAFEARRAY2D
  
    '-- Map DIB bits
    pvBuild_08bppSA tSA, oDIB08
    CopyMemory ByVal VarPtrArray(aBits()), VarPtr(tSA), 4
    
    '-- Get 8-bpp index
    PaletteIndex08 = aBits(x, y)

    '-- Unmap DIB bits
    CopyMemory ByVal VarPtrArray(aBits()), 0&, 4
End Function

'========================================================================================
' Private
'========================================================================================

Private Sub pvBuildInvIdxLUT(oPalOld As cPal8bpp, oPal As cPal8bpp, ByVal TransparentIdx As Integer)
  
  Dim lEnt As Long
    
    With oPalOld
        '-- Get nearest color index
        For lEnt = 0 To .Entries - 1
            oPal.ClosestIndex .rgbR(lEnt), .rgbG(lEnt), .rgbB(lEnt), m_aInvIdxLUT(lEnt)
        Next lEnt
    End With
    '-- Transparent frame [?]
    If (TransparentIdx <> -1) Then m_aInvIdxLUT(TransparentIdx) = oPal.Entries - 1
End Sub

Private Sub pvRemap(oDIB As cDIB, oOldPal As cPal8bpp, oPal As cPal8bpp, ByVal TransparentIdx As Integer)

  Dim tSA     As SAFEARRAY2D
  Dim aBits() As Byte

  Dim x As Long, y As Long
  Dim W As Long, H As Long
    
    '-- Build inverse index LUT
    pvBuildInvIdxLUT oOldPal, oPal, TransparentIdx
    
    '-- Map current 8-bpp DIB bits
    pvBuild_08bppSA tSA, oDIB
    CopyMemory ByVal VarPtrArray(aBits()), VarPtr(tSA), 4
   
    '-- Get dimensions
    W = oDIB.Width - 1
    H = oDIB.Height - 1
    
    '-- Update indexes
    For y = 0 To H
        For x = 0 To W
            aBits(x, y) = m_aInvIdxLUT(aBits(x, y))
        Next x
    Next y
    
    '-- Unmap DIB bits
    CopyMemory ByVal VarPtrArray(aBits()), 0&, 4
End Sub

'//

Private Sub pvBuild_08bppSA(tSA As SAFEARRAY2D, oDIB As cDIB)

    '-- 8-bpp DIB mapping
    With tSA
        .cbElements = 1
        .cDims = 2
        .Bounds(0).lLbound = 0
        .Bounds(0).cElements = oDIB.Height
        .Bounds(1).lLbound = 0
        .Bounds(1).cElements = oDIB.BytesPerScanline
        .pvData = oDIB.lpBits
    End With
End Sub
