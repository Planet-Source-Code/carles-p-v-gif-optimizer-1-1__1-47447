Attribute VB_Name = "mGIFOptimizer"
'================================================
' Module:        mGIFOptimizer.bas
' Author:        Carles P.V.
' Dependencies:  cGIF.cls
' Last revision: 2003.09.18
'================================================

Option Explicit

'-- API:

Private Type RECT2
    x1 As Long
    y1 As Long
    x2 As Long
    y2 As Long
End Type

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

Private Declare Function SetRect Lib "user32" (lpRect As RECT2, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function SetRectEmpty Lib "user32" (lpRect As RECT2) As Long
Private Declare Function CopyRect Lib "user32" (lpDestRect As RECT2, lpSourceRect As RECT2) As Long
Private Declare Function OffsetRect Lib "user32" (lpRect As RECT2, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function IntersectRect Lib "user32" (lpDestRect As RECT2, lpSrc1Rect As RECT2, lpSrc2Rect As RECT2) As Long
Private Declare Function EqualRect Lib "user32" (lpRect1 As RECT2, lpRect2 As RECT2) As Long
Private Declare Function IsRectEmpty Lib "user32" (lpRect As RECT2) As Long
Private Declare Function VarPtrArray Lib "msvbvm50" Alias "VarPtr" (Ptr() As Any) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDst As Any, lpSrc As Any, ByVal Length As Long)
Private Declare Sub FillMemory Lib "kernel32" Alias "RtlFillMemory" (lpDst As Any, ByVal Length As Long, ByVal Fill As Byte)

'//

Public Enum eFrameOptimizationMethod
    [fomMinimumBoundingRectangle] = 0
    [fomFrameDifferencing]
End Enum

'//

Public Function CheckGIF(oGIF As cGIF, oProgress As ucProgress) As Boolean
 
  Dim nFrm         As Integer
  Dim tSACurr      As SAFEARRAY2D
  Dim aBitsCurr()  As Byte
  
  Dim bIsTrns      As Boolean
  Dim aTrnsIdxCurr As Byte
  Dim aMaxIdx      As Byte
                    
  Dim rRectCurr    As RECT2
  Dim rRectNext    As RECT2
  Dim rMerge       As RECT2
  
  Dim x As Long, y As Long
  Dim W As Long, H As Long
  
    With oGIF
    
        oProgress.Max = 2 * .FramesCount
    
        '== Scan for transparent images without transparent pixels:
                
        For nFrm = 1 To .FramesCount
        
            oProgress = nFrm
            
            If (.FrameUseTransparentColor(nFrm)) Then
                
                '-- Map current 8-bpp DIB bits
                pvBuild_08bppSA tSACurr, .FrameDIBXOR(nFrm)
                CopyMemory ByVal VarPtrArray(aBitsCurr()), VarPtr(tSACurr), 4
                
                '-- Get frame transparent index and dimensions
                aTrnsIdxCurr = .FrameTransparentColorIndex(nFrm)
                W = .FrameDIBXOR(nFrm).Width - 1
                H = .FrameDIBXOR(nFrm).Height - 1
                
                '-- Scan for a not-transparent pixel
                bIsTrns = 0
                For y = 0 To H
                    For x = 0 To W
                        If (aBitsCurr(x, y) = aTrnsIdxCurr) Then bIsTrns = -1: GoTo Scan1Done
                    Next x
                Next y
            
Scan1Done:      '-- Unmap DIB bits
                CopyMemory ByVal VarPtrArray(aBitsCurr()), 0&, 4
                
                '-- No transparent pixel/s:
                '   Turn off transparency and reset index
                If (bIsTrns = 0) Then
                    .FrameUseTransparentColor(nFrm) = 0
                    .FrameTransparentColorIndex(nFrm) = 0
                End If
                
              Else
                '-- Reset index
                .FrameTransparentColorIndex(nFrm) = 0
            End If
        Next nFrm
    
        '== Scan for redundant disposal modes:
       
        For nFrm = 2 To .FramesCount

            oProgress = oProgress.Max + nFrm

            '-- Get current (n-1) and next frame (n) rects.
            SetRect rRectCurr, .FrameLeft(nFrm - 1), .FrameTop(nFrm - 1), .FrameLeft(nFrm - 1) + .FrameDIBXOR(nFrm - 1).Width, .FrameTop(nFrm - 1) + .FrameDIBXOR(nFrm - 1).Height
            SetRect rRectNext, .FrameLeft(nFrm), .FrameTop(nFrm), .FrameLeft(nFrm) + .FrameDIBXOR(nFrm).Width, .FrameTop(nFrm) + .FrameDIBXOR(nFrm).Height

            '-- If current frame is completely covered by next frame...
            If (.FrameDisposalMethod(nFrm - 1) = [dmRestoreToBackground] Or .FrameDisposalMethod(nFrm - 1) = [dmRestoreToPrevious]) Then
                If (.FrameUseTransparentColor(nFrm) = 0) Then
                    If (IntersectRect(rMerge, rRectCurr, rRectNext)) Then
                        If (EqualRect(rMerge, rRectCurr)) Then
                            .FrameDisposalMethod(nFrm - 1) = [dmNotSpecified]
                        End If
                    End If
                End If
            End If
        Next nFrm: .FrameDisposalMethod(.FramesCount) = [dmNotSpecified]
        
        '== Scan for 'out of palette bounds':
        
        If (.GlobalPaletteExists) Then
            If (.ScreenBackgroundColorIndex > .GlobalPaletteEntries - 1) Then
                .GlobalPaletteEntries = .ScreenBackgroundColorIndex + 1
            End If
        End If
        
        If (.GlobalPaletteExists) Then
            aMaxIdx = .GlobalPaletteEntries - 1
            For nFrm = 1 To .FramesCount
                If (Not .LocalPaletteUsed(nFrm)) Then
                    aMaxIdx = .FrameTransparentColorIndex(nFrm)
                End If
            Next nFrm
            If (aMaxIdx > .GlobalPaletteEntries - 1) Then
                .GlobalPaletteEntries = aMaxIdx + 1
                If (Not .LocalPaletteUsed(nFrm)) Then
                    .LocalPaletteEntries(nFrm) = .GlobalPaletteEntries
                End If
           End If
        End If
        
        For nFrm = 1 To .FramesCount
            If (.LocalPaletteUsed(nFrm)) Then
                If (.FrameUseTransparentColor(nFrm) And .FrameTransparentColorIndex(nFrm) > .LocalPaletteEntries(nFrm) - 1) Then
                    .LocalPaletteEntries(nFrm) = .FrameTransparentColorIndex(nFrm) + 1
                End If
            End If
        Next nFrm
    End With
    
    '-- End
    oProgress = 0
    CheckGIF = -1
End Function

Public Function OptimizeGlobalPalette(oGIF As cGIF, oProgress As ucProgress) As Byte

  Dim aPal(1023)  As Byte
  Dim bUseGlobal  As Boolean
  
  Dim nBefore     As Integer
  Dim aNPal(1023) As Byte
  Dim aGPal(1023) As Byte
  Dim bUsed()     As Boolean
  Dim aTrnEnt()   As Byte
  Dim aInvEnt()   As Byte
  Dim lIdx        As Long
  Dim lMax        As Long
  
  Dim nFrm        As Integer
  Dim tSA08       As SAFEARRAY2D
  Dim aBits08()   As Byte
  
  Dim x As Long, y As Long
  Dim W As Long, H As Long
    
    With oGIF
        
        oProgress.Max = 2 * .FramesCount
        
        If (.GlobalPaletteExists) Then
        
            '-- Check if at least one frame uses Global
            bUseGlobal = 0
            For nFrm = 1 To .FramesCount
                If (Not .LocalPaletteUsed(nFrm)) Then bUseGlobal = -1
            Next nFrm
            '-- Global not used, destroy it
            If (Not bUseGlobal) Then
                .GlobalPaletteExists = 0
                .GlobalPaletteEntries = 0
                CopyMemory ByVal .lpGlobalPalette, aPal(0), 1024
            End If
        End If
        
        If (.GlobalPaletteExists) Then
        
            '-- Store current number of entries and initialize arrays
            nBefore = .GlobalPaletteEntries
            ReDim bUsed(nBefore - 1)
            ReDim aTrnEnt(nBefore - 1)
            ReDim aInvEnt(nBefore - 1)
            
            '-- Store Global palette as RGBQUAD byte array
            CopyMemory aGPal(0), ByVal .lpGlobalPalette, 1024
            
            '-- Check all used entries
            For nFrm = 1 To .FramesCount
                
                If (Not .LocalPaletteUsed(nFrm)) Then
                
                    oProgress = nFrm
                
                    '-- Get dimensions
                    W = .FrameDIBXOR(nFrm).Width - 1
                    H = .FrameDIBXOR(nFrm).Height - 1
                
                    '-- Map current 8-bpp DIB bits
                    pvBuild_08bppSA tSA08, .FrameDIBXOR(nFrm)
                    CopyMemory ByVal VarPtrArray(aBits08()), VarPtr(tSA08), 4
                    
                    '-- Check used entries...
                    For y = 0 To H
                        For x = 0 To W
                            bUsed(aBits08(x, y)) = -1
                        Next x
                    Next y
                    
                    '-- Unmap DIB bits
                    CopyMemory ByVal VarPtrArray(aBits08()), 0&, 4
                End If
            Next nFrm
        
            '-- 'Strecth' palette...
            For lIdx = 0 To .GlobalPaletteEntries - 1
                If (bUsed(lIdx)) Then
                    aTrnEnt(lMax) = lIdx ' New index
                    aInvEnt(lIdx) = lMax ' Inverse index
                    lMax = lMax + 1      ' Current count
                End If
            Next lIdx
            
            '-- Any entry removed [?]
            If (lMax < .GlobalPaletteEntries) Then
        
                If (lMax > 0) Then lMax = lMax - 1
                '-- Build temp. palette with only used entries
                For lIdx = 0 To lMax
                    aNPal(4 * lIdx + 0) = aGPal(4 * aTrnEnt(lIdx) + 0)
                    aNPal(4 * lIdx + 1) = aGPal(4 * aTrnEnt(lIdx) + 1)
                    aNPal(4 * lIdx + 2) = aGPal(4 * aTrnEnt(lIdx) + 2)
                Next lIdx
                
                '-- Set as new Global palette
                CopyMemory ByVal .lpGlobalPalette, aNPal(0), 1024
                .GlobalPaletteEntries = lMax + 1
                
                '-- Update Background color index
                .ScreenBackgroundColorIndex = aInvEnt(.ScreenBackgroundColorIndex)
                
                '-- Set as XOR DIBs palette (and update indexes)
                For nFrm = 1 To .FramesCount
                    
                    If (Not .LocalPaletteUsed(nFrm)) Then
                    
                        oProgress = .FramesCount + nFrm
                        
                        '-- Set new frame DIB palette
                        .FrameDIBXOR(nFrm).SetPalette aNPal()
                        '-- Store temp. copy
                        CopyMemory ByVal .lpLocalPalette(nFrm), aNPal(0), 1024
                        .LocalPaletteEntries(nFrm) = .GlobalPaletteEntries
                        
                        '-- Get dimensions
                        W = .FrameDIBXOR(nFrm).Width - 1
                        H = .FrameDIBXOR(nFrm).Height - 1
                        
                        '-- Map current 8-bpp DIB bits
                        pvBuild_08bppSA tSA08, .FrameDIBXOR(nFrm)
                        CopyMemory ByVal VarPtrArray(aBits08()), VarPtr(tSA08), 4

                        '-- Update indexes...
                        For y = 0 To H
                            For x = 0 To W
                                aBits08(x, y) = aInvEnt(aBits08(x, y))
                            Next x
                        Next y

                        '-- Unmap DIB bits
                        CopyMemory ByVal VarPtrArray(aBits08()), 0&, 4
                         
                        '-- Update transparent index
                        .FrameTransparentColorIndex(nFrm) = aInvEnt(.FrameTransparentColorIndex(nFrm))
                    End If
                Next nFrm
            End If
            oProgress = 0

            '-- Return removed entries
            OptimizeGlobalPalette = (nBefore - .GlobalPaletteEntries)
        End If
    End With
End Function

Public Function OptimizeLocalPalette(oGIF As cGIF, ByVal nFrame As Integer) As Byte

  Dim nBefore     As Integer
  Dim aNPal(1023) As Byte
  Dim aLPal(1023) As Byte
  Dim bUsed()     As Boolean
  Dim aTrnEnt()   As Byte
  Dim aInvEnt()   As Byte
  Dim lIdx        As Long
  Dim lMax        As Long
  
  Dim tSA08       As SAFEARRAY2D
  Dim aBits08()   As Byte
  
  Dim x As Long, y As Long
  Dim W As Long, H As Long
    
    With oGIF
        
        If (.LocalPaletteUsed(nFrame)) Then
        
            '-- Store current number of entries and initialize arrays
            nBefore = .LocalPaletteEntries(nFrame)
            ReDim bUsed(nBefore - 1)
            ReDim aTrnEnt(nBefore - 1)
            ReDim aInvEnt(nBefore - 1)
            
            '-- Store Local palette as RGBQUAD byte array
            CopyMemory aLPal(0), ByVal .lpLocalPalette(nFrame), 1024
            
            '-- Get dimensions
            W = .FrameDIBXOR(nFrame).Width - 1
            H = .FrameDIBXOR(nFrame).Height - 1
            
            '-- Map current 8-bpp DIB bits
            pvBuild_08bppSA tSA08, .FrameDIBXOR(nFrame)
            CopyMemory ByVal VarPtrArray(aBits08()), VarPtr(tSA08), 4
            
            '-- Check used entries...
            For y = 0 To H
                For x = 0 To W
                    bUsed(aBits08(x, y)) = -1
                Next x
            Next y
            
            '-- 'Strecth' palette...
            For lIdx = 0 To .LocalPaletteEntries(nFrame) - 1
                If (bUsed(lIdx)) Then
                    aTrnEnt(lMax) = lIdx ' New index
                    aInvEnt(lIdx) = lMax ' Inverse index
                    lMax = lMax + 1      ' Current count
                End If
            Next lIdx
            
            '-- Any entry removed [?]
            If (lMax < .LocalPaletteEntries(nFrame)) Then
        
                If (lMax > 0) Then lMax = lMax - 1
                '-- Build temp. palette with only used entries
                For lIdx = 0 To lMax
                    aNPal(4 * lIdx + 0) = aLPal(4 * aTrnEnt(lIdx) + 0)
                    aNPal(4 * lIdx + 1) = aLPal(4 * aTrnEnt(lIdx) + 1)
                    aNPal(4 * lIdx + 2) = aLPal(4 * aTrnEnt(lIdx) + 2)
                Next lIdx
                
                '-- Set as XOR DIB palette
                .FrameDIBXOR(nFrame).SetPalette aNPal()
                        
                '-- Store temp. copy
                CopyMemory ByVal .lpLocalPalette(nFrame), aNPal(0), 1024
                .LocalPaletteEntries(nFrame) = lMax + 1
            
                '-- Update indexes...
                For y = 0 To H
                    For x = 0 To W
                        aBits08(x, y) = aInvEnt(aBits08(x, y))
                    Next x
                Next y
                 
                '-- Update transparent index
                .FrameTransparentColorIndex(nFrame) = aInvEnt(.FrameTransparentColorIndex(nFrame))
            End If
            
            '-- Unmap DIB bits
            CopyMemory ByVal VarPtrArray(aBits08()), 0&, 4
            
            '-- Return removed entries
            OptimizeLocalPalette = (nBefore - .LocalPaletteEntries(nFrame))
        End If
    End With
End Function

Public Function RemoveLocalPalettes(oGIF As cGIF, oProgress As ucProgress) As Integer

  Dim nLocalsOut     As Integer
  
  Dim lGPal(255)     As Long
  Dim aGPal(1023)    As Byte
  Dim lLPal(255)     As Long
  
  Dim nGMaxEnt       As Integer
  Dim lGEnt          As Long
  Dim lLEnt          As Long
  Dim bLEntExists()  As Boolean
  Dim bGEntChecked() As Boolean
  Dim nLCount        As Integer
  Dim aLTrnEnt()     As Byte
  
  Dim nTrnsIdx       As Integer
  Dim bTrnsExists    As Boolean
  
  Dim nFrm           As Integer
  Dim tSA08          As SAFEARRAY2D
  Dim aBits08()      As Byte
  
  Dim x As Long, y As Long
  Dim W As Long, H As Long
    
    With oGIF
        
        oProgress.Max = .FramesCount
        
        '-- Check if Global palette exists, store temp. copy and get max. index
        If (.GlobalPaletteExists) Then
            CopyMemory lGPal(0), ByVal .lpGlobalPalette, 1024
            nGMaxEnt = .GlobalPaletteEntries
          Else
            nGMaxEnt = 0
        End If
        
        For nFrm = 1 To .FramesCount
        
            oProgress = nFrm
            
            '-- Local palette used [?]
            If (.LocalPaletteUsed(nFrm)) Then
                
                '-- Store temp. copy (Long format)
                CopyMemory lLPal(0), ByVal .lpLocalPalette(nFrm), 1024
                
                '-- Redim arrays
                ReDim bLEntExists(.LocalPaletteEntries(nFrm) - 1)
                ReDim aLTrnEnt(.LocalPaletteEntries(nFrm) - 1)
                ReDim bGEntChecked(nGMaxEnt)
                nLCount = 0
                
                '-- Transparent index/flag
                nTrnsIdx = IIf(.FrameUseTransparentColor(nFrm), .FrameTransparentColorIndex(nFrm), -1)
                
                '-- Check which Local colors exist in Global palette
                If (nGMaxEnt > 0) Then
                    For lLEnt = 0 To .LocalPaletteEntries(nFrm) - 1
                        For lGEnt = 0 To nGMaxEnt - 1
                            If (lLPal(lLEnt) = lGPal(lGEnt) And bGEntChecked(lGEnt) = 0) Then
                                bGEntChecked(lGEnt) = -1 ' Global entry checked!
                                bLEntExists(lLEnt) = -1  ' Color in Global
                                aLTrnEnt(lLEnt) = lGEnt  ' Translated index
                                nLCount = nLCount + 1
                                Exit For
                            End If
                        Next lGEnt
                    Next lLEnt
                End If
                
                '-- Check if we can store all Local entries in Global palette
                If (.LocalPaletteEntries(nFrm) - nLCount <= 256 - nGMaxEnt) Then
                
                    '-- Remove this Local palette
                    .LocalPaletteUsed(nFrm) = 0: nLocalsOut = nLocalsOut + 1
                    
                    '-- Add colors to Global...
                    For lLEnt = 0 To .LocalPaletteEntries(nFrm) - 1
                        If (Not bLEntExists(lLEnt)) Then
                            lGPal(nGMaxEnt) = lLPal(lLEnt)
                            aLTrnEnt(lLEnt) = nGMaxEnt
                            nGMaxEnt = nGMaxEnt + 1
                        End If
                    Next lLEnt
                    If (nGMaxEnt > 256) Then nGMaxEnt = 256
                    
                    '-- Get dimensions
                    W = .FrameDIBXOR(nFrm).Width - 1
                    H = .FrameDIBXOR(nFrm).Height - 1
                    
                    '-- Map current 8-bpp DIB bits
                    pvBuild_08bppSA tSA08, .FrameDIBXOR(nFrm)
                    CopyMemory ByVal VarPtrArray(aBits08()), VarPtr(tSA08), 4

                    '-- Update indexes...
                    For y = 0 To H
                        For x = 0 To W
                            aBits08(x, y) = aLTrnEnt(aBits08(x, y))
                        Next x
                    Next y

                    '-- Unmap DIB bits
                    CopyMemory ByVal VarPtrArray(aBits08()), 0&, 4
                    
                    '-- Update transparent index
                    .FrameTransparentColorIndex(nFrm) = aLTrnEnt(.FrameTransparentColorIndex(nFrm))
                        
                    '-- Update Global
                    .GlobalPaletteEntries = nGMaxEnt
                    .GlobalPaletteExists = -1
                    CopyMemory ByVal .lpGlobalPalette, lGPal(0), 1024
               End If
            End If
        Next nFrm
        
        '-- Translate new Global palette to a byte array
        CopyMemory aGPal(0), lGPal(0), 1024
        
        '-- Update image palettes (using new Global)
        If (nLocalsOut > 0) Then
            For nFrm = 1 To .FramesCount
               If (.LocalPaletteUsed(nFrm) = 0) Then
                   .LocalPaletteEntries(nFrm) = .GlobalPaletteEntries
                   .FrameDIBXOR(nFrm).SetPalette aGPal()
                   CopyMemory ByVal .lpLocalPalette(nFrm), lGPal(0), 1024
               End If
            Next nFrm
        End If
    End With
    
    '-- Return number of removed Locals
    oProgress = 0
    RemoveLocalPalettes = nLocalsOut
End Function

Public Function OptimizeFrames(oGIF As cGIF, ByVal OptimizationMethod As eFrameOptimizationMethod, oProgress As ucProgress) As Boolean
 
  Dim tSAPrev       As SAFEARRAY2D
  Dim tSACurr       As SAFEARRAY2D
  Dim aBitsPrev()   As Byte
  Dim aBitsCurr()   As Byte
  
  Dim lPalPrev(255) As Long
  Dim lPalCurr(255) As Long
  Dim aPalXOR(1023) As Byte
  
  Dim bTrnsAdded    As Boolean
  Dim aTrnsAddedIdx As Byte
  Dim nTrnsIdxPrev  As Integer
  Dim nTrnsIdxCurr  As Integer
  Dim bIsTrnsCurr   As Boolean
  
  Dim xOffPrev      As Long
  Dim yOffPrev      As Long
  Dim xOffCurr      As Long
  Dim yOffCurr      As Long
  Dim xLUVPrev      As Long
  Dim yLUVPrev      As Long
  Dim xLUVCurr      As Long
  Dim yLUVCurr      As Long
                    
  Dim nFrm          As Integer
  Dim nFrmPal       As Integer
  Dim nRct          As Integer
  Dim rInit()       As RECT2
  Dim rCrop()       As RECT2
  Dim bCrop()       As Boolean
  Dim rMerge        As RECT2
  
  Dim x As Long, xi As Long, xf As Long
  Dim y As Long, yi As Long, yf As Long
  
    With oGIF
    
        '== Initialize
        
        '-- Redim. Crop rectangles arrays
        ReDim rInit(1 To .FramesCount)
        ReDim rCrop(1 To .FramesCount)
        ReDim bCrop(1 To .FramesCount)
        '-- Set max. prog.
        oProgress.Max = .FramesCount
        
        '== Define initial frame rects.
        
        For nFrm = 1 To .FramesCount
            SetRect rInit(nFrm), .FrameLeft(nFrm), .FrameTop(nFrm), .FrameLeft(nFrm) + .FrameDIBXOR(nFrm).Width, .FrameTop(nFrm) + .FrameDIBXOR(nFrm).Height
        Next nFrm
        
        '== Prepare transparency (if necessary)
        
        If (OptimizationMethod = [fomFrameDifferencing]) Then
            
            For nFrm = 2 To .FramesCount
                
                '-- Current frame will 'remain' on screen
                If (.FrameDisposalMethod(nFrm - 1) <> [dmRestoreToBackground] And .FrameDisposalMethod(nFrm - 1) <> [dmRestoreToPrevious]) Then
                    
                    '-- Only scan common (intersected) rectangle
                    If (IntersectRect(rMerge, rInit(nFrm), rInit(nFrm - 1))) Then
                
                        '-- Try to get a transparent entry...
                        If (.FrameUseTransparentColor(nFrm) = 0) Then
                            
                            '-- Frame has Local palette
                            If (.LocalPaletteUsed(nFrm)) Then
                                
                                '-- Is there space for an extra entry
                                If (.LocalPaletteEntries(nFrm) < 256) Then
                                    
                                    '-- Add to palette a blank entry
                                    .LocalPaletteEntries(nFrm) = .LocalPaletteEntries(nFrm) + 1
                                    '-- Use it as transparent entry
                                    .FrameTransparentColorIndex(nFrm) = .LocalPaletteEntries(nFrm) - 1
                                    '-- And 'turn on' transparency
                                    .FrameUseTransparentColor(nFrm) = -1
                                    
                                    '-- Restore/Set palette (un-mask frame)
                                    CopyMemory aPalXOR(0), ByVal .lpLocalPalette(nFrm), 1024
                                    .FrameDIBXOR(nFrm).SetPalette aPalXOR()
                                End If
                                
                            '-- Frame uses Global palette
                            Else
                            
                                '-- Is there space for an extra entry, or does already
                                '   exists (previously added) one?
                                If (.GlobalPaletteEntries() < 256 Or bTrnsAdded) Then
                                    
                                    If (bTrnsAdded) Then
                                    
                                        '-- A transparent entry already added. Use it
                                        .FrameTransparentColorIndex(nFrm) = aTrnsAddedIdx
                                        
                                      Else
                                        '-- No transparent entry. Add one...
                                        .GlobalPaletteEntries = .GlobalPaletteEntries + 1
                                        
                                        '-- This is the Global palette: update all temp. palette copies
                                        For nFrmPal = 1 To .FramesCount
                                            '-- Update number of entries
                                            If (.LocalPaletteUsed(nFrmPal) = 0) Then .LocalPaletteEntries(nFrmPal) = .GlobalPaletteEntries
                                            '-- Restore/Set palette (un-mask frame)
                                            CopyMemory aPalXOR(0), ByVal .lpGlobalPalette, 1024
                                            .FrameDIBXOR(nFrm).SetPalette aPalXOR()
                                        Next nFrmPal
                                        '-- Set new transparent entry
                                        .FrameTransparentColorIndex(nFrm) = .GlobalPaletteEntries - 1
                                        
                                        '-- Store this index. Enable 'bTrnsAdded' flag
                                        aTrnsAddedIdx = .GlobalPaletteEntries - 1
                                        bTrnsAdded = -1
                                    End If
                                    
                                    '-- 'Turn on' transparency
                                    .FrameUseTransparentColor(nFrm) = -1
                                End If
                            End If
                        End If
                    End If
                End If
            Next nFrm
        End If
        
        '== Start removing redundant pixels...
        
        '-- From last to first to avoid interfering
        For nFrm = .FramesCount - 1 To 1 Step -1
            
            '-- Current progress
            oProgress = .FramesCount - nFrm
            
            '-- Get current intersection rectangle
            IntersectRect rMerge, rInit(nFrm), rInit(nFrm + 1)
            
            '-- Initialize initial Crop rectangle
            Select Case OptimizationMethod
                Case [fomMinimumBoundingRectangle]: rCrop(nFrm + 1) = rInit(nFrm + 1)
                Case [fomFrameDifferencing]:        rCrop(nFrm + 1) = rMerge
            End Select
                
            '-- Current frame will 'remain' on screen
            If (.FrameDisposalMethod(nFrm) <> [dmRestoreToBackground] And .FrameDisposalMethod(nFrm) <> [dmRestoreToPrevious]) Then
                
                '-- Only scan common (intersected) rectangle
                If (IntersectRect(rMerge, rInit(nFrm), rInit(nFrm + 1))) Then
                    
                    '-- Map frame bits
                    pvBuild_08bppSA tSAPrev, .FrameDIBXOR(nFrm)
                    CopyMemory ByVal VarPtrArray(aBitsPrev()), VarPtr(tSAPrev), 4
                    pvBuild_08bppSA tSACurr, .FrameDIBXOR(nFrm + 1)
                    CopyMemory ByVal VarPtrArray(aBitsCurr()), VarPtr(tSACurr), 4
                    
                    '-- Store temp. copy of palettes (RGB Long format)
                    CopyMemory lPalPrev(0), ByVal .lpLocalPalette(nFrm), 1024
                    CopyMemory lPalCurr(0), ByVal .lpLocalPalette(nFrm + 1), 1024
            
                    '-- Get scan bounds and offsets, and transparent index/flag of both frame
                    xi = rMerge.x1
                    xf = rMerge.x2 - 1
                    yi = rMerge.y1
                    yf = rMerge.y2 - 1
                    xOffPrev = .FrameLeft(nFrm)
                    yOffPrev = .FrameTop(nFrm)
                    xOffCurr = .FrameLeft(nFrm + 1)
                    yOffCurr = .FrameTop(nFrm + 1)
                    nTrnsIdxPrev = IIf(.FrameUseTransparentColor(nFrm), .FrameTransparentColorIndex(nFrm), -1)
                    nTrnsIdxCurr = IIf(.FrameUseTransparentColor(nFrm + 1), .FrameTransparentColorIndex(nFrm + 1), -1)
                    
                    '-- Frame differencing
                    If (OptimizationMethod = [fomFrameDifferencing]) Then
                        
                        '-- Current frame has transparent entry. Can be processed:
                        If (.FrameUseTransparentColor(nFrm + 1)) Then
                            '-- Start scan...
                            For y = yi To yf
                                yLUVPrev = y - yOffPrev
                                yLUVCurr = y - yOffCurr
                                For x = xi To xf
                                    xLUVPrev = x - xOffPrev
                                    xLUVCurr = x - xOffCurr
                                    If (aBitsPrev(xLUVPrev, yLUVPrev) <> nTrnsIdxPrev) Then
                                        If (lPalCurr(aBitsCurr(xLUVCurr, yLUVCurr)) = lPalPrev(aBitsPrev(xLUVPrev, yLUVPrev))) Then aBitsCurr(xLUVCurr, yLUVCurr) = nTrnsIdxCurr
                                    End If
                                Next x
                            Next y
                        End If
                        
                    '-- Minimum Bounding Rectangle
                    Else
                        
                        '-- Previous frame contains current frame:
                        If (EqualRect(rMerge, rInit(nFrm + 1))) Then
                            
                            '-- Top:
                            For y = yi To yf
                                yLUVPrev = y - yOffPrev
                                yLUVCurr = y - yOffCurr
                                For x = xi To xf
                                    xLUVPrev = x - xOffPrev
                                    xLUVCurr = x - xOffCurr
                                    If ((aBitsCurr(xLUVCurr, yLUVCurr) <> nTrnsIdxCurr) Xor (aBitsPrev(xLUVPrev, yLUVPrev) <> nTrnsIdxPrev)) Or (lPalCurr(aBitsCurr(xLUVCurr, yLUVCurr)) <> lPalPrev(aBitsPrev(xLUVPrev, yLUVPrev))) Then rCrop(nFrm + 1).y1 = y: GoTo Check_y2
                                Next x
                            Next y
                            SetRectEmpty rCrop(nFrm + 1): GoTo ScanDone
                            
Check_y2:                   '-- Bottom:
                            For y = yf To yi Step -1
                                yLUVPrev = y - yOffPrev
                                yLUVCurr = y - yOffCurr
                                For x = xi To xf
                                    xLUVPrev = x - xOffPrev
                                    xLUVCurr = x - xOffCurr
                                    If ((aBitsCurr(xLUVCurr, yLUVCurr) <> nTrnsIdxCurr) Xor (aBitsPrev(xLUVPrev, yLUVPrev) <> nTrnsIdxPrev)) Or (lPalCurr(aBitsCurr(xLUVCurr, yLUVCurr)) <> lPalPrev(aBitsPrev(xLUVPrev, yLUVPrev))) Then rCrop(nFrm + 1).y2 = y + 1: GoTo Check_x1
                                Next x
                            Next y
                            SetRectEmpty rCrop(nFrm + 1): GoTo ScanDone
                            
Check_x1:                   yi = rCrop(nFrm + 1).y1
                            yf = rCrop(nFrm + 1).y2 - 1
                            
                            '-- Left:
                            For x = xi To xf
                                xLUVPrev = x - xOffPrev
                                xLUVCurr = x - xOffCurr
                                For y = yi To yf
                                    yLUVPrev = y - yOffPrev
                                    yLUVCurr = y - yOffCurr
                                    If ((aBitsCurr(xLUVCurr, yLUVCurr) <> nTrnsIdxCurr) Xor (aBitsPrev(xLUVPrev, yLUVPrev) <> nTrnsIdxPrev)) Or (lPalCurr(aBitsCurr(xLUVCurr, yLUVCurr)) <> lPalPrev(aBitsPrev(xLUVPrev, yLUVPrev))) Then rCrop(nFrm + 1).x1 = x: GoTo Check_x2
                                Next y
                            Next x
                            SetRectEmpty rCrop(nFrm + 1): GoTo ScanDone
                            
Check_x2:                   '-- Right:
                            For x = xf To xi Step -1
                                xLUVPrev = x - xOffPrev
                                xLUVCurr = x - xOffCurr
                                For y = yi To yf
                                    yLUVPrev = y - yOffPrev
                                    yLUVCurr = y - yOffCurr
                                    If ((aBitsCurr(xLUVCurr, yLUVCurr) <> nTrnsIdxCurr) Xor (aBitsPrev(xLUVPrev, yLUVPrev) <> nTrnsIdxPrev)) Or (lPalCurr(aBitsCurr(xLUVCurr, yLUVCurr)) <> lPalPrev(aBitsPrev(xLUVPrev, yLUVPrev))) Then rCrop(nFrm + 1).x2 = x + 1: GoTo ScanDone
                                Next y
                            Next x
                            SetRectEmpty rCrop(nFrm + 1)
                            
ScanDone:                   '-- Scan done.
                            '-- Need to crop [?]
                            bCrop(nFrm + 1) = Not (-EqualRect(rCrop(nFrm + 1), rMerge))
                            OffsetRect rCrop(nFrm + 1), -.FrameLeft(nFrm + 1), -.FrameTop(nFrm + 1)
                        End If
                    End If
                    
                    '-- Unmap frame DIB bits
                    CopyMemory ByVal VarPtrArray(aBitsPrev()), 0&, 4
                    CopyMemory ByVal VarPtrArray(aBitsCurr()), 0&, 4
                End If
            End If
        Next nFrm
         
        '== Remove redundant frames...
        If (OptimizationMethod = [fomMinimumBoundingRectangle]) Then
            Call pvRemoveRedundantFrames(oGIF, bCrop(), rCrop())
        End If
        
        '== Crop frames...
        If (OptimizationMethod = [fomMinimumBoundingRectangle]) Then
            For nFrm = 2 To .FramesCount
                If (bCrop(nFrm)) Then Call pvCropFrame(oGIF, nFrm, rCrop(nFrm))
            Next nFrm
        End If

        '== End/Success
        oProgress = 0
        OptimizeFrames = -1
    End With
End Function

Public Function CropTransparentImages(oGIF As cGIF, oProgress As ucProgress) As Boolean

  Dim nFrm          As Integer
  Dim tSA           As SAFEARRAY2D
  Dim aBits()       As Byte
  
  Dim bIsTrnsCurr   As Boolean
  Dim aTrnsIdx      As Byte
  
  Dim nRct          As Integer
  Dim rMerge        As RECT2
  Dim rInit()       As RECT2
  Dim rCrop()       As RECT2
  Dim bCrop()       As Boolean
  
  Dim bProcess      As Boolean
  
  Dim x As Long, xi As Long, xf As Long
  Dim y As Long, yi As Long, yf As Long
    
    With oGIF
    
        '== Initialize
        '-- Redim. Crop rectangles arrays
        ReDim rInit(1 To .FramesCount)
        ReDim rCrop(1 To .FramesCount)
        ReDim bCrop(1 To .FramesCount)
        '-- Set max. prog.
        oProgress.Max = .FramesCount
        
        '== Define initial frame rects.
        For nFrm = 1 To .FramesCount
            SetRect rInit(nFrm), .FrameLeft(nFrm), .FrameTop(nFrm), .FrameLeft(nFrm) + .FrameDIBXOR(nFrm).Width, .FrameTop(nFrm) + .FrameDIBXOR(nFrm).Height
        Next nFrm
        
        '== Calc. minimum bounding rectangles
        For nFrm = 1 To .FramesCount
            
            '-- Current progress
            oProgress = nFrm
            
            '-- Initialize crop rect.
            CopyRect rCrop(nFrm), rInit(nFrm)
            
            '-- Only try to crop transparent frames
            If (.FrameUseTransparentColor(nFrm)) Then
            
                '-- Check if current frame completely covers previous frame
                bProcess = 0
                Select Case nFrm
                    Case Is = 1 '-- Always process first frame
                        bProcess = -1
                    Case Is > 1 '-- Check others
                        If (IntersectRect(rMerge, rInit(nFrm - 1), rInit(nFrm))) Then
                            If (EqualRect(rMerge, rInit(nFrm - 1))) Then
                                bProcess = -1
                            End If
                        End If
                End Select

                '-- Minimum Bounding Rectangle
                If (bProcess) Then
                
                    '-- Map frame bits
                    pvBuild_08bppSA tSA, .FrameDIBXOR(nFrm)
                    CopyMemory ByVal VarPtrArray(aBits()), VarPtr(tSA), 4
                
                    '-- Get bounds
                    xi = 0: xf = .FrameDIBXOR(nFrm).Width - 1
                    yi = 0: yf = .FrameDIBXOR(nFrm).Height - 1
                    
                    '-- Transparent color index
                    aTrnsIdx = .FrameTransparentColorIndex(nFrm)
                    
                    '-- Top:
                    For y = yi To yf
                        For x = xi To xf
                            If (aBits(x, y) <> aTrnsIdx) Then rCrop(nFrm).y1 = y: GoTo Check_y2
                        Next x
                    Next y
                    SetRectEmpty rCrop(nFrm): GoTo ScanDone
                    
Check_y2:           '-- Bottom:
                    For y = yf To yi Step -1
                        For x = xi To xf
                            If (aBits(x, y) <> aTrnsIdx) Then rCrop(nFrm).y2 = y + 1: GoTo Check_x1
                        Next x
                    Next y
                    SetRectEmpty rCrop(nFrm): GoTo ScanDone
                    
Check_x1:           yi = rCrop(nFrm).y1
                    yf = rCrop(nFrm).y2 - 1
                                
                    '-- Left:
                    For x = xi To xf
                        For y = yi To yf
                            If (aBits(x, y) <> aTrnsIdx) Then rCrop(nFrm).x1 = x: GoTo Check_x2
                        Next y
                    Next x
                    SetRectEmpty rCrop(nFrm): GoTo ScanDone
                    
Check_x2:           '-- Right:
                    For x = xf To xi Step -1
                        For y = yi To yf
                            If (aBits(x, y) <> aTrnsIdx) Then rCrop(nFrm).x2 = x + 1: GoTo ScanDone
                        Next y
                    Next x
                    SetRectEmpty rCrop(nFrm)
                    
ScanDone:           '-- Scan done
                    CopyMemory ByVal VarPtrArray(aBits()), 0&, 4
                   
                    '-- Need to crop [?]
                    OffsetRect rCrop(nFrm), .FrameLeft(nFrm), .FrameTop(nFrm)
                    bCrop(nFrm) = Not -EqualRect(rCrop(nFrm), rInit(nFrm))
                    OffsetRect rCrop(nFrm), -.FrameLeft(nFrm), -.FrameTop(nFrm)
                End If
            End If
        Next nFrm
        
        '== Remove redundant frames...
        If (.FramesCount > 1) Then
            Call pvRemoveRedundantFrames(oGIF, bCrop(), rCrop())
        End If
        
        '== Crop frames...
        For nFrm = 1 To .FramesCount
            If (bCrop(nFrm)) Then
                If (Not (IsRectEmpty(rCrop(nFrm)) <> 0 And nFrm = 1)) Then
                    Call pvCropFrame(oGIF, nFrm, rCrop(nFrm))
                End If
            End If
        Next nFrm
        
        '== End/Success
        oProgress = 0
        CropTransparentImages = -1
    End With
End Function

Public Sub RemaskFrames(oGIF As cGIF, oProgress As ucProgress)
  
  Dim nFrm As Integer
  
    oProgress.Max = oGIF.FramesCount
    
    '-- Re-mask frames
    For nFrm = 1 To oGIF.FramesCount
        oProgress = nFrm
        oGIF.FrameMask nFrm, oGIF.FrameTransparentColorIndex(nFrm)
    Next nFrm
    oProgress = 0
End Sub

Private Sub pvRemoveRedundantFrames(oGIF As cGIF, bCrop() As Boolean, rCrop() As RECT2)

  Dim nFrm As Integer
  Dim nRct As Integer

    '-- Start from second frame
    nFrm = 2
    
    With oGIF
    
        Do  '-- Null rectangle [?]
            If (IsRectEmpty(rCrop(nFrm))) Then

                '-- Add delay time of removed frame to previous one
                .FrameDelay(nFrm - 1) = .FrameDelay(nFrm - 1) + .FrameDelay(nFrm)
                '-- Remove redundant frame
                .FrameRemove nFrm

                '-- Update temp. bounding rectangles array
                For nRct = nFrm To .FramesCount
                    rCrop(nRct) = rCrop(nRct + 1)
                    bCrop(nRct) = bCrop(nRct + 1)
                Next nRct

              Else
                '-- Next frame
                nFrm = nFrm + 1
            End If
        Loop Until nFrm > .FramesCount
    End With
End Sub

Private Sub pvCropFrame(oGIF As cGIF, ByVal nFrame As Integer, rCrop As RECT2)

  Dim oDIBBuff      As New cDIB
  Dim lpBitsDst     As Long
  Dim lpBitsSrc     As Long
  Dim lWidthDst     As Long
  Dim lScanWidthDst As Long
  Dim lScanWidthSrc As Long
  Dim y             As Long
  
  Dim aPalXOR(1023) As Byte
  Dim aPalAND(7)    As Byte

    With oGIF

        '-- Prepare palettes
        CopyMemory aPalXOR(0), ByVal .lpLocalPalette(nFrame), 1024
        FillMemory aPalAND(4), 3, &HFF
        
        '-- XOR DIB
        With .FrameDIBXOR(nFrame)
            
            '-- Make a temp. copy
            .CloneTo oDIBBuff
            '-- Resize current frame
            .Create rCrop.x2 - rCrop.x1, rCrop.y2 - rCrop.y1, [08_bpp]
            .SetPalette aPalXOR()
            
            '-- Get some props. for cropping process
            lpBitsDst = .lpBits
            lpBitsSrc = oDIBBuff.lpBits
            lScanWidthDst = .BytesPerScanline
            lScanWidthSrc = oDIBBuff.BytesPerScanline
            lWidthDst = .Width
            
            '-- Crop...
            For y = rCrop.y1 To rCrop.y2 - 1
                CopyMemory ByVal lpBitsDst + lScanWidthDst * (y - rCrop.y1), ByVal lpBitsSrc + (y * lScanWidthSrc) + rCrop.x1, lWidthDst
            Next y
        End With
        
        '-- AND DIB
        With .FrameDIBAND(nFrame)
            .Create rCrop.x2 - rCrop.x1, rCrop.y2 - rCrop.y1, [01_bpp]
            .SetPalette aPalAND()
        End With
        
        '-- Update frame position
        .FrameLeft(nFrame) = .FrameLeft(nFrame) + rCrop.x1
        .FrameTop(nFrame) = .FrameTop(nFrame) + rCrop.y1
    End With
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
