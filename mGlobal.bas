Attribute VB_Name = "mGlobal"
Option Explicit

Private Type tGO
    OGRemoveUnusedPaletteEntries   As Boolean
    OGMoveColorsFromLocalsToGlobal As Boolean
    OGRemoveRedundantPixels        As Boolean
    OGRemoveRedundantPixelsMethod  As Byte
    OGCropTransparentImages        As Boolean
    OGDisableInterlacing           As Boolean
    RGPaletteEntriesMode           As Byte
    RGPaletteEntriesFixedBPP       As Byte
    RGPaletteEntriesCustomEntries  As Integer
    RGPreserveColor                As Boolean
    RGPreservedColor               As Long
End Type

'//

Public g_sFilename As String    ' Last file path
Public g_oGIF      As New cGIF  ' Our GIF object
Public g_nFrame    As Integer   ' Current frame

Public g_oPattern  As New cTile ' Background pattern (transparent layer)
Public g_lColor    As Long      ' Current selected color

Public g_tGO       As tGO       ' Current settings
