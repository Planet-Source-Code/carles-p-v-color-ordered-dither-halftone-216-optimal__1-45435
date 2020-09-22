Attribute VB_Name = "mDither8bpp"
'================================================
' Module:        mDither8bpp.bas (simplified)
' Author:        Carles P.V.
' Dependencies:  cDIB.cls
'                cPal8bpp.cls
' Last revision: -
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

Private Type RGBQUAD
    B As Byte
    G As Byte
    R As Byte
    A As Byte
End Type

Private Declare Function VarPtrArray Lib "msvbvm50" Alias "VarPtr" (Ptr() As Any) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDst As Any, lpSrc As Any, ByVal Length As Long)

'//

'-- Public Enums.:

Public Enum impPaletteCts
    [ipBrowser] = 0
    [ipOptimal]
End Enum

'-- Property variables:

Private m_Palette   As impPaletteCts

'-- Private variables:

Private m_tPal(255) As RGBQUAD            '  8-bpp current palette entries

Private m_tSA32     As SAFEARRAY2D        ' 32-bpp SA
Private m_Bits32()  As RGBQUAD            ' 32-bpp mapping bits
Private m_tSA08     As SAFEARRAY2D        '  8-bpp SA
Private m_Bits08()  As Byte               '  8-bpp mapping bits

Private m_x As Long
Private m_y As Long
Private m_W As Long
Private m_H As Long

'//

Private m_OD_Thr(3, 3)            As Long ' Ordered dither matrix thresholds
Private m_OD_Oiv(3, 3)            As Long ' RGB4096 (optimal palette) inc. values
Private m_OD_Hiv(3, 3)            As Long ' Halftone inc. values

Private m_RGB4096_Inv(15, 15, 15) As Byte ' RGB4096 palette inverse index LUT
Private m_RGB4096_Trn(-8 To 262)  As Long ' RGB4096 translation LUT

Private m_HT216_Inv(5, 5, 5)      As Byte ' Halftone palette inverse index LUT
Private m_HT216_Trn(-51 To 280)   As Long ' Halftone translation LUT




'========================================================================================
' Module initialization
'========================================================================================

Public Sub InitializeLUTs()

  Dim lIdx As Long
  Dim R As Long
  Dim G As Long
  Dim B As Long
      
    '-- Ordered dither matrix thresholds (Bayer) and incs.
    m_OD_Thr(0, 0) = 16:  m_OD_Oiv(0, 0) = 1:  m_OD_Hiv(0, 0) = 3
    m_OD_Thr(1, 0) = 144: m_OD_Oiv(1, 0) = 9:  m_OD_Hiv(1, 0) = 29
    m_OD_Thr(2, 0) = 48:  m_OD_Oiv(2, 0) = 3:  m_OD_Hiv(2, 0) = 10
    m_OD_Thr(3, 0) = 176: m_OD_Oiv(3, 0) = 11: m_OD_Hiv(3, 0) = 35
    
    m_OD_Thr(0, 1) = 208: m_OD_Oiv(0, 1) = 13: m_OD_Hiv(0, 1) = 41
    m_OD_Thr(1, 1) = 80:  m_OD_Oiv(1, 1) = 5:  m_OD_Hiv(1, 1) = 16
    m_OD_Thr(2, 1) = 240: m_OD_Oiv(2, 1) = 15: m_OD_Hiv(2, 1) = 48
    m_OD_Thr(3, 1) = 112: m_OD_Oiv(3, 1) = 7:  m_OD_Hiv(3, 1) = 22
    
    m_OD_Thr(0, 2) = 64:  m_OD_Oiv(0, 2) = 4:  m_OD_Hiv(0, 2) = 13
    m_OD_Thr(1, 2) = 192: m_OD_Oiv(1, 2) = 12: m_OD_Hiv(1, 2) = 38
    m_OD_Thr(2, 2) = 32:  m_OD_Oiv(2, 2) = 2:  m_OD_Hiv(2, 2) = 6
    m_OD_Thr(3, 2) = 160: m_OD_Oiv(3, 2) = 10: m_OD_Hiv(3, 2) = 32
    
    m_OD_Thr(0, 3) = 256: m_OD_Oiv(0, 3) = 16: m_OD_Hiv(0, 3) = 51
    m_OD_Thr(1, 3) = 128: m_OD_Oiv(1, 3) = 8:  m_OD_Hiv(1, 3) = 26
    m_OD_Thr(2, 3) = 224: m_OD_Oiv(2, 3) = 14: m_OD_Hiv(2, 3) = 45
    m_OD_Thr(3, 3) = 96:  m_OD_Oiv(3, 3) = 6:  m_OD_Hiv(3, 3) = 19
    
    '-- Halfote-216 inverse indexes LUT
    For B = 0 To &H100 Step &H33
        For G = 0 To &H100 Step &H33
            For R = 0 To &H100 Step &H33
                '-- Set palette inverse index
                m_HT216_Inv(R \ &H33, G \ &H33, B \ &H33) = lIdx
                lIdx = lIdx + 1
            Next R
        Next G
    Next B
    '-- Halftone-216 translation LUT
    For lIdx = -51 To 280
        m_HT216_Trn(lIdx) = lIdx / &H33
        If (m_HT216_Trn(lIdx) < 0) Then m_HT216_Trn(lIdx) = 0
        If (m_HT216_Trn(lIdx) > 5) Then m_HT216_Trn(lIdx) = 5
    Next lIdx
    
    '-- RGB-4096 translation LUT
    For lIdx = -7 To 262
        m_RGB4096_Trn(lIdx) = lIdx / 17
        If (m_RGB4096_Trn(lIdx) < 0) Then m_RGB4096_Trn(lIdx) = 0
        If (m_RGB4096_Trn(lIdx) > 15) Then m_RGB4096_Trn(lIdx) = 15
    Next lIdx
End Sub

'========================================================================================
' Properties
'========================================================================================

Public Property Get Palette() As impPaletteCts
    Palette = m_Palette
End Property
Public Property Let Palette(ByVal New_Palette As impPaletteCts)
    m_Palette = New_Palette
End Property

'========================================================================================
' Methods
'========================================================================================

Public Sub Dither(oDIB32 As cDIB, oDIB08 As cDIB)
  
  Dim oTmpDIB32  As New cDIB
  Dim oPal       As New cPal8bpp
  Dim aPal(1023) As Byte
  
  Dim bfW As Long, bfH As Long
  Dim bfx As Long, bfy As Long
  
    '// Set palette
    
    Select Case m_Palette
            
        Case [ipOptimal]
            
            '-- Get Optimal from reduced DIB (speed up)
            oTmpDIB32.CreateFromStdPicture oDIB32.Image, -1
            oTmpDIB32.GetBestFitInfo 150, 150, bfx, bfy, bfW, bfH
            oTmpDIB32.Resize bfW, bfH
            oPal.CreateOptimal oTmpDIB32, 256, 8
            '-- Build 4096-colors palette inverse indexes LUT
            pvBuildRGB4096LUT oPal

        Case [ipBrowser]
        
            '-- Create Halftne-216 (6x6x6)
            oPal.CreateHalftone [216_phLevels]
    End Select
    
    '// Fill temp. palette copy (speed up)
    CopyMemory m_tPal(0), ByVal oPal.lpPalette, 1024
    
    '// Rebuild 8-bpp target DIB (create and set current palette)
    CopyMemory aPal(0), m_tPal(0), 1024
    oDIB08.Create oDIB32.Width, oDIB32.Height, [08_bpp]
    oDIB08.SetPalette aPal()
    
    '// Map source and target DIB bits (32-bpp/8-bpp)
    pvBuild_32bppSA m_tSA32, oDIB32
    pvBuild_08bppSA m_tSA08, oDIB08
    CopyMemory ByVal VarPtrArray(m_Bits32()), VarPtr(m_tSA32), 4
    CopyMemory ByVal VarPtrArray(m_Bits08()), VarPtr(m_tSA08), 4
   
    '// Get dimensions
    m_W = oDIB32.Width - 1
    m_H = oDIB32.Height - 1
   
    '// Dither...
    Select Case m_Palette
        Case [ipOptimal]: pvDitherToPalette_Ordered
        Case [ipBrowser]: pvDitherToHT216_Ordered
    End Select
    
    '// Unmap DIB bits
    CopyMemory ByVal VarPtrArray(m_Bits32()), 0&, 4
    CopyMemory ByVal VarPtrArray(m_Bits08()), 0&, 4
End Sub



'========================================================================================
' Private
'========================================================================================

Private Sub pvBuildRGB4096LUT(oPal As cPal8bpp)

  Dim R As Long
  Dim G As Long
  Dim B As Long

    '-- Build 4096-colors palette inverse indexes LUT
    For R = 0 To 15
    For G = 0 To 15
    For B = 0 To 15
        oPal.ClosestIndex R * 17, G * 17, B * 17, m_RGB4096_Inv(R, G, B)
    Next B, G, R
End Sub

'//

Private Sub pvDitherToHT216_Ordered()
  
  Dim lodX As Long, lodY As Long
  Dim lodT As Long, lodI As Long
  Dim newR As Long, newG As Long, newB As Long
  
    For m_y = 0 To m_H
        For m_x = 0 To m_W
        
            '-- Threshold/Inc.
            lodT = m_OD_Thr(lodX, lodY)
            lodI = m_OD_Hiv(lodX, lodY)
            
            '-- Inc ord. matrix column
            lodX = lodX + 1
            If (lodX = 4) Then lodX = 0
           
            '-- Ordered dither
            If (m_HT216_Trn(m_Bits32(m_x, m_y).B) < lodT) Then
                newB = m_HT216_Trn(m_Bits32(m_x, m_y).B + &H1A - lodI)
              Else
                newB = m_HT216_Trn(m_Bits32(m_x, m_y).B - lodI)
            End If
            If (m_HT216_Trn(m_Bits32(m_x, m_y).G) < lodT) Then
                newG = m_HT216_Trn(m_Bits32(m_x, m_y).G + &H1A - lodI)
              Else
                newG = m_HT216_Trn(m_Bits32(m_x, m_y).G - lodI)
            End If
            If (m_HT216_Trn(m_Bits32(m_x, m_y).R) < lodT) Then
                newR = m_HT216_Trn(m_Bits32(m_x, m_y).R + &H1A - lodI)
              Else
                newR = m_HT216_Trn(m_Bits32(m_x, m_y).R - lodI)
            End If
            
            '-- Set 8-bpp palette index
            m_Bits08(m_x, m_y) = m_HT216_Inv(newR, newG, newB)
        Next m_x
        
        '-- Reset column and inc ord. matrix row
        lodX = 0
        lodY = lodY + 1
        If lodY = 4 Then lodY = 0
    Next m_y
End Sub

Private Sub pvDitherToPalette_Ordered()

  Dim lodX As Long, lodY As Long
  Dim lodT As Long, lodI As Long
  Dim newR As Long, newG As Long, newB As Long
    
    For m_y = 0 To m_H
        For m_x = 0 To m_W
            
            '-- Threshold/Inc.
            lodT = m_OD_Thr(lodX, lodY)
            lodI = m_OD_Oiv(lodX, lodY)

            '-- Inc ord. matrix column
            lodX = lodX + 1
            If (lodX = 4) Then lodX = 0
            
            '-- Ordered dither
            If (m_RGB4096_Trn(m_Bits32(m_x, m_y).B) < lodT) Then
                newB = m_RGB4096_Trn(m_Bits32(m_x, m_y).B + 8 - lodI)
              Else
                newB = m_RGB4096_Trn(m_Bits32(m_x, m_y).B - lodI)
            End If
            If (m_RGB4096_Trn(m_Bits32(m_x, m_y).G) < lodT) Then
                newG = m_RGB4096_Trn(m_Bits32(m_x, m_y).G + 8 - lodI)
              Else
                newG = m_RGB4096_Trn(m_Bits32(m_x, m_y).G - lodI)
            End If
            If (m_RGB4096_Trn(m_Bits32(m_x, m_y).R) < lodT) Then
                newR = m_RGB4096_Trn(m_Bits32(m_x, m_y).R + 8 - lodI)
              Else
                newR = m_RGB4096_Trn(m_Bits32(m_x, m_y).R - lodI)
            End If
            
            '-- Set 8-bpp palette index
            m_Bits08(m_x, m_y) = m_RGB4096_Inv(newR, newG, newB)
        Next m_x
        
        '-- Reset column and inc ord. matrix
        lodX = 0
        lodY = lodY + 1
        If lodY = 4 Then lodY = 0
    Next m_y
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

Private Sub pvBuild_32bppSA(tSA As SAFEARRAY2D, oDIB As cDIB)

    '-- 32-bpp DIB mapping
    With tSA
        .cbElements = IIf(App.LogMode = 1, 1, 4)
        .cDims = 2
        .Bounds(0).lLbound = 0
        .Bounds(0).cElements = oDIB.Height
        .Bounds(1).lLbound = 0
        .Bounds(1).cElements = oDIB.Width
        .pvData = oDIB.lpBits
    End With
End Sub
