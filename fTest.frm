VERSION 5.00
Begin VB.Form fTest 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ordered dither"
   ClientHeight    =   6570
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   9510
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
   Icon            =   "fTest.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   438
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   634
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraTiming 
      Caption         =   "Timing"
      Height          =   1155
      Left            =   2850
      TabIndex        =   10
      Top             =   5235
      Width           =   1650
      Begin VB.Label lblTime 
         Caption         =   "0,0 ms"
         Height          =   300
         Left            =   180
         TabIndex        =   12
         Top             =   735
         Width           =   1050
      End
      Begin VB.Label lblLastTime 
         Caption         =   "Last dither time:"
         Height          =   240
         Left            =   180
         TabIndex        =   11
         Top             =   375
         Width           =   1260
      End
   End
   Begin VB.ComboBox cbZoom 
      Height          =   315
      ItemData        =   "fTest.frx":000C
      Left            =   8145
      List            =   "fTest.frx":001F
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   195
      Width           =   1185
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save 8-bpp bitmap"
      Height          =   480
      Left            =   7620
      TabIndex        =   1
      Top             =   5895
      Width           =   1695
   End
   Begin VB.Frame fraPalette 
      Caption         =   "Palette"
      Height          =   1155
      Left            =   180
      TabIndex        =   7
      Top             =   5235
      Width           =   2565
      Begin VB.OptionButton optPalette 
         Caption         =   "&Browser (Halftone-216)"
         Height          =   255
         Index           =   0
         Left            =   180
         TabIndex        =   8
         Top             =   360
         Value           =   -1  'True
         Width           =   2130
      End
      Begin VB.OptionButton optPalette 
         Caption         =   "&Optimal (Max.: 256 colors)"
         Height          =   255
         Index           =   1
         Left            =   180
         TabIndex        =   9
         Top             =   720
         Width           =   2280
      End
   End
   Begin VB.CommandButton cmdDither 
      Caption         =   "&Dither"
      Height          =   480
      Left            =   7620
      TabIndex        =   0
      Top             =   5250
      Width           =   1695
   End
   Begin OrderedDither.ucCanvas ucCanvasSrc 
      Height          =   4500
      Left            =   180
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   555
      Width           =   4500
      _ExtentX        =   7938
      _ExtentY        =   7938
   End
   Begin OrderedDither.ucCanvas ucCanvasTrg 
      Height          =   4500
      Left            =   4830
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   555
      Width           =   4500
      _ExtentX        =   7938
      _ExtentY        =   7938
   End
   Begin VB.Label lblZoom 
      Caption         =   "Zoom:"
      Height          =   270
      Left            =   7620
      TabIndex        =   13
      Top             =   255
      Width           =   690
   End
   Begin VB.Label lblTarget 
      Caption         =   "Target image (8-bpp)"
      Height          =   270
      Left            =   4830
      TabIndex        =   4
      Top             =   255
      Width           =   1800
   End
   Begin VB.Label lblSource 
      Caption         =   "Source image (32-bpp)"
      Height          =   270
      Left            =   195
      TabIndex        =   2
      Top             =   255
      Width           =   1800
   End
   Begin VB.Menu mnuImportTop 
      Caption         =   "&Import"
      Begin VB.Menu mnuImport 
         Caption         =   "&Load image..."
         Index           =   0
      End
      Begin VB.Menu mnuImport 
         Caption         =   "From &Clipboard"
         Index           =   1
      End
      Begin VB.Menu mnuImport 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuImport 
         Caption         =   "E&xit"
         Index           =   3
      End
   End
End
Attribute VB_Name = "fTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'================================================
' Project:       Ordered dither
' Author:        Carles P.V.
' Last revision: 2003.05.12
'================================================

Option Explicit

Private m_oTimer   As New cPCTimer ' HP timer
Private m_Filename As String       ' Temp. filename



Private Sub Form_Load()
    
    '-- Initialize module Look-UP tables
    mDither8bpp.InitializeLUTs
    
    '-- Default palette
    optPalette(0) = -1
    
    '-- Source canvas (No scroll)
    ucCanvasSrc.WorkMode = [cnvUserMode]
    
    '-- Reset zoom (100%)
    cbZoom.ListIndex = 0
End Sub

Private Sub Form_Paint()
    Line (0, 0)-(ScaleWidth, 0), vb3DShadow
    Line (0, 1)-(ScaleWidth, 1), vb3DHighlight
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    '-- Destroy objects
    Set ucCanvasSrc.DIB = Nothing
    Set ucCanvasTrg.DIB = Nothing
    Set m_oTimer = Nothing
    Set fTest = Nothing
End Sub

'//

Private Sub mnuImport_Click(Index As Integer)
    
  Dim sTmpFilename As String
    
    Select Case Index
    
        Case 0 '-- Load image...
        
            '-- Show open file dialog
            sTmpFilename = mDialogFile.GetFileName(m_Filename, "Supported formats|*.BMP;*.JPG;*.GIF", , "Open image", -1)
            
            If (Len(sTmpFilename)) Then
            
                m_Filename = sTmpFilename
                
                '-- Load image...
                DoEvents
                If (ucCanvasSrc.DIB.CreateFromStdPicture(LoadPicture(m_Filename), -1)) Then
                    ucCanvasSrc.Resize
                    ucCanvasSrc.Repaint
                    ucCanvasTrg.DIB.Destroy
                    ucCanvasTrg.Resize
                    ucCanvasTrg.Repaint
                  Else
                    MsgBox "Unexpected error loading image.", vbExclamation
                End If
            End If
            
        Case 1 '-- From Clipboard
            
            If (Clipboard.GetFormat(vbCFBitmap)) Then
                
                '-- Get from Clipboard
                DoEvents
                If (ucCanvasSrc.DIB.CreateFromStdPicture(Clipboard.GetData(vbCFBitmap), -1)) Then
                    ucCanvasSrc.Resize
                    ucCanvasSrc.Repaint
                    ucCanvasTrg.DIB.Destroy
                    ucCanvasTrg.Resize
                    ucCanvasTrg.Repaint
                  Else
                    MsgBox "Unexpected error loading image from Clipboard.", vbExclamation
                End If
              Else
                MsgBox "Nothing to import from Clipboard.", vbInformation
            End If
            
        Case 3 '-- Exit
            
            Unload Me
    End Select
End Sub

Private Sub optPalette_Click(Index As Integer)
    mDither8bpp.Palette = Index
End Sub

Private Sub cbZoom_Click()

    '-- Resize canvas
    ucCanvasTrg.Zoom = cbZoom.ListIndex + 1
    ucCanvasTrg.Resize
    ucCanvasTrg.Repaint
    ucCanvasSrc.Zoom = cbZoom.ListIndex + 1
    ucCanvasSrc.Resize
    
    '-- Force synchronization
    If (ucCanvasTrg.DIB.hDIB <> 0) Then
        ucCanvasTrg_Scroll
      Else
        ucCanvasSrc.Repaint
    End If
End Sub

'//

Private Sub ucCanvasTrg_Scroll()
  
  Dim x As Long
  Dim y As Long
    
    '-- Synchronize boths canvas (from Trg.)
    ucCanvasTrg.GetScrollPos x, y
    ucCanvasSrc.SetScrollPos x, y
    ucCanvasSrc.Repaint
End Sub

'//

Private Sub cmdDither_Click()

  Dim x As Long
  Dim y As Long

    If (ucCanvasSrc.DIB.hDIB = 0) Then
        MsgBox "Nothing to dither.", vbInformation
        Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
    
        '-- Reset timer
        m_oTimer.Reset
        '-- Dither...
        mDither8bpp.Dither ucCanvasSrc.DIB, ucCanvasTrg.DIB
        ucCanvasTrg.Resize
        '-- Get elapsed time
        lblTime.Caption = Format(m_oTimer.Elapsed, "#,#0.0 ms")

    Screen.MousePointer = vbDefault
    
    '-- Synchronize boths canvas (from Src.)
    ucCanvasSrc.GetScrollPos x, y
    ucCanvasTrg.SetScrollPos x, y
    ucCanvasTrg.Repaint
End Sub

Private Sub cmdSave_Click()

    If (ucCanvasTrg.DIB.hDIB = 0) Then
        MsgBox "Nothing to save.", vbInformation
        Exit Sub
    End If
    
    '-- Save 8-bpp bitmap
    ucCanvasTrg.DIB.Save App.Path & "\Test.bmp"
End Sub

