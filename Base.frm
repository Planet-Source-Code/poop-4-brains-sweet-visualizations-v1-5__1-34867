VERSION 5.00
Begin VB.Form Base 
   Caption         =   "Sweet Visualizations v1 by Kevin Fleet"
   ClientHeight    =   6795
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   10455
   Icon            =   "Base.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   453
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   697
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrFPS 
      Interval        =   1000
      Left            =   2880
      Top             =   1200
   End
   Begin VB.PictureBox grad 
      AutoRedraw      =   -1  'True
      BackColor       =   &H0000FF00&
      BorderStyle     =   0  'None
      Height          =   4500
      Left            =   6840
      Picture         =   "Base.frx":5C12
      ScaleHeight     =   300
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   4
      TabIndex        =   6
      Top             =   0
      Visible         =   0   'False
      Width           =   60
   End
   Begin VB.PictureBox Board 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   2160
      ScaleHeight     =   29
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   4
      Top             =   1560
      Width           =   255
   End
   Begin VB.Frame Stuff 
      BorderStyle     =   0  'None
      Height          =   336
      Left            =   720
      TabIndex        =   1
      Top             =   720
      Visible         =   0   'False
      Width           =   3360
      Begin VB.CommandButton StartButton 
         Caption         =   "&Start"
         Height          =   336
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   804
      End
      Begin VB.CommandButton StopButton 
         Caption         =   "S&top"
         Enabled         =   0   'False
         Height          =   336
         Left            =   864
         TabIndex        =   2
         Top             =   0
         Width           =   804
      End
   End
   Begin VB.ComboBox cmbV 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   315
      ItemData        =   "Base.frx":6A64
      Left            =   1560
      List            =   "Base.frx":6A86
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   0
      Width           =   6735
   End
   Begin VB.ComboBox DevicesBox 
      Height          =   315
      Left            =   1320
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Label lblX 
      BackColor       =   &H00000000&
      Caption         =   "Normal"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Menu mnuMenu 
      Caption         =   "Menu"
      Begin VB.Menu mnuFull 
         Caption         =   "Full Screen"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "Base"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------------------
' Deeth Stereo Oscilloscope v1.0
' A simple oscilloscope application -- now in <<stereo>>
'----------------------------------------------------------------------
' Opens a waveform audio device for 8-bit 11kHz input, and plots the
' waveform to a window.  Can only be resized to a certain minimum
' size defined by the Shape box.
'----------------------------------------------------------------------
' It would be good to make this use the same double-buffering
' scheme as the Spectrum Analyzer.
'----------------------------------------------------------------------
' Murphy McCauley (MurphyMc@Concentric.NET) 08/12/99
'----------------------------------------------------------------------

Option Explicit

Private DevHandle As Long
Private InData(0 To 511) As Byte
Private InOldD(0 To 511) As Byte

Private Inited As Boolean
Public MinHeight As Long, MinWidth As Long

Private Type WaveFormatEx
    FormatTag As Integer
    Channels As Integer
    SamplesPerSec As Long
    AvgBytesPerSec As Long
    BlockAlign As Integer
    BitsPerSample As Integer
    ExtraDataSize As Integer
End Type

Private Type WaveHdr
    lpData As Long
    dwBufferLength As Long
    dwBytesRecorded As Long
    dwUser As Long
    dwFlags As Long
    dwLoops As Long
    lpNext As Long 'wavehdr_tag
    Reserved As Long
End Type

Private Type WaveInCaps
    ManufacturerID As Integer      'wMid
    ProductID As Integer       'wPid
    DriverVersion As Long       'MMVERSIONS vDriverVersion
    ProductName(1 To 32) As Byte 'szPname[MAXPNAMELEN]
    Formats As Long
    Channels As Integer
    Reserved As Integer
End Type

Private Const WAVE_INVALIDFORMAT = &H0&                 '/* invalid format */
Private Const WAVE_FORMAT_1M08 = &H1&                   '/* 11.025 kHz, Mono,   8-bit
Private Const WAVE_FORMAT_1S08 = &H2&                   '/* 11.025 kHz, Stereo, 8-bit
Private Const WAVE_FORMAT_1M16 = &H4&                   '/* 11.025 kHz, Mono,   16-bit
Private Const WAVE_FORMAT_1S16 = &H8&                   '/* 11.025 kHz, Stereo, 16-bit
Private Const WAVE_FORMAT_2M08 = &H10&                  '/* 22.05  kHz, Mono,   8-bit
Private Const WAVE_FORMAT_2S08 = &H20&                  '/* 22.05  kHz, Stereo, 8-bit
Private Const WAVE_FORMAT_2M16 = &H40&                  '/* 22.05  kHz, Mono,   16-bit
Private Const WAVE_FORMAT_2S16 = &H80&                  '/* 22.05  kHz, Stereo, 16-bit
Private Const WAVE_FORMAT_4M08 = &H100&                 '/* 44.1   kHz, Mono,   8-bit
Private Const WAVE_FORMAT_4S08 = &H200&                 '/* 44.1   kHz, Stereo, 8-bit
Private Const WAVE_FORMAT_4M16 = &H400&                 '/* 44.1   kHz, Mono,   16-bit
Private Const WAVE_FORMAT_4S16 = &H800&                 '/* 44.1   kHz, Stereo, 16-bit

Private Const WAVE_FORMAT_PCM = 1

Private Const WHDR_DONE = &H1&              '/* done bit */
Private Const WHDR_PREPARED = &H2&          '/* set if this header has been prepared */
Private Const WHDR_BEGINLOOP = &H4&         '/* loop start block */
Private Const WHDR_ENDLOOP = &H8&           '/* loop end block */
Private Const WHDR_INQUEUE = &H10&          '/* reserved for driver */

Private Const WIM_OPEN = &H3BE
Private Const WIM_CLOSE = &H3BF
Private Const WIM_DATA = &H3C0

Private Declare Function waveInAddBuffer Lib "winmm" (ByVal InputDeviceHandle As Long, ByVal WaveHdrPointer As Long, ByVal WaveHdrStructSize As Long) As Long
Private Declare Function waveInPrepareHeader Lib "winmm" (ByVal InputDeviceHandle As Long, ByVal WaveHdrPointer As Long, ByVal WaveHdrStructSize As Long) As Long
Private Declare Function waveInUnprepareHeader Lib "winmm" (ByVal InputDeviceHandle As Long, ByVal WaveHdrPointer As Long, ByVal WaveHdrStructSize As Long) As Long

Private Declare Function waveInGetNumDevs Lib "winmm" () As Long
Private Declare Function waveInGetDevCaps Lib "winmm" Alias "waveInGetDevCapsA" (ByVal uDeviceID As Long, ByVal WaveInCapsPointer As Long, ByVal WaveInCapsStructSize As Long) As Long

Private Declare Function waveInOpen Lib "winmm" (WaveDeviceInputHandle As Long, ByVal WhichDevice As Long, ByVal WaveFormatExPointer As Long, ByVal CallBack As Long, ByVal CallBackInstance As Long, ByVal Flags As Long) As Long
Private Declare Function waveInClose Lib "winmm" (ByVal WaveDeviceInputHandle As Long) As Long

Private Declare Function waveInStart Lib "winmm" (ByVal WaveDeviceInputHandle As Long) As Long
Private Declare Function waveInReset Lib "winmm" (ByVal WaveDeviceInputHandle As Long) As Long
Private Declare Function waveInStop Lib "winmm" (ByVal WaveDeviceInputHandle As Long) As Long

Const vBar = 0
Const vCircle = 1
Const vColors = 2
Const vExplo = 3
Const vLines = 4
Const vScope = 5
Const vBackC = 6
Const vGradBars = 7
Const vIce = 8
Const vImp = 9

Const PI = 3.14

Dim VMode As Long

Sub InitDevices()
    Dim Caps As WaveInCaps, Which As Long
    DevicesBox.Clear
    For Which = 0 To waveInGetNumDevs - 1
        Call waveInGetDevCaps(Which, VarPtr(Caps), Len(Caps))
        'If Caps.Formats And WAVE_FORMAT_1M08 Then
        If Caps.Formats And WAVE_FORMAT_1S08 Then 'Now is 1S08 -- Check for devices that can do stereo 8-bit 11kHz
            Call DevicesBox.AddItem(StrConv(Caps.ProductName, vbUnicode), Which)
        End If
    Next
    If DevicesBox.ListCount = 0 Then
        MsgBox "You have no audio input devices!", vbCritical, "Ack!"
        End
    End If
    DevicesBox.ListIndex = 0
End Sub

Private Sub Form_Load()
Dim I As Long
For I = 0 To 255
CapSp(I) = 1
Next I
cmbV.ListIndex = 0
InitDevices
Me.Visible = True
StartButton_Click
End Sub

Private Sub Form_Resize()
If Me.WindowState = vbMinimized Then Exit Sub
cmbV.Width = Me.ScaleWidth
cmbV.Left = 0
Board.Top = cmbV.Height
Board.Left = 0
Board.Width = Me.ScaleWidth
Board.Height = Me.ScaleHeight - Board.Top
If lblX.Visible = True Then
cmbV.Width = Me.ScaleWidth - lblX.Width
cmbV.Left = lblX.Width
lblX.Height = cmbV.Height
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If DevHandle <> 0 Then
        Call DoStop
    End If
    Board.Cls
End Sub

Private Sub lblX_Click()
lblX.Visible = False
Me.WindowState = vbNormal
Me.BorderStyle = 2
End Sub

Private Sub mnuAbout_Click()
MsgBox "Sweet Visualizations v1 by Kevin Fleet " & vbCrLf & "Orignal Code by Murphy McCauly" & vbCrLf & "Please vote at PSC! - 4/30/2002", vbInformation, "About Sweet Visualizations"
End Sub

Private Sub mnuFull_Click()
Me.BorderStyle = 0
lblX.Visible = True
Me.WindowState = vbMaximized
End Sub

Private Sub StartButton_Click()
    Static WaveFormat As WaveFormatEx
    With WaveFormat
        .FormatTag = WAVE_FORMAT_PCM
        .Channels = 2 'Two channels -- left and right
        .SamplesPerSec = 11025 '11khz
        .BitsPerSample = 8
        .BlockAlign = (.Channels * .BitsPerSample) \ 8
        .AvgBytesPerSec = .BlockAlign * .SamplesPerSec
        .ExtraDataSize = 0
    End With
    
    Debug.Print "waveInOpen:"; waveInOpen(DevHandle, DevicesBox.ListIndex, VarPtr(WaveFormat), 0, 0, 0)
    
    If DevHandle = 0 Then
        Call MsgBox("Wave input device didn't open!", vbExclamation, "Ack!")
        Exit Sub
    End If
    Debug.Print " "; DevHandle
    Call waveInStart(DevHandle)
    
    Inited = True
       
    StopButton.Enabled = True
    StartButton.Enabled = False
    
    Call Visualize
End Sub


Private Sub StopButton_Click()
    Call DoStop
End Sub


Private Sub DoStop()
    Call waveInReset(DevHandle)
    Call waveInClose(DevHandle)
    DevHandle = 0
End Sub


Private Sub Visualize()
    Static Wave As WaveHdr
    
    Wave.lpData = VarPtr(InData(0))
    Wave.dwBufferLength = 512 'This is now 512 so there's still 256 samples per channel
    Wave.dwFlags = 0
    
    Do
    
        Call waveInPrepareHeader(DevHandle, VarPtr(Wave), Len(Wave))
        Call waveInAddBuffer(DevHandle, VarPtr(Wave), Len(Wave))
    
        Do
            'Nothing -- we're waiting for the audio driver to mark
            'this wave chunk as done.
        Loop Until ((Wave.dwFlags And WHDR_DONE) = WHDR_DONE) Or DevHandle = 0
        
        Call waveInUnprepareHeader(DevHandle, VarPtr(Wave), Len(Wave))
        
        If DevHandle = 0 Then
            'The device has closed...
            Exit Do
        End If
        
        Call DrawData
        
        DoEvents
    Loop While DevHandle <> 0 'While the audio device is open

End Sub

Function DrawData()
Static X As Long, g
Board.Cls

Select Case cmbV.ListIndex
Case vBar 'reg bars

    'right
    For X = 0 To 255
        Board.Line (0, X * 5)-(InData(X * 2), X * 5 + 3), vbGreen, BF
    Next X
    
    'left
    For X = 0 To 255
        Board.Line (Board.ScaleWidth, X * 5)-(Board.ScaleWidth - InData(X * 2), X * 5 + 3), vbRed, BF
    Next X

Case vCircle 'circle scope

    For X = 0 To 255
        Board.Circle (Board.ScaleWidth \ 2, X * 5), InData(X * 2) \ 2, vbBlue
    Next X

Case vColors 'colored squares
    Dim Width
    
    For X = 0 To 255 Step 5
        Width = InData(X * 2) * 2
        Board.Line (Board.ScaleWidth \ 2 - Width \ 2, Board.ScaleHeight \ 2 - Width \ 2)-(Board.ScaleWidth \ 2 + Width \ 2, Board.ScaleHeight \ 2 + Width \ 2), RGB(X, X, X), BF
    Next X
    
Case vExplo 'explo

    For X = 0 To 255
        Board.Circle (Board.ScaleWidth \ 2, Board.ScaleHeight \ 2), InData(X * 2), RGB(X, X, X)
    Next X
    
Case vLines 'lines
    
    For X = 0 To 254
        Board.Line (Board.ScaleWidth, Board.ScaleHeight \ 2)-(Board.ScaleWidth \ 2, InData(X * 2 + 2)), RGB(X, 0, 0)
        Board.Line (0, Board.ScaleHeight \ 2)-(Board.ScaleWidth \ 2, InData(X * 2)), RGB(X, 0, 0)
    Next X
    
Case vScope 'scope
    
    Dim stp As Long, Dx As Long
    stp = Board.ScaleWidth \ 255

    'right
    For X = 0 To 255
        Dx = X * stp
        Board.Line (Board.CurrentX, Board.CurrentY)-(Dx * 2, InData(X * 2)), vbBlue, BF
    Next X
    
    Board.CurrentX = 0
    Board.CurrentY = Board.ScaleWidth
    
    'left
    For X = 0 To 255
        Dx = X * stp
        Board.Line (Board.CurrentX, Board.CurrentY)-(Dx * 2, InData(X * 2 + 1)), vbRed, BF
    Next X
    
Case vBackC 'climate colors
    
    Dim Total As Double, Avg As Integer
    
    For X = 0 To 255
    Total = Total + InData(X)
    Next X
    Avg = Total / 255
    Board.Line (0, 0)-(Board.ScaleWidth, Board.ScaleHeight), RGB(Avg, 0, 0), BF
    Total = 0
    Avg = 0
Case vGradBars 'gradient bars
    
    For X = 0 To 255
    CapVal(X) = CapVal(X) + CapSp(X)
    CapSp(X) = CapSp(X) - 1
    If InData(X * 2) > CapVal(X) Then CapVal(X) = InData(X * 2) + 10: CapSp(X) = -5
    BitBlt Board.hDC, X * 5, Board.ScaleHeight - InData(X * 2), 4, InData(X * 2), grad.hDC, 0, grad.ScaleHeight - InData(X * 2), vbSrcCopy
    BitBlt Board.hDC, X * 5, Board.ScaleHeight - CapVal(X), 4, 3, grad.hDC, 0, grad.ScaleHeight - CapVal(X), vbSrcCopy
    Next X
    
Case vIce


    Dim N As Long, Color
    
    stp = Board.ScaleWidth \ (254 / 2)
    Board.CurrentY = Board.ScaleHeight / 2
    
    For N = 1 To MaxFade
        For X = 0 To (254 / 2)
            Dx = X * stp * 1.5
            Color = RGB(OldColor(N).r, OldColor(N).g, OldColor(N).b)
            Board.Line ((X + 2) * stp * 1.5, OldVal(X * 4 + 4, N))-(Dx, OldVal(X * 4, N)), Color
        Next X
        OldColor(N).r = OTrim(OldColor(N).r - 35)
        OldColor(N).g = OTrim(OldColor(N).g - 20)
        OldColor(N).b = OTrim(OldColor(N).b - 20)
    Next N

        OldI = OldI + 1: If OldI > 5 Then OldI = 1
        OldColor(OldI).r = 100
        OldColor(OldI).g = 100
        OldColor(OldI).b = 255
        For N = 0 To (254 / 2)
            OldVal(N * 4, OldI) = InData(N * 4)
        Next N
        
    Board.CurrentY = Board.ScaleHeight / 2
    Board.CurrentX = 0
    
    For X = 0 To (254 / 2)
        Dx = X * stp * 1.5
        Board.Line (Board.CurrentX, Board.CurrentY)-(Dx, InData(X * 4)), vbBlue
    Next X
    
Case vImp

Dim mak As Double, dstX, dstY, cr As Integer, cg As Integer, cb As Integer, ang As Long, olddstx, olddsty
mak = 360 / 255

olddstx = Me.ScaleWidth \ 2
olddsty = Me.ScaleHeight \ 2

For X = 0 To 255 Step 2
dstX = Cos((X * mak) / 180 * PI)
dstY = Sin((X * mak) / 180 * PI)
dstX = (Board.ScaleWidth \ 2) + (dstX * (InData(X) * 0.55))
dstY = (Board.ScaleHeight \ 2) + (dstY * (InData(X) * 0.55))
cr = InData(X) * 1.6
cg = InData(X) * 1.2
cb = InData(X) * 0.8
Board.DrawWidth = mak * 2.5
Board.Line (Board.ScaleWidth \ 2, Board.ScaleHeight \ 2)-(dstX, dstY), RGB(cr, cg, cb)
Board.Line (dstX, dstY)-(olddstx, olddsty), RGB(cr, cg * 0.7, cb * 0.7)
Board.DrawWidth = 1
olddstx = dstX
olddsty = dstY
Next X

End Select

Board.CurrentX = Board.ScaleWidth \ 2 - Board.TextWidth("FPS: " & TFPS) \ 2
Board.CurrentY = 10
Board.Print "FPS: " & TFPS

FPS = FPS + 1
End Function

Function OTrim(num As Long)
OTrim = num
If OTrim < 0 Then OTrim = 0
End Function

Private Sub tmrFPS_Timer()
TFPS = FPS
FPS = 0
End Sub
