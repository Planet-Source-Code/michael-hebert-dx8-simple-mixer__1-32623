VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "DirectX8 Simple Mixer"
   ClientHeight    =   5865
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   6945
   Icon            =   "SimpleMixer.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5865
   ScaleWidth      =   6945
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame10 
      Caption         =   "All Channels"
      Height          =   1335
      Left            =   5640
      TabIndex        =   36
      Top             =   4200
      Width           =   1095
      Begin VB.CommandButton allPlay 
         Caption         =   "Play"
         Height          =   255
         Left            =   120
         TabIndex        =   39
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton allPause 
         Caption         =   "Pause"
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   600
         Width           =   855
      End
      Begin VB.CommandButton allStop 
         Caption         =   "Stop"
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Top             =   960
         Width           =   855
      End
   End
   Begin VB.Frame Frame9 
      Caption         =   "Channel 4"
      Height          =   1335
      Left            =   4260
      TabIndex        =   32
      Top             =   4200
      Width           =   1095
      Begin VB.CommandButton ch4Play 
         Caption         =   "Play"
         Height          =   255
         Left            =   120
         TabIndex        =   35
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton ch4Pause 
         Caption         =   "Pause"
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   600
         Width           =   855
      End
      Begin VB.CommandButton ch4Stop 
         Caption         =   "Stop"
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   960
         Width           =   855
      End
   End
   Begin VB.Frame Frame8 
      Caption         =   "Channel 3"
      Height          =   1335
      Left            =   2880
      TabIndex        =   28
      Top             =   4200
      Width           =   1095
      Begin VB.CommandButton ch3Play 
         Caption         =   "Play"
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton ch3Pause 
         Caption         =   "Pause"
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   600
         Width           =   855
      End
      Begin VB.CommandButton ch3Stop 
         Caption         =   "Stop"
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   960
         Width           =   855
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "Channel 2"
      Height          =   1335
      Left            =   1500
      TabIndex        =   24
      Top             =   4200
      Width           =   1095
      Begin VB.CommandButton ch2Play 
         Caption         =   "Play"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton ch2Pause 
         Caption         =   "Pause"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   600
         Width           =   855
      End
      Begin VB.CommandButton ch2Stop 
         Caption         =   "Stop"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   960
         Width           =   855
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Channel 1"
      Height          =   1335
      Left            =   120
      TabIndex        =   20
      Top             =   4200
      Width           =   1095
      Begin VB.CommandButton ch1Stop 
         Caption         =   "Stop"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   960
         Width           =   855
      End
      Begin VB.CommandButton ch1Pause 
         Caption         =   "Pause"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   600
         Width           =   855
      End
      Begin VB.CommandButton ch1Play 
         Caption         =   "Play"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Balance Controls"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   4320
      TabIndex        =   18
      Top             =   240
      Width           =   2175
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Level Controls"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   2040
      TabIndex        =   17
      Top             =   240
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "File Names"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   360
      TabIndex        =   16
      Top             =   240
      Width           =   1455
   End
   Begin VB.HScrollBar ch4Pan 
      Height          =   255
      LargeChange     =   1000
      Left            =   4320
      Max             =   5000
      Min             =   -5000
      SmallChange     =   10
      TabIndex        =   14
      Top             =   3600
      Width           =   2175
   End
   Begin VB.HScrollBar ch4Vol 
      Height          =   255
      LargeChange     =   1000
      Left            =   2040
      Max             =   0
      Min             =   -10000
      SmallChange     =   10
      TabIndex        =   13
      Top             =   3600
      Value           =   -2500
      Width           =   2055
   End
   Begin VB.TextBox ch4Select 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   360
      TabIndex        =   12
      Text            =   "Click to Select"
      Top             =   3600
      Width           =   1455
   End
   Begin VB.HScrollBar ch3Pan 
      Height          =   255
      LargeChange     =   1000
      Left            =   4320
      Max             =   5000
      Min             =   -5000
      SmallChange     =   10
      TabIndex        =   10
      Top             =   2760
      Width           =   2175
   End
   Begin VB.HScrollBar ch3Vol 
      Height          =   255
      LargeChange     =   1000
      Left            =   2040
      Max             =   0
      Min             =   -10000
      SmallChange     =   10
      TabIndex        =   9
      Top             =   2760
      Value           =   -2500
      Width           =   2055
   End
   Begin VB.TextBox ch3Select 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   360
      TabIndex        =   8
      Text            =   "Click to Select"
      Top             =   2760
      Width           =   1455
   End
   Begin VB.HScrollBar ch2Pan 
      Height          =   255
      LargeChange     =   1000
      Left            =   4320
      Max             =   5000
      Min             =   -5000
      SmallChange     =   10
      TabIndex        =   6
      Top             =   1920
      Width           =   2175
   End
   Begin VB.HScrollBar ch2Vol 
      Height          =   255
      LargeChange     =   1000
      Left            =   2040
      Max             =   0
      Min             =   -10000
      SmallChange     =   10
      TabIndex        =   5
      Top             =   1920
      Value           =   -2500
      Width           =   2055
   End
   Begin VB.TextBox ch2Select 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   360
      TabIndex        =   4
      Text            =   "Click to Select"
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "Channel 1"
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   6615
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   960
         Top             =   120
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         DefaultExt      =   "wav"
         DialogTitle     =   "Select a Wave File..."
         Filter          =   "wav"
      End
      Begin VB.HScrollBar ch1Pan 
         Height          =   255
         LargeChange     =   1000
         Left            =   4200
         Max             =   5000
         Min             =   -5000
         SmallChange     =   10
         TabIndex        =   3
         Top             =   240
         Width           =   2175
      End
      Begin VB.HScrollBar ch1Vol 
         Height          =   255
         LargeChange     =   1000
         Left            =   1920
         Max             =   0
         Min             =   -10000
         SmallChange     =   10
         TabIndex        =   2
         Top             =   240
         Value           =   -2500
         Width           =   2055
      End
      Begin VB.TextBox ch1Select 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   240
         TabIndex        =   1
         Text            =   "Click to Select"
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Channel 2"
      Height          =   735
      Left            =   120
      TabIndex        =   7
      Top             =   1680
      Width           =   6615
      Begin MSComDlg.CommonDialog CommonDialog2 
         Left            =   960
         Top             =   120
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         DefaultExt      =   "wav"
         DialogTitle     =   "Select a Wave File..."
         Filter          =   "*.wav"
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Channel 3"
      Height          =   735
      Left            =   120
      TabIndex        =   11
      Top             =   2520
      Width           =   6615
      Begin MSComDlg.CommonDialog CommonDialog3 
         Left            =   1080
         Top             =   240
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         DefaultExt      =   "wav"
         DialogTitle     =   "Select a Wave File..."
         Filter          =   "*.wav"
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Channel 4"
      Height          =   735
      Left            =   120
      TabIndex        =   15
      Top             =   3360
      Width           =   6615
      Begin MSComDlg.CommonDialog CommonDialog4 
         Left            =   1080
         Top             =   120
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         DefaultExt      =   "wav"
         DialogTitle     =   "Select a Wave File..."
         Filter          =   "*.wav"
      End
   End
   Begin VB.Frame Frame5 
      Height          =   735
      Left            =   120
      TabIndex        =   19
      Top             =   0
      Width           =   6615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'DirectX8 Simple Mixer
' By Michael Hebert

'Define the DirectSound8 Object

Dim dx As New DirectX8
Dim ds As DirectSound8

'Create 4 Secondary Buffers

Dim dsBuffer(3) As DirectSoundSecondaryBuffer8

'Define Volume and Panning variables

Dim SetVolume As Long
Dim SetPan As Long

'Define FileName strings used for each channel

Dim ch1FileName As String
Dim ch2FileName As String
Dim ch3FileName As String
Dim ch4FileName As String

'Program Exit routine

Private Sub cmdExit_Click()

    Cleanup
    Unload Me
    
End Sub

'Make sure all buffers used by DirectX8 are killed

Private Sub Cleanup()

    If Not (dsBuffer(0) Is Nothing) Then dsBuffer(0).Stop
    Set dsBuffer(0) = Nothing
    If Not (dsBuffer(1) Is Nothing) Then dsBuffer(1).Stop
    Set dsBuffer(1) = Nothing
    If Not (dsBuffer(2) Is Nothing) Then dsBuffer(2).Stop
    Set dsBuffer(2) = Nothing
    If Not (dsBuffer(3) Is Nothing) Then dsBuffer(3).Stop
    Set dsBuffer(3) = Nothing
    Set ds = Nothing
    Set dx = Nothing
    
End Sub

'Load the form and create the DirectSound Object

Private Sub Form_Load()

    Me.Show
    On Local Error Resume Next
    Set ds = dx.DirectSoundCreate("")
    If Err.Number <> 0 Then
    MsgBox "Unable to start DirectSound"
    End
    End If
    ds.SetCooperativeLevel Me.hWnd, DSSCL_PRIORITY

End Sub

'Create DirectSound buffers and load files

Private Sub ch1Select_Click()

    'Describe the DirectSound buffer
    
    Dim bufferDesc As DSBUFFERDESC
    bufferDesc.lFlags = DSBCAPS_CTRLVOLUME Or DSBCAPS_CTRLPAN Or DSBCAPS_STATIC Or DSBCAPS_STICKYFOCUS
    
    'Setup the File Selection Dialog
    
    CommonDialog1.Filter = "Wave Files (*.wav)|*.wav"
    CommonDialog1.ShowOpen
    
        'Exit if Cancel clicked
    
        CommonDialog1.CancelError = True
            On Error GoTo CancelOpen
        
        'Load the file into the DirectSoundBuffer
        
    If CommonDialog1.FileName <> "" Then
        Set dsBuffer(0) = ds.CreateSoundBufferFromFile(CommonDialog1.FileName, bufferDesc)
        
        'Parse the Path to show only the selected FileName
        
        Dim x As Long
        x = InStrRev(CommonDialog1.FileName, "\")
        ch1Select.Text = LCase(Mid(CommonDialog1.FileName, x + 1))
    
    End If
    
    'Exit the subroutine in file successfully loaded
    
    Exit Sub
    
CancelOpen:

    Exit Sub
    
End Sub

'Repetitious code duplicating the selection routine for each channel

Private Sub ch2Select_Click()

    Dim bufferDesc As DSBUFFERDESC
    bufferDesc.lFlags = DSBCAPS_CTRLVOLUME Or DSBCAPS_CTRLPAN Or DSBCAPS_STATIC Or DSBCAPS_STICKYFOCUS
    CommonDialog2.Filter = "Wave Files (*.wav)|*.wav"
    CommonDialog2.ShowOpen
        CommonDialog2.CancelError = True
            On Error GoTo CancelOpen
    If CommonDialog2.FileName <> "" Then
        Set dsBuffer(1) = ds.CreateSoundBufferFromFile(CommonDialog2.FileName, bufferDesc)
        Dim x As Long
        x = InStrRev(CommonDialog2.FileName, "\")
        ch2Select.Text = LCase(Mid(CommonDialog2.FileName, x + 1))
    End If
Exit Sub

CancelOpen:
    Exit Sub

End Sub

Private Sub ch3Select_Click()

    Dim bufferDesc As DSBUFFERDESC
    bufferDesc.lFlags = DSBCAPS_CTRLVOLUME Or DSBCAPS_CTRLPAN Or DSBCAPS_STATIC Or DSBCAPS_STICKYFOCUS
    CommonDialog3.Filter = "Wave Files (*.wav)|*.wav"
    CommonDialog3.ShowOpen
        CommonDialog3.CancelError = True
            On Error GoTo CancelOpen
    If CommonDialog3.FileName <> "" Then
        Set dsBuffer(2) = ds.CreateSoundBufferFromFile(CommonDialog3.FileName, bufferDesc)
        Dim x As Long
        x = InStrRev(CommonDialog3.FileName, "\")
        ch3Select.Text = LCase(Mid(CommonDialog3.FileName, x + 1))
    End If
    Exit Sub

CancelOpen:
    Exit Sub
End Sub

Private Sub ch4Select_Click()

    Dim bufferDesc As DSBUFFERDESC
    bufferDesc.lFlags = DSBCAPS_CTRLVOLUME Or DSBCAPS_CTRLPAN Or DSBCAPS_STATIC Or DSBCAPS_STICKYFOCUS
    CommonDialog4.Filter = "Wave Files (*.wav)|*.wav"
    CommonDialog4.ShowOpen
        CommonDialog4.CancelError = True
            On Error GoTo CancelOpen
    If CommonDialog4.FileName <> "" Then
        Set dsBuffer(3) = ds.CreateSoundBufferFromFile(CommonDialog4.FileName, bufferDesc)
        Dim x As Long
        x = InStrRev(CommonDialog4.FileName, "\")
        ch4Select.Text = LCase(Mid(CommonDialog4.FileName, x + 1))
    End If
    Exit Sub
    
CancelOpen:
    Exit Sub

End Sub

Private Sub ch1Play_Click()

    'Exit the subroutine if there is nothing to play
    
    If dsBuffer(0) Is Nothing Then Exit Sub
    
    'Set the Volume and Panning parameters
    
    dsBuffer(0).SetVolume ch1Vol.Value
    dsBuffer(0).SetPan ch1Pan.Value
    
    'Loop the sound until told to stop
    
    dsBuffer(0).Play 1

End Sub

'Pause playing without changing play position in buffer

Private Sub ch1Pause_Click()

    If dsBuffer(0) Is Nothing Then Exit Sub
    dsBuffer(0).Stop

End Sub

'Stop playing and reset position to beginning of buffer

Private Sub ch1Stop_Click()

    If dsBuffer(0) Is Nothing Then Exit Sub
    dsBuffer(0).Stop
    dsBuffer(0).SetCurrentPosition 0

End Sub

'More repetitious code duplicating routines for remaining channels

Private Sub ch2Play_Click()

    If dsBuffer(1) Is Nothing Then Exit Sub
    dsBuffer(1).SetVolume ch2Vol.Value
    dsBuffer(1).SetPan ch2Pan.Value
    dsBuffer(1).Play 1

End Sub

Private Sub ch2Pause_Click()

    If dsBuffer(1) Is Nothing Then Exit Sub
    dsBuffer(1).Stop

End Sub

Private Sub ch2Stop_Click()

    If dsBuffer(1) Is Nothing Then Exit Sub
    dsBuffer(1).Stop
    dsBuffer(1).SetCurrentPosition 0

End Sub

Private Sub ch3Play_Click()

    If dsBuffer(2) Is Nothing Then Exit Sub
    dsBuffer(2).SetVolume ch3Vol.Value
    dsBuffer(2).SetPan ch3Pan.Value
    dsBuffer(2).Play 1

End Sub

Private Sub ch3Pause_Click()

    If dsBuffer(2) Is Nothing Then Exit Sub
    dsBuffer(2).Stop

End Sub

Private Sub ch3Stop_Click()

    If dsBuffer(2) Is Nothing Then Exit Sub
    dsBuffer(2).Stop
    dsBuffer(2).SetCurrentPosition 0

End Sub

Private Sub ch4Play_Click()

    If dsBuffer(3) Is Nothing Then Exit Sub
    dsBuffer(3).SetVolume ch4Vol.Value
    dsBuffer(3).SetPan ch4Pan.Value
    dsBuffer(3).Play 1

End Sub

Private Sub ch4Pause_Click()

    If dsBuffer(3) Is Nothing Then Exit Sub
    dsBuffer(3).Stop

End Sub

Private Sub ch4Stop_Click()

    If dsBuffer(3) Is Nothing Then Exit Sub
    dsBuffer(3).Stop
    dsBuffer(3).SetCurrentPosition 0

End Sub

Private Sub ch1Vol_Change()

    If dsBuffer(0) Is Nothing Then Exit Sub
    dsBuffer(0).SetVolume ch1Vol.Value

End Sub

Private Sub ch1Vol_Scroll()

    If dsBuffer(0) Is Nothing Then Exit Sub
    dsBuffer(0).SetVolume ch1Vol.Value

End Sub

Private Sub ch1Pan_Change()

    If dsBuffer(0) Is Nothing Then Exit Sub
    dsBuffer(0).SetPan ch1Pan.Value

End Sub

Private Sub ch1Pan_Scroll()

    If dsBuffer(0) Is Nothing Then Exit Sub
    dsBuffer(0).SetPan ch1Pan.Value

End Sub

Private Sub ch2Vol_Change()

    If dsBuffer(1) Is Nothing Then Exit Sub
    dsBuffer(1).SetVolume ch2Vol.Value

End Sub

Private Sub ch2Vol_Scroll()

    If dsBuffer(1) Is Nothing Then Exit Sub
    dsBuffer(1).SetVolume ch2Vol.Value

End Sub

Private Sub ch2Pan_Change()

    If dsBuffer(1) Is Nothing Then Exit Sub
    dsBuffer(1).SetPan ch2Pan.Value

End Sub

Private Sub ch2Pan_Scroll()

    If dsBuffer(1) Is Nothing Then Exit Sub
    dsBuffer(1).SetPan ch2Pan.Value

End Sub

Private Sub ch3Vol_Change()

    If dsBuffer(2) Is Nothing Then Exit Sub
    dsBuffer(2).SetVolume ch3Vol.Value

End Sub

Private Sub ch3Vol_Scroll()

    If dsBuffer(2) Is Nothing Then Exit Sub
    dsBuffer(2).SetVolume ch3Vol.Value

End Sub

Private Sub ch3Pan_Change()

    If dsBuffer(2) Is Nothing Then Exit Sub
    dsBuffer(2).SetPan ch3Pan.Value

End Sub

Private Sub ch3Pan_Scroll()

    If dsBuffer(2) Is Nothing Then Exit Sub
    dsBuffer(2).SetPan ch3Pan.Value

End Sub

Private Sub ch4Vol_Change()

    If dsBuffer(3) Is Nothing Then Exit Sub
    dsBuffer(3).SetVolume ch4Vol.Value

End Sub

Private Sub ch4Vol_Scroll()

    If dsBuffer(3) Is Nothing Then Exit Sub
    dsBuffer(3).SetVolume ch4Vol.Value

End Sub

Private Sub ch4Pan_Change()

    If dsBuffer(3) Is Nothing Then Exit Sub
    dsBuffer(3).SetPan ch4Pan.Value

End Sub

Private Sub ch4Pan_Scroll()

    If dsBuffer(3) Is Nothing Then Exit Sub
    dsBuffer(3).SetPan ch4Pan.Value

End Sub

'Play all the channels simultaneously

Private Sub allPlay_Click()

    ch1Play_Click
    ch2Play_Click
    ch3Play_Click
    ch4Play_Click

End Sub

'Pause all the channels simultaneously

Private Sub allPause_Click()

    ch1Pause_Click
    ch2Pause_Click
    ch3Pause_Click
    ch4Pause_Click

End Sub

'Stop all the channels simultaneously

Private Sub allStop_Click()

    ch1Stop_Click
    ch2Stop_Click
    ch3Stop_Click
    ch4Stop_Click

End Sub
