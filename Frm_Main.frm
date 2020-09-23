VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Frm_Main 
   AutoRedraw      =   -1  'True
   Caption         =   "FunTalk 1.01 By Dodo2479@hotmail.com"
   ClientHeight    =   2085
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   5775
   Icon            =   "Frm_Main.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2085
   ScaleWidth      =   5775
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Chk_I3D 
      Caption         =   "Room (Off)"
      Height          =   375
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CheckBox Chk_Waves 
      Caption         =   "Waves (Off)"
      Height          =   375
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CheckBox Chk_Flanger 
      Caption         =   "Flanger (Off)"
      Height          =   375
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CheckBox Chk_Echo 
      Caption         =   "Echo (Off)"
      Height          =   375
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CheckBox Chk_Gargle 
      Caption         =   "Gargle (Off) "
      Height          =   375
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CheckBox Chk_Distor 
      Caption         =   "Distorsion (Off)"
      Height          =   375
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CheckBox Chk_Compres 
      Caption         =   "Compres (Off)"
      Height          =   375
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   960
      Width           =   1215
   End
   Begin VB.CheckBox Chk_Chorus 
      Caption         =   "Chorus (Off)"
      Height          =   375
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   960
      Width           =   1215
   End
   Begin VB.CheckBox Chk_Pitch 
      Caption         =   "Pitch (Off)"
      Height          =   375
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "It's called Pitch, But it realy the frequenti of the buffer..."
      Top             =   960
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Pitch"
      Height          =   2055
      Left            =   1440
      TabIndex        =   10
      Top             =   0
      Width           =   615
      Begin ComctlLib.Slider S_Freq 
         Height          =   1335
         Left            =   240
         TabIndex        =   11
         Top             =   480
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   2355
         _Version        =   327682
         Orientation     =   1
         Min             =   -90000
         Max             =   -22583
         SelStart        =   -22583
         TickStyle       =   3
         TickFrequency   =   6000
         Value           =   -22583
      End
      Begin VB.Label txtlbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "High"
         Height          =   195
         Index           =   7
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   330
      End
      Begin VB.Label txtlbl 
         AutoSize        =   -1  'True
         Caption         =   "Low"
         Height          =   195
         Index           =   6
         Left            =   120
         TabIndex        =   12
         Top             =   1800
         Width           =   300
      End
   End
   Begin VB.Frame FR_Balance 
      Caption         =   "Balance"
      Height          =   855
      Left            =   2160
      TabIndex        =   5
      Top             =   0
      Width           =   2415
      Begin ComctlLib.Slider S_Balance 
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   450
         _Version        =   327682
         Min             =   -10000
         Max             =   10000
         TickFrequency   =   10000
      End
      Begin VB.Label txtlbl 
         AutoSize        =   -1  'True
         Caption         =   "Right"
         Height          =   195
         Index           =   4
         Left            =   1920
         TabIndex        =   9
         Top             =   240
         Width           =   375
      End
      Begin VB.Label txtlbl 
         AutoSize        =   -1  'True
         Caption         =   "Left"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   270
      End
      Begin VB.Label txtlbl 
         AutoSize        =   -1  'True
         Caption         =   "Middle"
         Height          =   195
         Index           =   3
         Left            =   840
         TabIndex        =   7
         Top             =   240
         Width           =   465
      End
   End
   Begin VB.Frame Fr_vol 
      Caption         =   "Volume (70%)"
      Height          =   2055
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   1215
      Begin ComctlLib.Slider S_Volume 
         Height          =   1695
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   2990
         _Version        =   327682
         Orientation     =   1
         LargeChange     =   1000
         Max             =   5000
         SelStart        =   1500
         TickFrequency   =   1000
         Value           =   1500
      End
      Begin VB.Label txtlbl 
         AutoSize        =   -1  'True
         Caption         =   "Min."
         Height          =   195
         Index           =   1
         Left            =   600
         TabIndex        =   4
         Top             =   1680
         Width           =   300
      End
      Begin VB.Label txtlbl 
         AutoSize        =   -1  'True
         Caption         =   "Max."
         Height          =   195
         Index           =   0
         Left            =   600
         TabIndex        =   3
         Top             =   360
         Width           =   345
      End
   End
   Begin VB.CommandButton Cmd_FT 
      Caption         =   "Start Talking"
      Height          =   735
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
   Begin VB.Menu Mnu_main 
      Caption         =   "FunTalk"
      Begin VB.Menu Mnu_inf 
         Caption         =   "About"
      End
      Begin VB.Menu line0 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_pref 
         Caption         =   "Preferences"
      End
      Begin VB.Menu line1 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_exit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "Frm_Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'###########
'#
'# This is the Funtalk 1.01 source
'#
'# Firts of all, Not all the code came from me...
'# For the recording and playback function...
'#
'# I'm olso a dutch radio amateur, mostly listening.
'# For my hobby i found some articles about listening with a computer.
'# Gerald Youngblood, AC5OG has written some good articles about
'# Digital Signal Processing and Software Defined Radios
'# He olso has includes the visual basic source for this.
'#
'# This brought me to the idea to make this...
'# I used some sources written by him. (Thank for that)
'#
'# I've you like... Please visit this link : http://www.flex-radio.com/articles_files
'# Take a look good at part 2 and part 3
'# It's very usefull.. (i think...)
'#
'# This is just a simple example with directx and a few functions of it.
'# Hopefully you'll like my idea on voice manipulation.
'#
'# Greetzz
'#
'#  Geert, Dodo2479@hotmail.com
'#############

Option Explicit

'Define Constants
Const CaptureSize As Long = 4096                'Capture Buffer size

'Define DirectX Objects
Dim DX As New DirectX8                          'DirectX object
Dim DS As DirectSound8                          'DirectSound object
Dim DsPBuffer As DirectSoundPrimaryBuffer8      'Primary buffer object
Dim DscBuffer As DirectSoundCaptureBuffer8      'Capture Buffer object
Dim DsBuffer As DirectSoundSecondaryBuffer8     'Output Buffer object
Dim DSC As DirectSoundCapture8                  'Capture object
Dim DspEnum As DirectSoundEnum8                 'Enumarate object for playback device
Dim DscEnum As DirectSoundEnum8                 'Enumarate object for capture device

'Define Type Definitions
Dim DscBufDesc As DSCBUFFERDESC                 'Capture buffer description
Dim DsBufDesc As DSBUFFERDESC                   'DirectSound buffer description
Dim DspBufDesc As WAVEFORMATEX                  'Primary buffer description
Dim DspCur As DSCURSORS                         'DirectSound Play Cursor

'Create I/O Sound Buffers
Dim InputBuf(CaptureSize) As Integer            'Demodulator Input Buffer
Dim OutputBuf(CaptureSize) As Integer           'Demodulator Output Buffer

'Define pointers and counters
Dim Pass As Long                                'Number of capture passes
Dim InPtr As Long                               'Capture Buffer block pointer
Dim OutPtr As Long                              'Output Buffer block pointer
Dim StartAddr As Long                           'Buffer block starting address
Dim EndAddr As Long                             'Ending buffer block address
Dim CaptureBytes As Long                        'Capture bytes to read

'Define loop counter variables for timing the capture event cycle
Dim TimeStart As Double                         'DirectX avg. loop timer variables
Dim TimeEnd As Double

'Set up Event variables for the Capture Buffer
Implements DirectXEvent8                        'Allows DirectX Events
Dim hEvent(1) As Long                           'Handle for DirectX Event
Dim nPos(1) As DSBPOSITIONNOTIFY                'Notify position array
Dim FirstPass As Boolean                        'Denotes first pass from Start

'Setup For Effects
Dim DsFreq As Long
Dim DsPan As Long
Dim DsVolume As Long

Dim DsPi As Long

'Start Talking.......
Sub FTOn()
Select Case IsChanged
    Case True
        UnloadDX                            'unload old dx
        LoadDevices                         'load new dx
        SetEvents                           'Set new event points
        IsChanged = False                   'Prefs. aren't changed anymore
        GoTo StartTalking                   'let's start the fun...
        
    Case False
        GoTo StartTalking
End Select

StartTalking:

Dim FXdesc() As DSEFFECTDESC


'Dim lResult As Long
Dim FXItems As Long
Dim FxI As Long

If IsRunning = True Then DsBuffer.Stop

FXItems = 0
If Chk_Echo.Value = 1 Then FXItems = FXItems + 1
If Chk_Chorus.Value = 1 Then FXItems = FXItems + 1
If Chk_Flanger.Value = 1 Then FXItems = FXItems + 1
If Chk_Compres.Value = 1 Then FXItems = FXItems + 1
If Chk_Distor.Value = 1 Then FXItems = FXItems + 1
If Chk_Gargle.Value = 1 Then FXItems = FXItems + 1
If Chk_Waves.Value = 1 Then FXItems = FXItems + 1
If Chk_I3D.Value = 1 Then FXItems = FXItems + 1

ReDim FXdesc(FXItems) As DSEFFECTDESC
ReDim lResult(FXItems) As Long

  FxI = 0
'apply the effexts...
    If Chk_Echo.Value = 1 Then
        FXdesc(FxI).guidDSFXClass = DSFX_STANDARD_ECHO
        FxI = FxI + 1
    End If
    If Chk_Chorus.Value = 1 Then
        FXdesc(FxI).guidDSFXClass = DSFX_STANDARD_CHORUS
        FxI = FxI + 1
    End If
    If Chk_Flanger.Value = 1 Then
        FXdesc(FxI).guidDSFXClass = DSFX_STANDARD_FLANGER
        FxI = FxI + 1
    End If
    If Chk_Compres.Value = 1 Then
        FXdesc(FxI).guidDSFXClass = DSFX_STANDARD_COMPRESSOR
        FxI = FxI + 1
    End If
    If Chk_Distor.Value = 1 Then
        FXdesc(FxI).guidDSFXClass = DSFX_STANDARD_DISTORTION
        FxI = FxI + 1
    End If
    If Chk_Gargle.Value = 1 Then
        FXdesc(FxI).guidDSFXClass = DSFX_STANDARD_GARGLE
        FxI = FxI + 1
    End If
    If Chk_Waves.Value = 1 Then
        FXdesc(FxI).guidDSFXClass = DSFX_STANDARD_WAVES_REVERB
        FxI = FxI + 1
    End If
    If Chk_I3D.Value = 1 Then
        FXdesc(FxI).guidDSFXClass = DSFX_STANDARD_I3DL2REVERB
        FxI = FxI + 1
    End If
    
 DsBuffer.SetFX FXItems, FXdesc, lResult
    
    'Set volume and balance
    With Frm_Main
        DsBuffer.SetVolume (0 - .S_Volume.Value)
        DsBuffer.SetPan .S_Balance.Value
    End With

    'set freq. case
    Select Case Chk_Pitch.Value
        Case 0
            DsBuffer.SetFrequency DSBFREQUENCY_ORIGINAL
        Case 1
            DsBuffer.SetFrequency (0 - S_Freq.Value)
    End Select

    DscBuffer.Start DSCBSTART_LOOPING       'Start Capture Looping
    IsRunning = True                        'Set flag to receive mode
    FirstPass = True                        'This is the first pass after Start
    OutPtr = 0
End Sub

'Turn Capture/Playback Off
Private Sub FTOff()
    IsRunning = False                        'Reset IsRunning flag
    FirstPass = False                        'Reset FirstPass flag
    DscBuffer.Stop                           'Stop Capture Loop
    DsBuffer.Stop                            'Stop Playback Loop
    DsBuffer.SetCurrentPosition 0            'rewind buffer
    OutPtr = 0
End Sub

'Enumarate the capture and playback devices
Public Sub EnumarateDX()
    Dim Idev As Integer
    
    Set DX = New DirectX8
    Set DspEnum = DX.GetDSEnum          'Playback devices...
    Set DscEnum = DX.GetDSCaptureEnum   'Capture devices...
    
    If DspEnum.GetCount <> 0 Then       'check if a playback device is found
        If DscEnum.GetCount <> 0 Then   'check is a capture device is found
            
            
            'If the device are found add then to it's combo
            With Frm_Prefs
                    For Idev = 1 To DspEnum.GetCount
                        .CB_Pdev.AddItem DspEnum.GetDescription(Idev)
                    Next Idev
                Idev = Empty
                    For Idev = 1 To DscEnum.GetCount
                        .CB_Cdev.AddItem DscEnum.GetDescription(Idev)
                    Next Idev
                .CB_Pdev.ListIndex = 0
                .CB_Cdev.ListIndex = 0
            End With
            Exit Sub
        Else
            MsgBox "FunTalk 1.01, Critical Error..." & vbCrLf & vbCrLf & "Funtalk wasn't able to find a 'Capture' device." & vbCrLf & "You can't run Funtalk 1.01 without a 'Capture' device..." & vbCrLf & vbCrLf & "Sorry, But Funtalk will close now...", vbCritical, "FunTalk 1.01 Critical Error."
            Call CloseApp
            Exit Sub
        End If
    Else
        MsgBox "FunTalk 1.01, Critical Error..." & vbCrLf & vbCrLf & "Funtalk wasn't able to find a 'Playback' device." & vbCrLf & "You can't run Funtalk 1.01 without a 'Playback' device..." & vbCrLf & vbCrLf & "Sorry, But Funtalk will close now...", vbCritical, "FunTalk 1.01 Critical Error."
        Call CloseApp
        Exit Sub
    End If
End Sub
'Set up the DirectSound Objects and the Capture and Play Buffer descriptions
Sub LoadDevices()

   On Local Error Resume Next
    Set DS = DX.DirectSoundCreate(DspEnum.GetGuid(FpDev))           'DirectSound object
    Set DSC = DX.DirectSoundCaptureCreate(DscEnum.GetGuid(FcDev))   'DirectSound Capture
    
    'Check to se if Sound Card is properly installed
    If Err.Number <> 0 Then
        MsgBox "Unable to start DirectSound. Check proper sound card installation"
        End
    End If
       
    'Set the cooperative level to allow formatting of the Primary Buffer
    DS.SetCooperativeLevel Me.hwnd, DSSCL_PRIORITY
    
    'Set up format for capture buffer
    With DscBufDesc
        With .fxFormat
            .nFormatTag = WAVE_FORMAT_PCM
            .nChannels = FChan                          'Stereo (I&Q)
            .lSamplesPerSec = FRate                    'Sample rate
            .nBitsPerSample = FBits                    '16 bit samples
            .nBlockAlign = .nBitsPerSample / 8 * .nChannels
            .lAvgBytesPerSec = .lSamplesPerSec * .nBlockAlign
        End With
        .lFlags = DSCBCAPS_DEFAULT
        .lBufferBytes = (DscBufDesc.fxFormat.nBlockAlign * CaptureSize) 'Buffer Size
        CaptureBytes = .lBufferBytes \ 2            'Bytes for 1/2 of capture buffer
    End With
    
    Set DscBuffer = DSC.CreateCaptureBuffer(DscBufDesc)       'Create the capture buffer
    
    ' Set up format for secondary playback buffer
    With DsBufDesc
        .fxFormat = DscBufDesc.fxFormat
        .lBufferBytes = DscBufDesc.lBufferBytes * 2  'Play is 2X Capture Buffer Size
        .lFlags = DSBCAPS_GLOBALFOCUS Or DSBCAPS_GETCURRENTPOSITION2 Or DSBCAPS_CTRLFREQUENCY Or DSBCAPS_CTRLFX Or DSBCAPS_CTRLVOLUME Or DSBCAPS_CTRLPAN
    End With
            
    DspBufDesc = DsBufDesc.fxFormat                           'Set Primary Buffer format
    DsPBuffer.SetFormat DspBufDesc                            'to same as Secondary Buffer
    
    Set DsBuffer = DS.CreateSoundBuffer(DsBufDesc)            'Create the secondary buffer

End Sub

'Set events for capture buffer notification at 0 and 1/2
Sub SetEvents()

    hEvent(0) = DX.CreateEvent(Me)
    hEvent(1) = DX.CreateEvent(Me)
    
    'Buffer Event 0 sets Write at 50% of buffer
    nPos(0).hEventNotify = hEvent(0)
    nPos(0).lOffset = (DscBufDesc.lBufferBytes \ 2) - 1  'First half of capture buffer
    
    'Buffer Event 1 Write at 100% of buffer
    nPos(1).hEventNotify = hEvent(1)
    nPos(1).lOffset = DscBufDesc.lBufferBytes - 1        'Second half of capture buffer
    
    DscBuffer.SetNotificationPositions 2, nPos()  'Set number of notification positions
    
End Sub



Private Sub DirectXEvent8_DXCallback(ByVal eventid As Long)

    StartTimer                          'Save loop start time
    
    Select Case eventid                 'Determine which Capture Block is ready
        Case hEvent(0)
            InPtr = 0                   'First half of Capture Buffer
        Case hEvent(1)
            InPtr = 1                   'Second half of Capture Buffer
    End Select
            
    StartAddr = InPtr * CaptureBytes    'Capture buffer starting address
       
    'Read from DirectX circular Capture Buffer to InputBuf
    DscBuffer.ReadBuffer StartAddr, CaptureBytes, InputBuf(0), DSCBLOCK_DEFAULT
    StartAddr = OutPtr * CaptureBytes   'Play buffer starting address
    EndAddr = OutPtr + CaptureBytes - 1 'Play buffer ending address
        
    With DsBuffer                                    'Reference DirectSoundBuffer
        
            .GetCurrentPosition DspCur         'Get current Play position
            'If true the write is overlapping the lWrite cursor due to load
            If DspCur.lWrite >= StartAddr _
                And DspCur.lWrite <= EndAddr Then

                FirstPass = True                'Restart play buffer
                OutPtr = 0
                StartAddr = 0
                
            End If
            
            'If true the write is overlapping the lPlay cursor due to load
            If DspCur.lPlay >= StartAddr _
                And DspCur.lPlay <= EndAddr Then

                FirstPass = True                'Restart play buffer
                OutPtr = 0
                StartAddr = 0
                
            End If
                        
        'Write OutputBuf to DirectX circular Secondary Buffer
        .WriteBuffer StartAddr, CaptureBytes, InputBuf(0), DSBLOCK_DEFAULT
        
    
        OutPtr = IIf(OutPtr >= 3, 0, OutPtr + 1)    'Counts 0 to 3
                
        If FirstPass = True Then        'On FirstPass wait 4 counts before starting
            Pass = Pass + 1             'the Secondary Play buffer looping at 0
            If Pass = 5 Then            'This puts the Play buffer three Capture cycles
                FirstPass = False       'after the current one
                Pass = 0                'Reset the Pass counter
                .SetCurrentPosition 0   'Set playback position to zero
                .Play DSBPLAY_LOOPING   'Start playback looping

            End If
        End If
        
    End With
    
    StopTimer                           'Display average loop time in immediate window
        
End Sub

Public Static Sub StartTimer()
    TimeStart = Timer
End Sub

Public Static Sub StopTimer()
    TimeEnd = Timer
End Sub

Public Sub UnloadDX()
On Error Resume Next
    If IsRunning = True Then
        DsBuffer.Stop                        'Stop Playback
        DscBuffer.Stop                       'Stop Capture
    End If
        
    Dim I As Integer

    For I = 0 To UBound(hEvent)                 'Kill DirectX Events
        DoEvents
        If hEvent(I) Then DX.DestroyEvent hEvent(I)
    Next

    Set DX = Nothing                            'Kill DirectX objects
    Set DS = Nothing
    Set DSC = Nothing
    Set DsBuffer = Nothing
    Set DscBuffer = Nothing
    
End Sub

'Let's Start
Private Sub Form_Load()
    Me.Show
End Sub

Private Sub Chk_Chorus_Click()
    Select Case Chk_Chorus.Value
        Case 0
            Chk_Chorus.Caption = "Chorus (Off)"
            If IsRunning = True Then FTOn
        Case 1
            Chk_Chorus.Caption = "Chorus (On)"
            If IsRunning = True Then FTOn
    End Select
End Sub

Private Sub Chk_Compres_Click()
   Select Case Chk_Compres.Value
        Case 0
            Chk_Compres.Caption = "Compres (Off)"
            If IsRunning = True Then FTOn
        Case 1
            Chk_Compres.Caption = "Compres (On)"
            If IsRunning = True Then FTOn
    End Select
End Sub
Private Sub Chk_Distor_Click()
   Select Case Chk_Distor.Value
        Case 0
            Chk_Distor.Caption = "Distorsion (Off)"
            If IsRunning = True Then FTOn
        Case 1
            Chk_Distor.Caption = "Distorsion (On)"
            If IsRunning = True Then FTOn
    End Select
End Sub
Private Sub Chk_Echo_Click()
    Select Case Chk_Echo.Value
        Case 0
            Chk_Echo.Caption = "Echo (Off)"
            If IsRunning = True Then FTOn
        Case 1
            Chk_Echo.Caption = "Echo (On)"
            If IsRunning = True Then FTOn
    End Select
End Sub

Private Sub Chk_Flanger_Click()
    Select Case Chk_Flanger.Value
        Case 0
            Chk_Flanger.Caption = "Flanger (Off)"
            If IsRunning = True Then FTOn
        Case 1
            Chk_Flanger.Caption = "Flanger (On)"
            If IsRunning = True Then FTOn
    End Select
End Sub
Private Sub Chk_Gargle_Click()
    Select Case Chk_Gargle.Value
        Case 0
            Chk_Gargle.Caption = "Gargle (Off)"
            If IsRunning = True Then FTOn
        Case 1
            Chk_Gargle.Caption = "Gargle (On)"
            If IsRunning = True Then FTOn
    End Select
End Sub
Private Sub Chk_Pitch_Click()
    Select Case Chk_Pitch.Value
        Case 0
            Chk_Pitch.Caption = "Pitch (Off)"
            S_Freq.Value = 44100
            If IsRunning = True Then FTOn
        Case 1
            Chk_Pitch.Caption = "Pitch (On)"
            If IsRunning = True Then FTOn
    End Select
End Sub
Private Sub Chk_Waves_Click()
    Select Case Chk_Waves.Value
        Case 0
            Chk_Waves.Caption = "Waves (Off)"
            If IsRunning = True Then FTOn
        Case 1
            Chk_Waves.Caption = "Waves (On)"
            If IsRunning = True Then FTOn
    End Select
End Sub
Private Sub Chk_I3D_Click()
    Select Case Chk_I3D.Value
        Case 0
            Chk_I3D.Caption = "Room (Off)"
            If IsRunning = True Then FTOn
        Case 1
            Chk_I3D.Caption = "Room (On)"
            If IsRunning = True Then FTOn
    End Select
End Sub
Private Sub Cmd_FT_Click()
    Select Case IsRunning
        Case True
            FTOff 'Stop running
            Cmd_FT.Caption = "Start Talking"
        Case False
            FTOn 'Start running
            Cmd_FT.Caption = "Stop Talking"
    End Select
End Sub

'Shutdown the normal way
Private Sub Form_Unload(Cancel As Integer)
    If IsClosing = False Then mnu_exit_Click
End Sub
'Shutdown the app.
Private Sub mnu_exit_Click()
    CloseApp
End Sub

Private Sub Mnu_inf_Click()
    IsAbout = True
    SetBG Frm_Splash, "About"
    Frm_Splash.Show
End Sub

Public Sub mnu_pref_Click()
    Select Case IsRunning
        Case True
            MsgBox "Could open the Preferences in talking mode." & vbCrLf & vbCrLf & "Please hit the 'Stop talking' button before you try to edit the preferences.", vbInformation, "Can't open preferences..."
            Exit Sub
        Case False
            'load the preferences
            Call LoadPrefs
            
                Select Case IsPrefVis
                    Case True 'Preferences is visable
                        IsPrefVis = False
                        Frm_Main.Cmd_FT.Enabled = True
                        Frm_Prefs.Visible = False
                
                    Case False 'preferences isn't viable
                        IsPrefVis = True
                        Frm_Main.Cmd_FT.Enabled = False
                        Frm_Prefs.Visible = True
                End Select
    End Select
End Sub

Private Sub S_Balance_Scroll()
    If IsRunning = False Then Exit Sub
    DsBuffer.SetPan S_Balance.Value
End Sub

Private Sub S_Freq_Scroll()
'Oké I know this is NOt a real pitching function,
'But you've got olmost the same result.
    If Chk_Pitch.Value = 0 Then Chk_Pitch.Value = 1
    'if talking is on reset it
    If IsRunning = True Then FTOn
End Sub

Private Sub S_Volume_Scroll()
    'calculate percentage
    Fr_vol.Caption = "Volume (" & "" & CInt(100 - Abs(100 * (S_Volume.Value / 5000))) & "%)"
    'exit if we're not taling
    If IsRunning = False Then Exit Sub
    'set volume if talking
    DsBuffer.SetVolume (0 - S_Volume.Value)
End Sub

