VERSION 5.00
Begin VB.Form Frm_Prefs 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FunTalk Settings..."
   ClientHeight    =   3660
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4740
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3660
   ScaleWidth      =   4740
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Audio Preferences"
      Height          =   3615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4695
      Begin VB.Frame Frame3 
         Caption         =   "Sample Rate (hz)"
         Height          =   1575
         Left            =   120
         TabIndex        =   13
         Top             =   1920
         Width           =   1455
         Begin VB.OptionButton Op_Rate 
            Caption         =   "44100"
            Height          =   255
            Index           =   0
            Left            =   360
            TabIndex        =   17
            Top             =   360
            Width           =   855
         End
         Begin VB.OptionButton Op_Rate 
            Caption         =   "22050"
            Height          =   195
            Index           =   1
            Left            =   360
            TabIndex        =   16
            Top             =   650
            Width           =   855
         End
         Begin VB.OptionButton Op_Rate 
            Caption         =   "11025"
            Height          =   255
            Index           =   2
            Left            =   360
            TabIndex        =   15
            Top             =   840
            Width           =   855
         End
         Begin VB.OptionButton Op_Rate 
            Caption         =   "8000"
            Height          =   195
            Index           =   3
            Left            =   360
            TabIndex        =   14
            Top             =   1100
            Width           =   855
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Sample Bits"
         Height          =   975
         Left            =   1680
         TabIndex        =   10
         Top             =   1920
         Width           =   1215
         Begin VB.OptionButton Op_Bits 
            Caption         =   "8 Bits"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   12
            Top             =   360
            Width           =   735
         End
         Begin VB.OptionButton Op_Bits 
            Caption         =   "16 Bits"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   11
            Top             =   600
            Width           =   855
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Sample Channels"
         Height          =   975
         Left            =   3000
         TabIndex        =   7
         Top             =   1920
         Width           =   1575
         Begin VB.OptionButton Op_Chan 
            Caption         =   "1 Mono"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   9
            Top             =   360
            Width           =   1215
         End
         Begin VB.OptionButton Op_Chan 
            Caption         =   "2 Stereo"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   8
            Top             =   600
            Width           =   1215
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Audio Devices"
         Height          =   1575
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   4455
         Begin VB.ComboBox CB_Cdev 
            Height          =   315
            Left            =   240
            Style           =   2  'Dropdown List
            TabIndex        =   19
            Top             =   1080
            Width           =   4095
         End
         Begin VB.ComboBox CB_Pdev 
            Height          =   315
            Left            =   240
            Style           =   2  'Dropdown List
            TabIndex        =   18
            Top             =   480
            Width           =   4095
         End
         Begin VB.Label txtlbl 
            AutoSize        =   -1  'True
            Caption         =   "Recording Device :"
            Height          =   195
            Index           =   1
            Left            =   240
            TabIndex        =   6
            Top             =   840
            Width           =   1380
         End
         Begin VB.Label txtlbl 
            AutoSize        =   -1  'True
            Caption         =   "Playback Device :"
            Height          =   195
            Index           =   0
            Left            =   240
            TabIndex        =   5
            Top             =   240
            Width           =   1305
         End
      End
      Begin VB.CommandButton Cmd_sSave 
         Caption         =   "Save"
         Height          =   375
         Left            =   1680
         TabIndex        =   3
         Top             =   3120
         Width           =   855
      End
      Begin VB.CommandButton Cmd_sCancel 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   2640
         TabIndex        =   2
         Top             =   3120
         Width           =   855
      End
      Begin VB.CommandButton Cmd_sOk 
         Caption         =   "OK"
         Height          =   375
         Left            =   3600
         TabIndex        =   1
         Top             =   3120
         Width           =   975
      End
   End
End
Attribute VB_Name = "Frm_Prefs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'#################################
'# FunTalk 1.01 Preferences Form #
'######################################################
'# This is where the dx setting are saved and loaded. #
'######################################################

Private Sub Cmd_sCancel_Click()
' if a user wants to abort the changes then load the last
' saved settings and set a new oldsettings, and close the settings form
    Frm_Main.mnu_pref_Click
End Sub

Private Sub Cmd_sOk_Click()
Dim Msg As String
'Set newprefs string to controll the old one
NewPrf = GetPrefs
'controll the changes... if made any
If NewPrf <> OldPrf Then
    Msg = MsgBox("You've changed some settings but didn't save then." & vbCrLf & vbCrLf & "Would you like to save then ?", vbQuestion & vbYesNo, "Save Settings ?")
        
        Select Case Msg
            Case 6 ' Yes save them...
                Cmd_sSave_Click         'Save the settings...
                Cmd_sOk_Click           'Run the OK command again (just checking)
        
            Case 7 ' No don't save them
                Call LoadPrefs          'Load the old settings.
                Cmd_sOk_Click        'Run the OK command again (Just checking)
        End Select
Else 'no changes are made... (or the are saved...)
Frm_Main.mnu_pref_Click
'    IsPrefVis = False
'    Me.Hide
End If
End Sub

Private Sub Cmd_sSave_Click()
    'Save the settings and set a new oldprefs
    Call SavePrefs
    'Tell the app. the preferences are changed
    IsChanged = True
    'Tell the user the settings are save...
    MsgBox "Settings Succesfully saved.", vbInformation, "Settings Saved"
End Sub

Private Sub Form_Load()
    Me.Icon = Frm_Main.Icon
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If IsClosing = False Then Cmd_sCancel_Click
End Sub

