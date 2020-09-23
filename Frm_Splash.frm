VERSION 5.00
Begin VB.Form Frm_Splash 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   Caption         =   "FunTalk 1.01"
   ClientHeight    =   3075
   ClientLeft      =   -15
   ClientTop       =   -15
   ClientWidth     =   4305
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3075
   ScaleWidth      =   4305
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Cmd_close 
      Caption         =   "Close"
      Height          =   375
      Left            =   3120
      TabIndex        =   1
      Top             =   2520
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Timer SplashTimer 
      Interval        =   3000
      Left            =   3240
      Top             =   1200
   End
   Begin VB.Label About 
      BackStyle       =   0  'Transparent
      Caption         =   "              Written By Dodo2479@hotmail.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   855
      Left            =   600
      TabIndex        =   2
      Top             =   1800
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.Label txtlbl 
      BackStyle       =   0  'Transparent
      Caption         =   "             Written by Dodo2479@hotmail.com"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000C&
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   1920
      Width           =   4095
   End
   Begin VB.Image LogoPic 
      Height          =   720
      Index           =   10
      Left            =   2400
      Picture         =   "Frm_Splash.frx":0000
      Top             =   840
      Width           =   720
   End
   Begin VB.Image LogoPic 
      Height          =   720
      Index           =   9
      Left            =   1920
      Picture         =   "Frm_Splash.frx":0472
      Top             =   840
      Width           =   720
   End
   Begin VB.Image LogoPic 
      Height          =   720
      Index           =   8
      Left            =   1560
      Picture         =   "Frm_Splash.frx":0955
      Top             =   840
      Width           =   720
   End
   Begin VB.Image LogoPic 
      Height          =   720
      Index           =   7
      Left            =   3120
      Picture         =   "Frm_Splash.frx":0AAF
      Top             =   120
      Width           =   720
   End
   Begin VB.Image LogoPic 
      Height          =   720
      Index           =   6
      Left            =   1200
      Picture         =   "Frm_Splash.frx":0FC9
      Top             =   840
      Width           =   720
   End
   Begin VB.Image LogoPic 
      Height          =   720
      Index           =   5
      Left            =   2760
      Picture         =   "Frm_Splash.frx":143B
      Top             =   120
      Width           =   720
   End
   Begin VB.Image LogoPic 
      Height          =   720
      Index           =   4
      Left            =   2400
      Picture         =   "Frm_Splash.frx":18BC
      Top             =   120
      Width           =   720
   End
   Begin VB.Image LogoPic 
      Height          =   720
      Index           =   3
      Left            =   2040
      Picture         =   "Frm_Splash.frx":1DAA
      Top             =   120
      Width           =   720
   End
   Begin VB.Image LogoPic 
      Height          =   720
      Index           =   2
      Left            =   1560
      Picture         =   "Frm_Splash.frx":229C
      Top             =   120
      Width           =   720
   End
   Begin VB.Image LogoPic 
      Height          =   720
      Index           =   1
      Left            =   1080
      Picture         =   "Frm_Splash.frx":278A
      Top             =   120
      Width           =   720
   End
   Begin VB.Image LogoPic 
      Height          =   720
      Index           =   0
      Left            =   600
      Picture         =   "Frm_Splash.frx":2C50
      Top             =   120
      Width           =   720
   End
End
Attribute VB_Name = "Frm_Splash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Cmd_close_Click()
Unload Me
End Sub

Private Sub Form_Load()
'select what to do...
If IsAbout = True Then
    SplashTimer.Enabled = False
    txtlbl.Visible = False
    About.Visible = True
    Cmd_close.Visible = True
End If
    Me.Visible = True
End Sub

Private Sub SplashTimer_Timer()
    Load Frm_Main
    SplashTimer.Enabled = False
    Frm_Main.Show
    Unload Me
End Sub

