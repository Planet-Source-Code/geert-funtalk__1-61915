Attribute VB_Name = "Mod_Main"
'######
'# Main Module for FunTalk 1.01
'# Written by Dodo2479@hotmail.com
'######

Option Explicit

'Some program runners...
Public IsClosing As Boolean        'To tell the app. it's going to close
Public IsRunning As Boolean        'To tell if're talking
Public IsPrefVis As Boolean        'Status of preferences form.
Public IsChanged As Boolean        'For checked if preferences are changed
Public IsAbout As Boolean          'For the splash screen...
Public FRate As Long               'For setting the sample frequent (hz)
Public Xrate As Long
Public FChan As Long               'For setting the sample channels (Mono/Stereo)
Public FBits As Long               'For setting the sample bits (Bits per sample)
Public FcDev As Long               'The selected capture device (from the Settings file)
Public FpDev As Long               'The selected playback device (from the settings file)
Public FPref As String             'The settings file
Public OldPrf As String            'For checking changes in the preferences
Public NewPrf As String            'Same as OldPrf
Public FL As Integer               'The FreeFile

Sub Main()
FL = FreeFile
FPref = App.Path & "\Funtalk.dat"

IsClosing = False
IsRunning = False
IsPrefVis = False
IsAbout = False
IsChanged = True

Frm_Main.EnumarateDX 'check and get playback & capture devices...
Frm_Main.Visible = False
Call LoadPrefs       'Load preferences...
SetBG Frm_Splash, "Splash"
Load Frm_Splash      'show the splash
End Sub

Sub CloseApp()
'tell the app. it's going to close
IsClosing = True

Frm_Main.UnloadDX

'unload the forms...
Unload Frm_Prefs
Unload Frm_Main

'game over
End
End Sub

Function SetBG(Formname As Object, Color As String)

    Dim DLoop As Integer
    With Formname
    .DrawStyle = vbInsideSolid
    .DrawMode = vbCopyPen
    .ScaleMode = vbPixels
    .DrawWidth = 2
    .ScaleHeight = 256
    End With

    For DLoop = 0 To 255
        Select Case Color
            Case "Splash"
                Formname.Line (0, DLoop)-(Screen.Width, DLoop - 1), RGB(175, 255 - DLoop, 255), B
            Case "About"
                Formname.Line (0, DLoop)-(Screen.Width, DLoop - 1), RGB(175, 255 - DLoop, 175), B
        End Select
    Next DLoop

End Function
