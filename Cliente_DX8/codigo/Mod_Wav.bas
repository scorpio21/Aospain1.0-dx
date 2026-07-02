Attribute VB_Name = "Mod_Wav"
Option Explicit

Public Const SND_SYNC = &H0
Public Const SND_ASYNC = &H1
Public Const SND_NODEFAULT = &H2
Public Const SND_LOOP = &H8
Public Const SND_NOSTOP = &H10

Public Const SND_CLICK = "click.Wav"
Public Const SND_PASOS1 = "23.Wav"
Public Const SND_PASOS2 = "24.Wav"
Public Const SND_NAVEGANDO = "50.wav"
Public Const SND_OVER = "click2.Wav"
Public Const SND_DICE = "cupdice.Wav"
Public Const SNDCHANNEL_CLICK = 0

' [DX8] Todo el audio via clsAudio (Audio en modDX8Requires.bas)
' Stub de compatibilidad para el codigo que llama LoadWavetoDSBuffer
Function LoadWavetoDSBuffer(DS As DirectSound8, DSB As DirectSoundSecondaryBuffer8, sFile As String) As Boolean
    LoadWavetoDSBuffer = True
End Function

Sub PlayWaveDS(File As String)
    If Fx = 1 Then Exit Sub
    On Error Resume Next
    Call Audio.PlayWave(File)
End Sub
