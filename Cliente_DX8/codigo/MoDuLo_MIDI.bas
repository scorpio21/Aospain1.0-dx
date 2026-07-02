Attribute VB_Name = "MoDuLo_MIDI"
Option Explicit

' [DX8] Todo el MIDI ahora se maneja via clsAudio (Audio en modDX8Requires.bas)
' Perf/Seg/Loader/SegState eliminados - clsAudio los maneja internamente

Public Const MIdi_Inicio = 6

Public CurMidi As String
Public LoopMidi As Byte
Public IsPlayingCheck As Boolean

Public Sub CargarMIDI(Archivo As String)
    If Musica = 1 Then Exit Sub
    On Error Resume Next
    Call Audio.PlayMIDI(Archivo)
    IsPlayingCheck = True
End Sub

Public Sub Stop_Midi()
    On Error Resume Next
    Call Audio.StopMidi
    IsPlayingCheck = False
End Sub

Public Sub Play_Midi()
    If Musica = 1 Then Exit Sub
    On Error Resume Next
    Call Audio.PlayMIDI(CurMidi)
    IsPlayingCheck = True
End Sub

Function Sonando() As Boolean
    If Musica = 1 Then Exit Function
    On Error Resume Next
    Sonando = Audio.MusicActivated And IsPlayingCheck
End Function
