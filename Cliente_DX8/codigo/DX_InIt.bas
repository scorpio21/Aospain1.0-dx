Attribute VB_Name = "Mod_DX"
Option Explicit

' [DX8] Motor grafico: Direct3D8 via clsDX8Engine (modDX8Requires.bas)
' [DX8] Audio: DirectSound8 + DirectMusic8 via clsAudio (modDX8Requires.bas)
' No se necesitan declaraciones propias de DX aqui - todo esta en modDX8Requires.bas

Public Sub IniciarObjetosDirectX()
    On Error Resume Next

    Call AddtoRichTextBox(frmCargando.status, "Iniciando DirectX8 Engine....", 0, 0, 0, 0, 0, True)

    Set SurfaceDB = New clsSurfaceManDynDX8
    Call engine.Engine_Init
    Call engine.setup_ambient

    Call AddtoRichTextBox(frmCargando.status, "Hecho", 0, 251, 0, 1, 0, False)

    If Musica = 0 Or Fx = 0 Then
        Call AddtoRichTextBox(frmCargando.status, "Iniciando Audio....", 0, 0, 0, 0, 0, True)
        Call Audio.Initialize(frmMain.hwnd, DirSound, DirMidi)
        Call AddtoRichTextBox(frmCargando.status, "Hecho", 0, 251, 0, 1, 0, False)
    End If

    Call AddtoRichTextBox(frmCargando.status, "!DirectX OK!", 0, 251, 0, 1, 0)
End Sub

Public Sub LiberarObjetosDX()
    On Error Resume Next
    Call engine.Engine_Deinit
    ' clsAudio se libera automaticamente en Class_Terminate
End Sub
