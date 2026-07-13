VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDX8Engine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' [DX8] API GDI para crear fuentes sin depender de StdFont/IFont
Private Declare Function CreateFontGDI Lib "gdi32" Alias "CreateFontA" ( _
    ByVal nHeight As Long, ByVal nWidth As Long, ByVal nEscapement As Long, _
    ByVal nOrientation As Long, ByVal fnWeight As Long, ByVal fdwItalic As Long, _
    ByVal fdwUnderline As Long, ByVal fdwStrikeOut As Long, ByVal fdwCharSet As Long, _
    ByVal fdwOutputPrecision As Long, ByVal fdwClipPrecision As Long, _
    ByVal fdwQuality As Long, ByVal fdwPitchAndFamily As Long, _
    ByVal lpszFace As String) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Public bRunning As Boolean

Private Const FVF = D3DFVF_XYZRHW Or D3DFVF_TEX1 Or D3DFVF_DIFFUSE Or D3DFVF_SPECULAR
Private Const FVF2 = D3DFVF_XYZRHW Or D3DFVF_DIFFUSE Or D3DFVF_SPECULAR Or D3DFVF_TEX2

Dim font_count As Long
Dim font_last As Long

Private font_list() As D3DXFont

Public Enum FontAlignment
    fa_center = DT_CENTER
    fa_top = DT_TOP
    fa_left = DT_LEFT
    fa_topleft = DT_TOP Or DT_LEFT
    fa_bottomleft = DT_BOTTOM Or DT_LEFT
    fa_bottom = DT_BOTTOM
    fa_right = DT_RIGHT
    fa_bottomright = DT_BOTTOM Or DT_RIGHT
    fa_topright = DT_TOP Or DT_RIGHT
End Enum

Dim texture As Direct3DTexture8
Dim TransTexture As Direct3DTexture8

Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long

Public FPS As Integer
Private FramesPerSecCounter As Integer
Private timerElapsedTime As Single
Private timerTicksPerFrame As Double
Private particletimer As Single
Public engineBaseSpeed As Single
Private lFrameTimer As Long
Private lFrameLimiter As Long
Private ScrollPixelsPerFrameX As Byte
Private ScrollPixelsPerFrameY As Byte

Private TileBufferPixelOffsetX As Integer
Private TileBufferPixelOffsetY As Integer

Private MainViewTop As Integer
Private MainViewLeft As Integer
Private MainDestRect   As RECT
Private MainViewRect   As RECT
Private BackBufferRect As RECT

Private MainViewWidth As Integer
Private MainViewHeight As Integer

Private MouseTileX As Byte
Private MouseTileY As Byte
Private RLluvia(7)  As RECT  'RECT de la lluvia
Private iFrameIndex As Byte  'Frame actual de la LL
Private llTick      As Long  'Contador
Private LTLluvia(4) As Integer
Private WindowTileWidth As Integer
Private WindowTileHeight As Integer

Private HalfWindowTileWidth As Integer
Private HalfWindowTileHeight As Integer
Private Const GrhFogata As Integer = 1521

Private Type Light
    active As Boolean 'Do we ignore this light?
    id As Long
    map_x As Integer 'Coordinates
    map_y As Integer
    color As Long 'Start colour
    range As Byte
End Type

'Light list
Dim light_list() As Light
Dim light_count As Long
Dim light_last As Long

Private Enum PARTICLE_STATUS
    Alive = 0
    Dead = 1
End Enum

Private Type PARTICLE
    x As Single     'World Space Coordinates
    y As Single
    Z As Single
    vX As Single    'Speed and Direction
    vY As Single
    vZ As Single
    StartColor As D3DCOLORVALUE
    EndColor As D3DCOLORVALUE
    CurrentColor As D3DCOLORVALUE
    lifeTime As Long    'How long Mr. Particle Exists
    created As Long 'When this particle was created...
    status As PARTICLE_STATUS 'Does he even exist?
End Type

Private Type pa_gro
    PrtData() As PARTICLE
    PrtVertList() As TLVERTEX
    Position As D3DVECTOR
    Light As D3DLIGHT8
    type As Integer
    nParticles As Long
    ParticleSize As Single
    gravity As Single
    XWind As Single
    ZWind As Single
    YWind As Single
    XVariation As Single
    YVariation As Single
    ZVariation As Single
    x As Single
    y As Single
    Z As Single
    activated As Boolean
    texture As Integer
    size As Single
    life As Integer
End Type

Dim particle_group_list() As pa_gro
Dim particle_group_count As Integer
Dim particle_group_last As Integer

'Dim fxs_list() As New clsFXS
'Dim fxs_count As Integer
'Dim fxs_last As Integer


Dim valoreBlur As Long
Dim dimeTex As Long
Dim tex As Direct3DTexture8
Dim D3DbackBuffer As Direct3DSurface8
Dim zTarget As Direct3DSurface8
Dim stencil As Direct3DSurface8
Dim superTex As Direct3DSurface8
Dim blur As Boolean
Public blur_factor As Byte

Dim bump_map_texture As Direct3DTexture8
Dim bump_map_texture_ex As Direct3DTexture8
Dim bump_map_supported As Boolean
Dim bump_map_powa As Boolean
'Sets a Grh animation to loop indefinitely.







#Const HARDCODED = False 'True ' == MÁS FPS ^^






Public Function change_day_effect(ByVal o_r As Long, ByVal og As Long, ByVal ob As Long, ByVal r_ As Long, ByVal g_ As Long, ByVal b_ As Long) As Long
Dim r As Long
Dim g As Long
Dim b As Long
r = o_r
g = og
b = ob
    
    If o_r > r_ Then
        r = r - 1
    ElseIf o_r < r_ Then
        r = r + 1
    End If

    If og > g_ Then
        g = g - 1
    ElseIf og < r Then
        g = g + 1
    End If

    If ob > b_ Then
        b = b - 1
    ElseIf og < r Then
        b = b + 1
    End If
If r > 255 Then r = 255
If g > 255 Then g = 255
If b > 255 Then b = 255
day_r_old = r
day_g_old = g
day_b_old = b

If r < 0 Or g < 0 Or b < 0 Then r = 1: g = 1: b = 1
change_day_effect = ARGB(r, g, b, 255)

End Function

Public Sub setup_ambient()

'Noche 87, 61, 43
luz_dia(0).r = 87
luz_dia(0).g = 61
luz_dia(0).b = 43
luz_dia(1).r = 87
luz_dia(1).g = 61
luz_dia(1).b = 43
luz_dia(2).r = 87
luz_dia(2).g = 61
luz_dia(2).b = 43
luz_dia(3).r = 87
luz_dia(3).g = 61
luz_dia(3).b = 43
'4 am 124,117,91
luz_dia(4).r = 124
luz_dia(4).g = 127
luz_dia(4).b = 91
'5,6 am 143,137,135
luz_dia(5).r = 143
luz_dia(5).g = 137
luz_dia(5).b = 135
luz_dia(6).r = 143
luz_dia(6).g = 137
luz_dia(6).b = 135
'7 am 212,205,207
luz_dia(7).r = 212
luz_dia(7).g = 205
luz_dia(7).b = 207
luz_dia(8).r = 212
luz_dia(8).g = 205
luz_dia(8).b = 207
luz_dia(9).r = 212
luz_dia(9).g = 205
luz_dia(9).b = 207
luz_dia(10).r = 212
luz_dia(10).g = 205
luz_dia(10).b = 207
luz_dia(11).r = 212
luz_dia(11).g = 205
luz_dia(11).b = 207
luz_dia(12).r = 212
luz_dia(12).g = 205
luz_dia(12).b = 207
'Dia 255, 255, 255
luz_dia(12).r = 255
luz_dia(12).g = 255
luz_dia(12).b = 255
luz_dia(13).r = 255
luz_dia(13).g = 255
luz_dia(13).b = 255
'Medio Dia 255, 200, 255
luz_dia(14).r = 255
luz_dia(14).g = 250
luz_dia(14).b = 255
luz_dia(15).r = 255
luz_dia(15).g = 240
luz_dia(15).b = 255
luz_dia(16).r = 255
luz_dia(16).g = 230
luz_dia(16).b = 255
'17/18 0, 100, 255
luz_dia(17).r = 230
luz_dia(17).g = 230
luz_dia(17).b = 255
'18/19 0, 100, 255
luz_dia(18).r = 230
luz_dia(18).g = 230
luz_dia(18).b = 255
'19/20 156, 142, 83
luz_dia(19).r = 156
luz_dia(19).g = 142
luz_dia(19).b = 83
luz_dia(20).r = 87
luz_dia(20).g = 61
luz_dia(20).b = 43
luz_dia(21).r = 87
luz_dia(21).g = 61
luz_dia(21).b = 43
luz_dia(22).r = 87
luz_dia(22).g = 61
luz_dia(22).b = 43
luz_dia(23).r = 87
luz_dia(23).g = 61
luz_dia(23).b = 43
luz_dia(24).r = 87
luz_dia(24).g = 61
luz_dia(24).b = 43
            
End Sub


Private Function GetElapsedTime() As Single
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Gets the time that past since the last call
'**************************************************************
    Dim start_time As Currency
    Static end_time As Currency
    Static timer_freq As Currency

    'Get the timer frequency
    If timer_freq = 0 Then
        QueryPerformanceFrequency timer_freq
    End If
    
    'Get current time
    Call QueryPerformanceCounter(start_time)
    
    'Calculate elapsed time
    GetElapsedTime = (start_time - end_time) / timer_freq * 1000
    
    'Get next end time
    Call QueryPerformanceCounter(end_time)
End Function



Public Sub Text_Render(ByVal font As D3DXFont, ByVal Text As String, ByVal Top As Long, ByVal Left As Long, _
                                ByVal Width As Long, ByVal Height As Long, ByVal color As Long, ByVal format As Long, Optional ByVal shadow As Boolean = False)
'*****************************************************
'****** Coded by Menduz (lord.yo.wo@gmail.com) *******
'*****************************************************
    Dim TextRect As RECT
    Dim ShadowRect As RECT
    
    TextRect.Top = Top
    TextRect.Left = Left
    TextRect.bottom = Top + Height
    TextRect.Right = Left + Width
    
    If shadow Then
        ShadowRect.Top = Top - 1
        ShadowRect.Left = Left - 2
        ShadowRect.bottom = (Top + Height) - 1
        ShadowRect.Right = (Left + Width) - 2
        D3DX.DrawText font, &HFF000000, Text, ShadowRect, format
    End If
    
    D3DX.DrawText font, color, Text, TextRect, format
End Sub
Public Sub Text_Render_ext(ByVal Text As String, ByVal Top As Long, ByVal Left As Long, _
                                ByVal Width As Long, ByVal Height As Long, ByVal color As Long, Optional ByVal shadow As Boolean = False, Optional ByVal center As Boolean = False)
    If center = True Then
        Call Text_Render(font_list(1), Text, Top, Left, Width, Height, color, fa_center, shadow)
    Else
        Call Text_Render(font_list(1), Text, Top, Left, Width, Height, color, DT_TOP Or DT_LEFT, shadow)
    End If
End Sub

Private Sub Font_Make(ByVal font_index As Long, ByVal style As String, ByVal bold As Boolean, _
                        ByVal italic As Boolean, ByVal size As Long)
    If font_index > font_last Then
        font_last = font_index
        ReDim Preserve font_list(1 To font_last)
    End If
    font_count = font_count + 1
    
    ' [DX8] Crear fuente via API GDI directamente (sin StdFont/IFont)
    Dim hFont As Long
    hFont = CreateFontGDI(size * -1, 0, 0, 0, IIf(bold, 700, 400), italic, 0, 0, 0, 0, 0, 0, 0, style)
    Set font_list(font_index) = D3DX.CreateFont(D3DDevice, hFont)
    DeleteObject hFont
End Sub


Public Function Font_Create(ByVal style As String, ByVal size As Long, ByVal bold As Boolean, _
                            ByVal italic As Boolean) As Long
On Error GoTo ErrorHandler:
    Font_Create = Font_Next_Open
    Font_Make Font_Create, style, bold, italic, size
ErrorHandler:
    Font_Create = 0
End Function

Private Function Font_Next_Open() As Long
    Font_Next_Open = font_last + 1
End Function

Private Function Font_Check(ByVal font_index As Long) As Boolean
'*****************************************************
'****** Coded by Menduz (lord.yo.wo@gmail.com) *******
'*****************************************************
    If font_index > 0 And font_index <= font_last Then
        Font_Check = True
    End If
End Function

Function MakeVector(ByVal x As Single, ByVal y As Single, ByVal Z As Single) As D3DVECTOR
'*****************************************************
'****** Coded by Menduz (lord.yo.wo@gmail.com) *******
'*****************************************************
  MakeVector.x = x
  MakeVector.y = y
  MakeVector.Z = Z
End Function

Public Sub Engine_Init()
'*****************************************************
'****** Coded by Menduz (lord.yo.wo@gmail.com) *******
'*****************************************************
'On Error GoTo ErrHandler:

    Dim DispMode As D3DDISPLAYMODE
    Dim DispModeBK As D3DDISPLAYMODE
    Dim D3DWindow As D3DPRESENT_PARAMETERS
    Dim ColorKeyVal As Long
    
    Set SurfaceDB = New clsSurfaceManDynDX8
    
    Set dX = New DirectX8
    Set D3D = dX.Direct3DCreate()
    Set D3DX = New D3DX8
    
    D3D.GetAdapterDisplayMode D3DADAPTER_DEFAULT, DispMode
    D3D.GetAdapterDisplayMode D3DADAPTER_DEFAULT, DispModeBK
    
    
    With D3DWindow
        .Windowed = True
        .SwapEffect = D3DSWAPEFFECT_COPY
        .BackBufferFormat = DispMode.format
        .BackBufferWidth = frmMain.renderer.ScaleWidth
        .BackBufferHeight = frmMain.renderer.ScaleHeight
        .EnableAutoDepthStencil = 1
        .AutoDepthStencilFormat = D3DFMT_D16
        .hDeviceWindow = frmMain.renderer.hwnd
    End With
    DispMode.format = D3DFMT_X8R8G8B8
    If D3D.CheckDeviceFormat(0, D3DDEVTYPE_HAL, DispMode.format, 0, D3DRTYPE_TEXTURE, D3DFMT_A8R8G8B8) = D3D_OK Then
        Dim Caps8 As D3DCAPS8
        D3D.GetDeviceCaps 0, D3DDEVTYPE_HAL, Caps8
        If (Caps8.TextureOpCaps And D3DTEXOPCAPS_DOTPRODUCT3) = D3DTEXOPCAPS_DOTPRODUCT3 Then
            bump_map_supported = True
        Else
            bump_map_supported = False
            DispMode.format = DispModeBK.format
        End If
    Else
        bump_map_supported = False
        DispMode.format = DispModeBK.format
    End If
    Set D3DDevice = D3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, frmMain.renderer.hwnd, D3DCREATE_SOFTWARE_VERTEXPROCESSING, _
                                                            D3DWindow)
                                                            
                                                            
    HalfWindowTileHeight = (frmMain.renderer.ScaleHeight / 32) \ 2
    HalfWindowTileWidth = (frmMain.renderer.ScaleWidth / 32) \ 2
    
    TileBufferSize = 9
    TileBufferPixelOffsetX = (TileBufferSize - 1) * 32
    TileBufferPixelOffsetY = (TileBufferSize - 1) * 32
    
    D3DDevice.SetVertexShader FVF
    
    '//Transformed and lit vertices dont need lighting
    '   so we disable it...
    D3DDevice.SetRenderState D3DRS_LIGHTING, False
    
    D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
    D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
    D3DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, True
    
    Call SurfaceDB.Init(D3DX, D3DDevice, General_Get_Free_Ram_Bytes)

    engineBaseSpeed = 0.02
    base_light = &HFFFFFFFF ' blanco opaco por defecto (sin sistema de luces dinamicas)
    
    ' [Adaptado] MapData ya se redimensiona en InitTileEngine (TileEngine.bas)
    ' No la tocamos aca para no pisar datos cargados despues
    
    'Set FPS value to 60 for startup
    FPS = 60
    FramesPerSecCounter = 60
    
    ScrollPixelsPerFrameX = 9
    ScrollPixelsPerFrameY = 9
    
    ' [Adaptado] UserPos y los bordes Min/Max ya se calculan correctamente
    ' en InitTileEngine (TileEngine.bas) usando WindowTileWidth/Height.
    ' No los tocamos aca para evitar pisar esos valores o desbordar el tipo Byte
    ' (MinXBorder/MaxXBorder son Byte; ScaleWidth/64 puede dar valores invalidos).

    
    'partículas
    D3DDevice.SetRenderState D3DRS_POINTSIZE, Engine_FToDW(2)
    D3DDevice.SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_MODULATE
    D3DDevice.SetRenderState D3DRS_POINTSPRITE_ENABLE, 1
    D3DDevice.SetRenderState D3DRS_POINTSCALE_ENABLE, 0
    ' [Adaptado] Se elimino la inicializacion de motion blur (D3DbackBuffer, zTarget,
    ' stencil, tex, superTex) porque dimeTex nunca se asignaba (quedaba en 0),
    ' lo que hacia fallar D3DX.CreateTexture(device, 0, 0, ...) en tiempo de ejecucion.
    ' El blur tampoco se usa en Render() (fue simplificado a solo EndScene/Present).
    blur_factor = 10
    'bump mapping
    
    Font_Create "Tahoma", 8, True, 0
    'Font_Create "Verdana", 8, False, 0
    
bRunning = True
Exit Sub
errHandler:
Debug.Print "Error Number Returned: " & Err.Number
bRunning = False
End Sub

Public Sub Engine_Deinit()
' [Adaptado] Se elimino "Erase charlist" (esa variable no existe en este engine
' adaptado) y el "End" (mataba toda la aplicacion de golpe, sin dejar que el
' resto de LiberarObjetosDX/cierre de conexion terminen su trabajo).
Set D3DDevice = Nothing
Set D3D = Nothing
Set dX = Nothing
End Sub


Private Function CreateTLVertex(x As Single, y As Single, Z As Single, rhw As Single, color As Long, Specular As Long, tu As Single, tv As Single) As TLVERTEX
'*****************************************************
'****** Coded by Menduz (lord.yo.wo@gmail.com) *******
'*****************************************************
    CreateTLVertex.x = x
    CreateTLVertex.y = y
    CreateTLVertex.Z = Z
    CreateTLVertex.rhw = rhw
    CreateTLVertex.color = color
    CreateTLVertex.Specular = Specular
    CreateTLVertex.tu = tu
    CreateTLVertex.tv = tv
End Function

Public Sub Engine_ActFPS()
    If GetTickCount - lFrameTimer > 1000 Then
        FPS = FramesPerSecCounter
        FramesPerSecCounter = 0
        lFrameTimer = GetTickCount
    End If
End Sub

Function InMapBounds(ByVal x As Integer, ByVal y As Integer) As Boolean
'*****************************************************************
'Checks to see if a tile position is in the maps bounds
'*****************************************************************
    If x < XMinMapSize Or x > XMaxMapSize Or y < YMinMapSize Or y > YMaxMapSize Then
        Exit Function
    End If
    
    InMapBounds = True
End Function

' [Adaptado] Se elimino Draw_Grh (usaba MapData.light_value, incompatible
' con el tipo MapBlock de este proyecto). No es invocada por codigo propio.


Public Sub Engine_Blur_Set(ByVal bool As Boolean)
'*****************************************************
'****** Coded by Menduz (lord.yo.wo@gmail.com) *******
'*****************************************************
blur = bool
If blur = True Then
    D3DDevice.SetTextureStageState 0, D3DTSS_ALPHAARG1, D3DTA_TFACTOR
Else
    D3DDevice.SetTextureStageState 0, D3DTSS_ALPHAARG1, D3DTA_TEXTURE
End If
End Sub

Public Sub Engine_Blur_Toggle()
'*****************************************************
'****** Coded by Menduz (lord.yo.wo@gmail.com) *******
'*****************************************************
blur = Not blur
If blur = True Then
    D3DDevice.SetTextureStageState 0, D3DTSS_ALPHAARG1, D3DTA_TFACTOR
Else
    D3DDevice.SetTextureStageState 0, D3DTSS_ALPHAARG1, D3DTA_TEXTURE
End If
End Sub

Public Sub Engine_BeginFrame()
'*****************************************************
'****** Adaptado: BeginScene + Clear, llamado antes de*
'****** dibujar cualquier sprite con                  *
'****** Device_Box_Textured_Render                    *
'*****************************************************
D3DDevice.BeginScene
D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, 0, 1#, 0
End Sub

Public Sub Render()
'*****************************************************
'****** Adaptado: EndScene + Present.                 *
'****** El dibujo real (mapa, chars, dialogos) ya lo   *
'****** hizo RenderScreen/MostrarTexto/etc en          *
'****** General.bas, entre Engine_BeginFrame y este    *
'****** Sub, usando Device_Box_Textured_Render y       *
'****** Text_Render_ext por cada elemento.             *
'*****************************************************

D3DDevice.EndScene
D3DDevice.Present ByVal 0, ByVal 0, 0, ByVal 0

lFrameLimiter = GetTickCount
FramesPerSecCounter = FramesPerSecCounter + 1
timerElapsedTime = GetElapsedTime()
timerTicksPerFrame = timerElapsedTime * engineBaseSpeed
particletimer = timerElapsedTime * 0.05
End Sub



' [Adaptado] Se eliminaron ShowNextFrame() y RenderScreen() del engine original.
' Colisionaban en nombre con el RenderScreen propio de TileEngine.bas, y usaban
' Draw_Grh/MapData.light_value (incompatible con el tipo MapBlock de este proyecto).
' El render de mapa ya lo hace RenderScreen en TileEngine.bas, que llama
' directamente a Device_Box_Textured_Render.


Private Function Geometry_Create_TLVertex(ByVal x As Single, ByVal y As Single, ByVal Z As Single, _
                                            ByVal rhw As Single, ByVal color As Long, ByVal Specular As Long, tu As Single, _
                                            ByVal tv As Single) As TLVERTEX
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'**************************************************************
    Geometry_Create_TLVertex.x = x
    Geometry_Create_TLVertex.y = y
    Geometry_Create_TLVertex.Z = Z
    Geometry_Create_TLVertex.rhw = rhw
    Geometry_Create_TLVertex.color = color
    Geometry_Create_TLVertex.Specular = Specular
    Geometry_Create_TLVertex.tu = tu
    Geometry_Create_TLVertex.tv = tv
End Function

Private Function Geometry_Create_TLVertex2(x As Single, y As Single, Z As Single, rhw As Single, color As Long, Specular As Long, tu1 As Single, tv1 As Single, tu2 As Single, tv2 As Single) As TLVERTEX2
'mz
Geometry_Create_TLVertex2.x = x
Geometry_Create_TLVertex2.y = y
Geometry_Create_TLVertex2.Z = Z
Geometry_Create_TLVertex2.rhw = rhw
Geometry_Create_TLVertex2.color = color
Geometry_Create_TLVertex2.Specular = Specular
Geometry_Create_TLVertex2.tu1 = tu1
Geometry_Create_TLVertex2.tv1 = tv1
Geometry_Create_TLVertex2.tu2 = tu2
Geometry_Create_TLVertex2.tv2 = tv2
End Function

Private Sub Geometry_Create_Box(ByRef verts() As TLVERTEX, ByRef dest As RECT, ByRef src As RECT, ByRef rgb_list() As Long, _
                                Optional ByRef Textures_Width As Long, Optional ByRef Textures_Height As Long, Optional ByVal angle As Single)
'**************************************************************
'Author: Aaron Perkins
'Modified by Juan Martín Sotuyo Dodero
'Last Modify Date: 11/17/2002
'
' * v1      * v3
' |\        |
' |  \      |
' |    \    |
' |      \  |
' |        \|
' * v0      * v2
'**************************************************************
    Dim x_center As Single
    Dim y_center As Single
    Dim radius As Single
    Dim x_Cor As Single
    Dim y_Cor As Single
    Dim left_point As Single
    Dim right_point As Single
    Dim temp As Single
    
    If angle > 0 Then
        'Center coordinates on screen of the square
        x_center = dest.Left + (dest.Right - dest.Left) / 2
        y_center = dest.Top + (dest.bottom - dest.Top) / 2
        
        'Calculate radius
        radius = Sqr((dest.Right - x_center) ^ 2 + (dest.bottom - y_center) ^ 2)
        
        'Calculate left and right points
        temp = (dest.Right - x_center) / radius
        right_point = Atn(temp / Sqr(-temp * temp + 1))
        left_point = PI - right_point
    End If
    
    'Calculate screen coordinates of sprite, and only rotate if necessary
    If angle = 0 Then
        x_Cor = dest.Left
        y_Cor = dest.bottom
    Else
        x_Cor = x_center + Cos(-left_point - angle) * radius
        y_Cor = y_center - Sin(-left_point - angle) * radius
    End If
    
    
    '0 - Bottom left vertex
    If Textures_Width And Textures_Height Then
        verts(0) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(0), 0, src.Left / Textures_Width, (src.bottom + 1) / Textures_Height)
    Else
        verts(0) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(0), 0, 0, 0)
    End If
    'Calculate screen coordinates of sprite, and only rotate if necessary
    If angle = 0 Then
        x_Cor = dest.Left
        y_Cor = dest.Top
    Else
        x_Cor = x_center + Cos(left_point - angle) * radius
        y_Cor = y_center - Sin(left_point - angle) * radius
    End If
    
    
    '1 - Top left vertex
    If Textures_Width And Textures_Height Then
        verts(1) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(1), 0, src.Left / Textures_Width, src.Top / Textures_Height)
    Else
        verts(1) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(1), 0, 0, 0)
    End If
    'Calculate screen coordinates of sprite, and only rotate if necessary
    If angle = 0 Then
        x_Cor = dest.Right
        y_Cor = dest.bottom
    Else
        x_Cor = x_center + Cos(-right_point - angle) * radius
        y_Cor = y_center - Sin(-right_point - angle) * radius
    End If
    
    
    '2 - Bottom right vertex
    If Textures_Width And Textures_Height Then
        verts(2) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(2), 0, (src.Right + 1) / Textures_Width, (src.bottom + 1) / Textures_Height)
    Else
        verts(2) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(2), 0, 0, 0)
    End If
    'Calculate screen coordinates of sprite, and only rotate if necessary
    If angle = 0 Then
        x_Cor = dest.Right
        y_Cor = dest.Top
    Else
        x_Cor = x_center + Cos(right_point - angle) * radius
        y_Cor = y_center - Sin(right_point - angle) * radius
    End If
    
    
    '3 - Top right vertex
    If Textures_Width And Textures_Height Then
        verts(3) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(3), 0, (src.Right + 1) / Textures_Width, src.Top / Textures_Height)
    Else
        verts(3) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(3), 0, 0, 0)
    End If

End Sub

Private Sub Device_Box_Textured_Render_Advance(ByVal grhindex As Long, ByVal dest_x As Integer, ByVal dest_y As Integer, ByVal src_width As Integer, _
                                            ByVal src_height As Integer, ByRef rgb_list() As Long, ByVal src_x As Integer, _
                                            ByVal src_y As Integer, ByVal dest_width As Integer, Optional ByVal dest_height As Integer, _
                                            Optional ByVal alpha_blend As Boolean, Optional ByVal angle As Single)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/15/2003
'Copies the Textures allowing resizing
'Modified by Juan Martín Sotuyo Dodero
'**************************************************************
    Static src_rect As RECT
    Static dest_rect As RECT
    Static temp_verts(3) As TLVERTEX
    Static d3dTextures As D3D8Textures
    Static light_value(0 To 3) As Long

    If grhindex = 0 Then Exit Sub
    Set d3dTextures.texture = SurfaceDB.GetTexture(GrhData(grhindex).FileNum, d3dTextures.texwidth, d3dTextures.texheight)
    
    light_value(0) = rgb_list(0)
    light_value(1) = rgb_list(1)
    light_value(2) = rgb_list(2)
    light_value(3) = rgb_list(3)
        
    'If Not char_current_blind Then
    '    If (light_value(0) = 0) Then light_value(0) = 0
    '    If (light_value(1) = 0) Then light_value(1) = 0
    '    If (light_value(2) = 0) Then light_value(2) = 0
    '    If (light_value(3) = 0) Then light_value(3) = 0
    'Else
        light_value(0) = &HFFFFFFFF 'blind_color
        light_value(1) = &HFFFFFFFF 'blind_color
        light_value(2) = &HFFFFFFFF 'blind_color
        light_value(3) = &HFFFFFFFF 'blind_color
    'End If
    
    'Set up the source rectangle
    With src_rect
        .bottom = src_y + src_height
        .Left = src_x
        .Right = src_x + src_width
        .Top = src_y
    End With
        
    'Set up the destination rectangle
    With dest_rect
        .bottom = dest_y + dest_height
        .Left = dest_x
        .Right = dest_x + dest_width
        .Top = dest_y
    End With
    
    'Set up the TempVerts(3) vertices
    Geometry_Create_Box temp_verts(), dest_rect, src_rect, light_value(), d3dTextures.texwidth, d3dTextures.texheight, angle
        
    'Set Textures
    D3DDevice.SetTexture 0, d3dTextures.texture
    
    If alpha_blend Then
       'Set Rendering for alphablending
        D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_ONE
        D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_ONE
    End If
    
    'Draw the triangles that make up our square Textures
    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, temp_verts(0), Len(temp_verts(0))
    
    If alpha_blend Then
        'Set Rendering for colokeying
        D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
        D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
    End If

End Sub

Public Sub Device_Box_Textured_Render(ByVal grhindex As Long, ByVal dest_x As Integer, ByVal dest_y As Integer, ByVal src_width As Integer, _
                                            ByVal src_height As Integer, ByRef rgb_list() As Long, ByVal src_x As Integer, _
                                            ByVal src_y As Integer, Optional ByVal alpha_blend As Boolean, Optional ByVal angle As Single)
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 2/12/2004
'Just copies the Textures
'**************************************************************
    Static src_rect As RECT
    Static dest_rect As RECT
    Static temp_verts(3) As TLVERTEX
    Static d3dTextures As D3D8Textures
    Static light_value(0 To 3) As Long
    
    If grhindex = 0 Then Exit Sub
    Set d3dTextures.texture = SurfaceDB.GetTexture(GrhData(grhindex).FileNum, d3dTextures.texwidth, d3dTextures.texheight)
    
    light_value(0) = rgb_list(0)
    light_value(1) = rgb_list(1)
    light_value(2) = rgb_list(2)
    light_value(3) = rgb_list(3)
    
    'If Not char_current_blind Then
        If (light_value(0) = 0) Then light_value(0) = base_light
        If (light_value(1) = 0) Then light_value(1) = base_light
        If (light_value(2) = 0) Then light_value(2) = base_light
        If (light_value(3) = 0) Then light_value(3) = base_light
    'Else
    '    light_value(0) = &HFFFFFFFF 'blind_color
    '    light_value(1) = &HFFFFFFFF 'blind_color
    '    light_value(2) = &HFFFFFFFF 'blind_color
    '    light_value(3) = &HFFFFFFFF 'blind_color
    'End If
        
    'Set up the source rectangle
    With src_rect
        .bottom = src_y + src_height
        .Left = src_x
        .Right = src_x + src_width
        .Top = src_y
    End With
                
    'Set up the destination rectangle
    With dest_rect
        .bottom = dest_y + src_height
        .Left = dest_x
        .Right = dest_x + src_width
        .Top = dest_y
    End With
    
    'Set up the TempVerts(3) vertices
    Geometry_Create_Box temp_verts(), dest_rect, src_rect, light_value(), d3dTextures.texwidth, d3dTextures.texheight, angle
    
    'Set Textures
    D3DDevice.SetTexture 0, d3dTextures.texture
    
    If alpha_blend Then
       'Set Rendering for alphablending
        D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_ONE
        D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_ONE
    End If
    
    'Draw the triangles that make up our square Textures
    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, temp_verts(0), Len(temp_verts(0))
    
    If alpha_blend Then
        'Set Rendering for colokeying
        D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
        D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
    End If

End Sub


' [Adaptado] Se eliminaron las funciones Engine_MoveScreen, Char_Render, Start,
' SetCharacterFx, Char_Move_by_Head, Char_Move_by_Pos y EstaPCarea del engine original.
' Dependian del tipo Char y Enum E_Heading propios del cliente de referencia,
' incompatibles con el tipo Char de este proyecto (TileEngine.bas).
' Este cliente ya tiene su propio sistema de render de personajes y movimiento
' en TileEngine.bas/TileEnginegm.bas y General.bas, que llama directamente a
' Device_Box_Textured_Render para cada sprite.


' [Adaptado] Se eliminaron Particle_Group_Make, Particles_Update y CreateColorVal
' del engine original (sistema de particulas). Escribian en
' MapData(map_x, map_y).particle_group, campo que no existe en el tipo MapBlock
' de este proyecto. No son usadas por codigo propio.

Private Function Engine_FToDW(f As Single) As Long
' single > long
Dim buf As D3DXBuffer
    Set buf = D3DX.CreateBuffer(4)
    D3DX.BufferSetData buf, 0, 4, 1, f
    D3DX.BufferGetData buf, 0, 4, 1, Engine_FToDW
End Function

Private Function VectorToRGBA(Vec As D3DVECTOR, fHeight As Single) As Long
Dim r As Integer, g As Integer, b As Integer, A As Integer
    r = 127 * Vec.x + 128
    g = 127 * Vec.y + 128
    b = 127 * Vec.Z + 128
    A = 255 * fHeight
    VectorToRGBA = D3DColorARGB(A, r, g, b)
End Function



' [Adaptado] Se elimino el sistema de luces dinamicas (Light_Remove, Light_Create,
' Light_Move, Light_Render_All, etc.) y Particle_Group_Make/Particles_Update
' del engine original. Escribian en MapData(x,y).light_value, campo que no
' existe en el tipo MapBlock de este proyecto. No son usadas por codigo propio.

' [Adaptado] Se elimino RemoveCharAparence (dependia de E_Heading, SetCharacterFx
' y campos del tipo Char propios del cliente de referencia). No es usada por
' este proyecto.



