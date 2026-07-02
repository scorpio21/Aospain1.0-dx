Attribute VB_Name = "Mod_Config"
Option Explicit

' Rutas configurables via rutas.ini
' Si el archivo no existe, se usan las rutas por defecto del proyecto.
'
' Formato de rutas.ini (en la misma carpeta que el .exe):
'
' [Rutas]
' Graficos=Graficos
' Sonidos=Wav
' Musica=Midi
' Mapas=Mapas
'
' Tambien se pueden usar rutas absolutas o relativas:
' Musica=..\Recursos\musica\Midi
' Graficos=C:\MisGraficos\AO

Private Const INI_FILE = "\rutas.ini"
Private Const SECTION = "Rutas"
Private Const BUF_SIZE = 512

' Rutas cargadas en memoria (se leen una sola vez al iniciar)
Private m_Graficos As String
Private m_Sonidos  As String
Private m_Musica   As String
Private m_Mapas    As String
Private m_Loaded   As Boolean

Private Function IniPath() As String
    IniPath = App.Path & INI_FILE
End Function

Private Function ReadIni(ByVal key As String, ByVal default_val As String) As String
    Dim buf As String
    buf = Space(BUF_SIZE)
    Dim ret As Long
    ret = getprivateprofilestring(SECTION, key, default_val, buf, BUF_SIZE, IniPath)
    ReadIni = Left(buf, ret)
End Function

Public Sub WriteIni(ByVal key As String, ByVal value As String)
    writeprivateprofilestring SECTION, key, value, IniPath
End Sub

' Carga las rutas del INI (llamar una vez al inicio)
Public Sub CargarRutas()
    If m_Loaded Then Exit Sub
    
    ' Leer con defaults = los valores de Config_Inicio
    m_Graficos = ReadIni("Graficos", Config_Inicio.DirGraficos)
    m_Sonidos  = ReadIni("Sonidos",  Config_Inicio.DirSonidos)
    m_Musica   = ReadIni("Musica",   Config_Inicio.DirMusica)
    m_Mapas    = ReadIni("Mapas",    Config_Inicio.DirMapas)
    
    m_Loaded = True
    
    ' Si no existe el ini, crearlo con los valores por defecto para que el usuario lo pueda editar
    If Dir(IniPath) = "" Then
        WriteIni "Graficos", m_Graficos
        WriteIni "Sonidos",  m_Sonidos
        WriteIni "Musica",   m_Musica
        WriteIni "Mapas",    m_Mapas
    End If
End Sub

' Convierte una ruta relativa/absoluta en ruta absoluta con backslash final
Private Function ResolverRuta(ByVal ruta As String) As String
    Dim resultado As String
    ' Si es ruta absoluta (empieza con letra: o \\) la usamos tal cual
    If Len(ruta) >= 2 Then
        If Mid(ruta, 2, 1) = ":" Or Left(ruta, 2) = "\\" Then
            resultado = ruta
        Else
            ' Ruta relativa al directorio del exe
            resultado = App.Path & "\" & ruta
        End If
    Else
        resultado = App.Path & "\" & ruta
    End If
    ' Asegurar backslash final
    If Right(resultado, 1) <> "\" Then resultado = resultado & "\"
    ResolverRuta = resultado
End Function

' Funciones de acceso a rutas resueltas
Public Function RutaGraficos() As String
    If Not m_Loaded Then CargarRutas
    RutaGraficos = ResolverRuta(m_Graficos)
End Function

Public Function RutaSonidos() As String
    If Not m_Loaded Then CargarRutas
    RutaSonidos = ResolverRuta(m_Sonidos)
End Function

Public Function RutaMusica() As String
    If Not m_Loaded Then CargarRutas
    RutaMusica = ResolverRuta(m_Musica)
End Function

Public Function RutaMapas() As String
    If Not m_Loaded Then CargarRutas
    RutaMapas = ResolverRuta(m_Mapas)
End Function
