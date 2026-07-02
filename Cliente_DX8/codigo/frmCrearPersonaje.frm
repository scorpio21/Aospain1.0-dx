VERSION 5.00
Begin VB.Form frmCrearPersonaje 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picHead 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   360
      Index           =   0
      Left            =   9150
      ScaleHeight     =   24
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   24
      TabIndex        =   41
      Top             =   6855
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.PictureBox picHead 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   360
      Index           =   4
      Left            =   10770
      ScaleHeight     =   24
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   24
      TabIndex        =   40
      Top             =   6855
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.PictureBox picHead 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   360
      Index           =   3
      Left            =   10365
      ScaleHeight     =   24
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   24
      TabIndex        =   39
      Top             =   6855
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.PictureBox picHead 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   360
      Index           =   2
      Left            =   9960
      ScaleHeight     =   24
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   24
      TabIndex        =   38
      Top             =   6855
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.PictureBox picHead 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   360
      Index           =   1
      Left            =   9555
      ScaleHeight     =   24
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   24
      TabIndex        =   37
      Top             =   6855
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.PictureBox picPJ 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   9840
      ScaleHeight     =   65
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   41
      TabIndex        =   36
      Top             =   7335
      Width           =   615
   End
   Begin VB.PictureBox picTemp 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   990
      Left            =   9840
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   39
      TabIndex        =   35
      Top             =   7335
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.ComboBox lstProfesion 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      ItemData        =   "frmCrearPersonaje.frx":0000
      Left            =   5925
      List            =   "frmCrearPersonaje.frx":0037
      Style           =   2  'Dropdown List
      TabIndex        =   31
      Top             =   2700
      Width           =   2820
   End
   Begin VB.ComboBox lstGenero 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      ItemData        =   "frmCrearPersonaje.frx":00D1
      Left            =   5925
      List            =   "frmCrearPersonaje.frx":00DB
      Style           =   2  'Dropdown List
      TabIndex        =   30
      Top             =   2250
      Width           =   2820
   End
   Begin VB.ComboBox lstRaza 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      ItemData        =   "frmCrearPersonaje.frx":00EE
      Left            =   5925
      List            =   "frmCrearPersonaje.frx":0107
      Style           =   2  'Dropdown List
      TabIndex        =   29
      Top             =   1800
      Width           =   2820
   End
   Begin VB.ComboBox lstHogar 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      ItemData        =   "frmCrearPersonaje.frx":0142
      Left            =   9015
      List            =   "frmCrearPersonaje.frx":0152
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   3510
      Width           =   2565
   End
   Begin VB.TextBox txtNombre 
      BackColor       =   &H80000012&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   3570
      TabIndex        =   0
      Top             =   1275
      Width           =   4815
   End
   Begin VB.Image HeadPJ 
      Height          =   225
      Index           =   0
      Left            =   8835
      Picture         =   "frmCrearPersonaje.frx":017F
      Top             =   6900
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image HeadPJ 
      Height          =   225
      Index           =   1
      Left            =   11220
      Picture         =   "frmCrearPersonaje.frx":0491
      Top             =   6900
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image DirPJ 
      Height          =   225
      Index           =   1
      Left            =   10320
      Picture         =   "frmCrearPersonaje.frx":07A3
      Top             =   8295
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image DirPJ 
      Height          =   225
      Index           =   0
      Left            =   9720
      Picture         =   "frmCrearPersonaje.frx":0AB5
      Top             =   8295
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      BorderStyle     =   3  'Dot
      Index           =   0
      Visible         =   0   'False
      X1              =   663
      X2              =   663
      Y1              =   457
      Y2              =   481
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      BorderStyle     =   3  'Dot
      Index           =   1
      Visible         =   0   'False
      X1              =   688
      X2              =   688
      Y1              =   457
      Y2              =   481
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      BorderStyle     =   3  'Dot
      Index           =   2
      Visible         =   0   'False
      X1              =   663
      X2              =   689
      Y1              =   456
      Y2              =   456
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      BorderStyle     =   3  'Dot
      Index           =   3
      Visible         =   0   'False
      X1              =   663
      X2              =   689
      Y1              =   481
      Y2              =   481
   End
   Begin VB.Label lblEspecialidad 
      BackStyle       =   0  'Transparent
      Caption         =   "Capinteria, Pesca, Mineria y Algo mas que va aqui"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   4080
      TabIndex        =   42
      Top             =   8160
      Width           =   1815
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   43
      Left            =   7770
      MouseIcon       =   "frmCrearPersonaje.frx":0DC7
      MousePointer    =   99  'Custom
      Top             =   7800
      Width           =   135
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   42
      Left            =   8400
      MouseIcon       =   "frmCrearPersonaje.frx":0F19
      MousePointer    =   99  'Custom
      Top             =   7800
      Width           =   135
   End
   Begin VB.Label Skill 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   21
      Left            =   7995
      TabIndex        =   34
      Top             =   7725
      Width           =   270
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "+3"
      ForeColor       =   &H00FFFF80&
      Height          =   195
      Left            =   4020
      TabIndex        =   33
      Top             =   4260
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image Image1 
      Height          =   3120
      Left            =   3270
      Stretch         =   -1  'True
      Top             =   4710
      Width           =   2475
   End
   Begin VB.Label Puntos 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   7200
      TabIndex        =   32
      Top             =   8535
      Width           =   270
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   3
      Left            =   7770
      MouseIcon       =   "frmCrearPersonaje.frx":106B
      MousePointer    =   99  'Custom
      Top             =   3465
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   5
      Left            =   7785
      MouseIcon       =   "frmCrearPersonaje.frx":11BD
      MousePointer    =   99  'Custom
      Top             =   3675
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   7
      Left            =   7785
      MouseIcon       =   "frmCrearPersonaje.frx":130F
      MousePointer    =   99  'Custom
      Top             =   3885
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   9
      Left            =   7770
      MouseIcon       =   "frmCrearPersonaje.frx":1461
      MousePointer    =   99  'Custom
      Top             =   4110
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   11
      Left            =   7770
      MouseIcon       =   "frmCrearPersonaje.frx":15B3
      MousePointer    =   99  'Custom
      Top             =   4335
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   13
      Left            =   7755
      MouseIcon       =   "frmCrearPersonaje.frx":1705
      MousePointer    =   99  'Custom
      Top             =   4545
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   15
      Left            =   7770
      MouseIcon       =   "frmCrearPersonaje.frx":1857
      MousePointer    =   99  'Custom
      Top             =   4785
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   17
      Left            =   7770
      MouseIcon       =   "frmCrearPersonaje.frx":19A9
      MousePointer    =   99  'Custom
      Top             =   4965
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   19
      Left            =   7770
      MouseIcon       =   "frmCrearPersonaje.frx":1AFB
      MousePointer    =   99  'Custom
      Top             =   5175
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   21
      Left            =   7770
      MouseIcon       =   "frmCrearPersonaje.frx":1C4D
      MousePointer    =   99  'Custom
      Top             =   5385
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   23
      Left            =   7755
      MouseIcon       =   "frmCrearPersonaje.frx":1D9F
      MousePointer    =   99  'Custom
      Top             =   5610
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   25
      Left            =   7755
      MouseIcon       =   "frmCrearPersonaje.frx":1EF1
      MousePointer    =   99  'Custom
      Top             =   5820
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   27
      Left            =   7755
      MouseIcon       =   "frmCrearPersonaje.frx":2043
      MousePointer    =   99  'Custom
      Top             =   6015
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   1
      Left            =   7770
      MouseIcon       =   "frmCrearPersonaje.frx":2195
      MousePointer    =   99  'Custom
      Top             =   3240
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   0
      Left            =   8355
      MouseIcon       =   "frmCrearPersonaje.frx":22E7
      MousePointer    =   99  'Custom
      Top             =   3270
      Width           =   180
   End
   Begin VB.Image command1 
      Height          =   150
      Index           =   2
      Left            =   8355
      MouseIcon       =   "frmCrearPersonaje.frx":2439
      MousePointer    =   99  'Custom
      Top             =   3495
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   4
      Left            =   8370
      MouseIcon       =   "frmCrearPersonaje.frx":258B
      MousePointer    =   99  'Custom
      Top             =   3720
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   150
      Index           =   6
      Left            =   8370
      MouseIcon       =   "frmCrearPersonaje.frx":26DD
      MousePointer    =   99  'Custom
      Top             =   3945
      Width           =   180
   End
   Begin VB.Image command1 
      Height          =   150
      Index           =   8
      Left            =   8370
      MouseIcon       =   "frmCrearPersonaje.frx":282F
      MousePointer    =   99  'Custom
      Top             =   4155
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   10
      Left            =   8370
      MouseIcon       =   "frmCrearPersonaje.frx":2981
      MousePointer    =   99  'Custom
      Top             =   4380
      Width           =   165
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   12
      Left            =   8355
      MouseIcon       =   "frmCrearPersonaje.frx":2AD3
      MousePointer    =   99  'Custom
      Top             =   4605
      Width           =   165
   End
   Begin VB.Image command1 
      Height          =   120
      Index           =   14
      Left            =   8370
      MouseIcon       =   "frmCrearPersonaje.frx":2C25
      MousePointer    =   99  'Custom
      Top             =   4800
      Width           =   135
   End
   Begin VB.Image command1 
      Height          =   120
      Index           =   16
      Left            =   8370
      MouseIcon       =   "frmCrearPersonaje.frx":2D77
      MousePointer    =   99  'Custom
      Top             =   4995
      Width           =   135
   End
   Begin VB.Image command1 
      Height          =   120
      Index           =   18
      Left            =   8370
      MouseIcon       =   "frmCrearPersonaje.frx":2EC9
      MousePointer    =   99  'Custom
      Top             =   5220
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   20
      Left            =   8370
      MouseIcon       =   "frmCrearPersonaje.frx":301B
      MousePointer    =   99  'Custom
      Top             =   5400
      Width           =   165
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   22
      Left            =   8370
      MouseIcon       =   "frmCrearPersonaje.frx":316D
      MousePointer    =   99  'Custom
      Top             =   5640
      Width           =   165
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   24
      Left            =   8370
      MouseIcon       =   "frmCrearPersonaje.frx":32BF
      MousePointer    =   99  'Custom
      Top             =   5850
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   120
      Index           =   26
      Left            =   8370
      MouseIcon       =   "frmCrearPersonaje.frx":3411
      MousePointer    =   99  'Custom
      Top             =   6075
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   150
      Index           =   28
      Left            =   8370
      MouseIcon       =   "frmCrearPersonaje.frx":3563
      MousePointer    =   99  'Custom
      Top             =   6285
      Width           =   165
   End
   Begin VB.Image command1 
      Height          =   150
      Index           =   29
      Left            =   7755
      MouseIcon       =   "frmCrearPersonaje.frx":36B5
      MousePointer    =   99  'Custom
      Top             =   6270
      Width           =   165
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   30
      Left            =   8370
      MouseIcon       =   "frmCrearPersonaje.frx":3807
      MousePointer    =   99  'Custom
      Top             =   6495
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   150
      Index           =   31
      Left            =   7770
      MouseIcon       =   "frmCrearPersonaje.frx":3959
      MousePointer    =   99  'Custom
      Top             =   6465
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   32
      Left            =   8370
      MouseIcon       =   "frmCrearPersonaje.frx":3AAB
      MousePointer    =   99  'Custom
      Top             =   6720
      Width           =   180
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   33
      Left            =   7755
      MouseIcon       =   "frmCrearPersonaje.frx":3BFD
      MousePointer    =   99  'Custom
      Top             =   6690
      Width           =   180
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   34
      Left            =   8370
      MouseIcon       =   "frmCrearPersonaje.frx":3D4F
      MousePointer    =   99  'Custom
      Top             =   6945
      Width           =   135
   End
   Begin VB.Image command1 
      Height          =   150
      Index           =   35
      Left            =   7755
      MouseIcon       =   "frmCrearPersonaje.frx":3EA1
      MousePointer    =   99  'Custom
      Top             =   6915
      Width           =   165
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   36
      Left            =   8370
      MouseIcon       =   "frmCrearPersonaje.frx":3FF3
      MousePointer    =   99  'Custom
      Top             =   7170
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   180
      Index           =   37
      Left            =   7755
      MouseIcon       =   "frmCrearPersonaje.frx":4145
      MousePointer    =   99  'Custom
      Top             =   7125
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   120
      Index           =   38
      Left            =   8370
      MouseIcon       =   "frmCrearPersonaje.frx":4297
      MousePointer    =   99  'Custom
      Top             =   7395
      Width           =   135
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   39
      Left            =   7770
      MouseIcon       =   "frmCrearPersonaje.frx":43E9
      MousePointer    =   99  'Custom
      Top             =   7335
      Width           =   165
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   40
      Left            =   8370
      MouseIcon       =   "frmCrearPersonaje.frx":453B
      MousePointer    =   99  'Custom
      Top             =   7590
      Width           =   165
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   41
      Left            =   7815
      MouseIcon       =   "frmCrearPersonaje.frx":468D
      MousePointer    =   99  'Custom
      Top             =   7560
      Width           =   135
   End
   Begin VB.Image boton 
      Height          =   645
      Index           =   2
      Left            =   315
      MouseIcon       =   "frmCrearPersonaje.frx":47DF
      MousePointer    =   99  'Custom
      Top             =   3015
      Width           =   660
   End
   Begin VB.Image boton 
      Height          =   255
      Index           =   1
      Left            =   1905
      MouseIcon       =   "frmCrearPersonaje.frx":4931
      MousePointer    =   99  'Custom
      Top             =   8565
      Width           =   765
   End
   Begin VB.Image boton 
      Height          =   210
      Index           =   0
      Left            =   9075
      MouseIcon       =   "frmCrearPersonaje.frx":4A83
      MousePointer    =   99  'Custom
      Top             =   8565
      Width           =   1200
   End
   Begin VB.Label Skill 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   20
      Left            =   7995
      TabIndex        =   28
      Top             =   7515
      Width           =   270
   End
   Begin VB.Label Skill 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   19
      Left            =   7995
      TabIndex        =   27
      Top             =   7306
      Width           =   270
   End
   Begin VB.Label Skill 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   18
      Left            =   7995
      TabIndex        =   26
      Top             =   7092
      Width           =   270
   End
   Begin VB.Label Skill 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   17
      Left            =   7995
      TabIndex        =   25
      Top             =   6878
      Width           =   270
   End
   Begin VB.Label Skill 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   16
      Left            =   7995
      TabIndex        =   24
      Top             =   6664
      Width           =   270
   End
   Begin VB.Label Skill 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   15
      Left            =   7995
      TabIndex        =   23
      Top             =   6450
      Width           =   270
   End
   Begin VB.Label Skill 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   14
      Left            =   7995
      TabIndex        =   22
      Top             =   6236
      Width           =   270
   End
   Begin VB.Label Skill 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   13
      Left            =   7995
      TabIndex        =   21
      Top             =   6022
      Width           =   270
   End
   Begin VB.Label Skill 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   12
      Left            =   7995
      TabIndex        =   20
      Top             =   5808
      Width           =   270
   End
   Begin VB.Label Skill 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   11
      Left            =   7995
      TabIndex        =   19
      Top             =   5594
      Width           =   270
   End
   Begin VB.Label Skill 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   10
      Left            =   7995
      TabIndex        =   18
      Top             =   5380
      Width           =   270
   End
   Begin VB.Label Skill 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   9
      Left            =   7995
      TabIndex        =   17
      Top             =   5166
      Width           =   270
   End
   Begin VB.Label Skill 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   8
      Left            =   7995
      TabIndex        =   16
      Top             =   4952
      Width           =   270
   End
   Begin VB.Label Skill 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   7
      Left            =   7995
      TabIndex        =   15
      Top             =   4738
      Width           =   270
   End
   Begin VB.Label Skill 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   6
      Left            =   7995
      TabIndex        =   14
      Top             =   4524
      Width           =   270
   End
   Begin VB.Label Skill 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   5
      Left            =   7995
      TabIndex        =   13
      Top             =   4310
      Width           =   270
   End
   Begin VB.Label Skill 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   7995
      TabIndex        =   12
      Top             =   4096
      Width           =   270
   End
   Begin VB.Label Skill 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   7995
      TabIndex        =   11
      Top             =   3882
      Width           =   270
   End
   Begin VB.Label Skill 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   7995
      TabIndex        =   10
      Top             =   3668
      Width           =   270
   End
   Begin VB.Label Skill 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   7995
      TabIndex        =   9
      Top             =   3240
      Width           =   270
   End
   Begin VB.Label Skill 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   7995
      TabIndex        =   8
      Top             =   3450
      Width           =   270
   End
   Begin VB.Image imgHogar 
      Height          =   2850
      Left            =   8775
      Picture         =   "frmCrearPersonaje.frx":4BD5
      Top             =   3900
      Width           =   2985
   End
   Begin VB.Label lbCarisma 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "11"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   5310
      TabIndex        =   6
      Top             =   3570
      Width           =   225
   End
   Begin VB.Label lbSabiduria 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Left            =   3840
      TabIndex        =   5
      Top             =   4260
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Label lbInteligencia 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "5"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   5310
      TabIndex        =   4
      Top             =   3240
      Width           =   210
   End
   Begin VB.Label lbConstitucion 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "8"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   5310
      TabIndex        =   3
      Top             =   3900
      Width           =   225
   End
   Begin VB.Label lbAgilidad 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "10"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   5310
      TabIndex        =   2
      Top             =   2910
      Width           =   225
   End
   Begin VB.Label lbFuerza 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "12"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   5310
      TabIndex        =   1
      Top             =   2580
      Width           =   210
   End
End
Attribute VB_Name = "frmCrearPersonaje"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Argentum Online 0.9.0.9
'
'Copyright (C) 2002 Marquez Pablo Ignacio
'Copyright (C) 2002 Otto Perez
'Copyright (C) 2002 Aaron Perkins
'Copyright (C) 2002 Mataas Fernando Pequeao
'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation; either version 2 of the License, or
'any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License for more details.
'
'You should have received a copy of the GNU General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'
'You can contact me at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 namero 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Cadigo Postal 1900
'Pablo Ignacio Marquez

Option Explicit

Public SkillPoints As Byte

'--- CABEZAS ---
Private PrimerCabeza As Integer
Private UltimaCabeza As Integer
Private UserDir As Byte
'---------------

Function CheckData() As Boolean
If UserRaza = "" Then
    MsgBox "Seleccione la raza del personaje."
    Exit Function
End If

If UserSexo = "" Then
    MsgBox "Seleccione el sexo del personaje."
    Exit Function
End If

If UserClase = "" Then
    MsgBox "Seleccione la clase del personaje."
    Exit Function
End If

If UserClase = "" Then
    MsgBox "Seleccione la clase del personaje."
    Exit Function
End If

If UserHogar = "" Then
    MsgBox "Seleccione el hogar del personaje."
    Exit Function
End If

If SkillPoints > 0 Then
    MsgBox "Asigne los skillpoints del personaje."
    Exit Function
End If

Dim i As Integer
For i = 1 To NUMATRIBUTOS
    If UserAtributos(i) = 0 Then
        MsgBox "Los atributos del personaje son invalidos."
        Exit Function
    End If
Next i

CheckData = True


End Function

Private Sub boton_Click(Index As Integer)

Call PlayWaveDS(SND_CLICK)

Select Case Index
    Case 0
        
        Dim i As Integer
        Dim k As Object
        i = 1
        For Each k In Skill
            UserSkills(i) = k.Caption
            i = i + 1
        Next
        
        UserName = txtNombre.Text
        
        If Right$(UserName, 1) = " " Then
                UserName = RTrim$(UserName)
                MsgBox "Nombre invalido, se han removido los espacios al final del nombre"
        End If
        
        UserRaza = lstRaza.List(lstRaza.ListIndex)
        UserSexo = lstGenero.List(lstGenero.ListIndex)
        UserClase = lstProfesion.List(lstProfesion.ListIndex)
        
        UserAtributos(1) = Val(lbFuerza.Caption)
        UserAtributos(2) = Val(lbInteligencia.Caption)
        UserAtributos(3) = Val(lbAgilidad.Caption)
        UserAtributos(4) = Val(lbCarisma.Caption)
        UserAtributos(5) = Val(lbConstitucion.Caption)
        
        UserHogar = lstHogar.List(lstHogar.ListIndex)
        
        'Barrin 3/10/03
        If CheckData() Then
            frmMain.Socket1.HostName = CurServerIp
            frmMain.Socket1.RemotePort = CurServerPort
            
            Me.MousePointer = 11
            EstadoLogin = CrearNuevoPj
         
            If Not frmMain.Socket1.Connected Then
                
                frmMensaje.Show
                    frmMensaje.msg.Caption = "Error: Se ha perdido la conexion con el server."
                Unload Me
            Else
                Call Login
            End If
        End If
        
    Case 1
        If Musica = 0 Then
            CurMidi = DirMidi & MIdi_Inicio & ".mid"
            LoopMidi = 1
            Call CargarMIDI(CurMidi)
            Call Play_Midi
        End If
        
        frmConnect.Picture = LoadPicture(App.Path & "\Graficos\conectar.png")
        Me.Visible = False
        
        
    Case 2
        Call PlayWaveDS(SND_DICE)
        Call TirarDados
      
End Select


End Sub


Function RandomNumber(ByVal LowerBound As Variant, ByVal UpperBound As Variant) As Single

Randomize Timer

RandomNumber = (UpperBound - LowerBound + 1) * Rnd + LowerBound
If RandomNumber > UpperBound Then RandomNumber = UpperBound

End Function


Private Sub TirarDados()
lbFuerza.Caption = CInt(RandomNumber(1, 6) + RandomNumber(1, 6) + RandomNumber(1, 6))
lbInteligencia.Caption = CInt(RandomNumber(1, 6) + RandomNumber(1, 6) + RandomNumber(1, 6))
lbAgilidad.Caption = CInt(RandomNumber(1, 6) + RandomNumber(1, 6) + RandomNumber(1, 6))
lbCarisma.Caption = CInt(RandomNumber(1, 6) + RandomNumber(1, 6) + RandomNumber(1, 6))
lbConstitucion.Caption = CInt(RandomNumber(1, 6) + RandomNumber(1, 6) + RandomNumber(1, 6))
End Sub

Private Sub Command1_Click(Index As Integer)
Call PlayWaveDS(SND_CLICK)

Dim indice
If Index Mod 2 = 0 Then
    If SkillPoints > 0 Then
        indice = Index \ 2
        Skill(indice).Caption = Val(Skill(indice).Caption) + 1
        SkillPoints = SkillPoints - 1
    End If
Else
    If SkillPoints < 10 Then
        
        indice = Index \ 2
        If Val(Skill(indice).Caption) > 0 Then
            Skill(indice).Caption = Val(Skill(indice).Caption) - 1
            SkillPoints = SkillPoints + 1
        End If
    End If
End If

Puntos.Caption = SkillPoints
End Sub

Private Sub Form_Load()
    ' [CODE] - Inicializacion de AoSpain
    SkillPoints = 10
    Puntos.Caption = SkillPoints
    Me.Picture = LoadPicture(App.Path & "\graficos\CP-Interface.jpg")
    imgHogar.Picture = LoadPicture(App.Path & "\graficos\CP-Ullathorpe.jpg")

    Dim i As Integer
    lstProfesion.Clear
    For i = LBound(ListaClases) To UBound(ListaClases)
        lstProfesion.AddItem ListaClases(i)
    Next i

    ' Inicializar variables globales de creacion
    UserSexo = ""
    UserRaza = ""
    UserHogar = ""
    UserHead = 0
    UserDir = SOUTH 'Mirando abajo por defecto

    ' Limpiar PictureBox de vista previa
    picPJ.Cls

    ' Configurar listas
    lstProfesion.ListIndex = 0
    lstGenero.ListIndex = 0
    lstRaza.ListIndex = 0
    lstHogar.ListIndex = 0

    '--- INICIALIZAR CABEZAS Y VISTA PREVIA ---
    Call LoadHeadRanges
    '---------------------------

    Image1.Picture = LoadPicture(App.Path & "\graficos\" & lstProfesion.Text & ".jpg")
    Call TirarDados
End Sub


Private Sub lstProfesion_Click()
On Error Resume Next
Image1.Picture = LoadPicture(App.Path & "\graficos\" & lstProfesion.Text & ".jpg")
End Sub

Private Sub txtNombre_Change()
txtNombre.Text = LTrim(txtNombre.Text)
End Sub

Private Sub txtNombre_GotFocus()
MsgBox "Sea cuidadoso al seleccionar el nombre de su personaje, Argentum es un juego de rol, un mundo magico y fantastico, si selecciona un nombre obsceno o con connotacian politica los administradores borraran su personaje y no habra ninguna posibilidad de recuperarlo."
End Sub

Private Sub txtNombre_KeyPress(KeyAscii As Integer)
 KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

'==========================================================
' SISTEMA DE CABEZAS DINAMICO (NADA HARCODEADO)
'==========================================================

Private Sub LoadHeadRanges()
    Dim Section As String
    Dim RazaStr As String
    Dim SexoStr As String
    
    RazaStr = Replace$(lstRaza.List(lstRaza.ListIndex), " ", "")
    SexoStr = lstGenero.List(lstGenero.ListIndex)
    
    Section = UCase$(RazaStr) & "-" & UCase$(SexoStr)
    
    PrimerCabeza = Val(GetVar(App.Path & "\init\Cabezas.ini", Section, "PrimerCabeza"))
    UltimaCabeza = Val(GetVar(App.Path & "\init\Cabezas.ini", Section, "UltimaCabeza"))
    
    If PrimerCabeza = 0 Then PrimerCabeza = 1
    If UltimaCabeza = 0 Then UltimaCabeza = 1
    
    UserHead = PrimerCabeza
    
    Call UpdateHeadSelection
End Sub

Private Sub UpdateHeadSelection()
    Dim Head As Integer
    Dim i As Integer
    
    'Limpiar PictureBoxes si es necesario (opcional si se usa AutoRedraw)
    For i = 0 To 4
        picHead(i).Cls
    Next i
    
    'Mostramos 5 cabezas en el carrusel
    'picHead(2) es la seleccionada (centro)
    
    Call DrawHead(UserHead, 2)
    
    'Derecha
    Head = UserHead
    For i = 3 To 4
        Head = CheckCabeza(Head + 1)
        Call DrawHead(Head, i)
    Next i
    
    'Izquierda
    Head = UserHead
    For i = 1 To 0 Step -1
        Head = CheckCabeza(Head - 1)
        Call DrawHead(Head, i)
    Next i
    
    'Actualizamos la vista previa del PJ completo
    Call UpdatePJPreview
End Sub

Private Sub DrawHead(ByVal HeadIdx As Integer, ByVal PicIdx As Integer)
    Dim SR As RECT, DR As RECT
    Dim Grh As Long
    
    ' Limpiar el PictureBox antes de intentar dibujar
    picHead(PicIdx).Cls
    
    ' Prevenir errores si el array no se cargó o el índice es inválido
    On Error GoTo ErrorHandler
    
    ' Si HeadIdx es inválido, no dibujamos y salimos
    If HeadIdx < PrimerCabeza Or HeadIdx > UltimaCabeza Then Exit Sub
    
    ' Validar dirección
    If UserDir < 1 Or UserDir > 4 Then UserDir = 3
    
    ' Grh index de la cabeza en la direccion actual
    ' Usamos el array Miscabezas (el oficial del motor de AoSpain)
    Grh = Miscabezas(HeadIdx).Head(UserDir)
    
    ' Si no hay grafico para esta direccion, probamos con el frente (Sur = 3)
    If Grh <= 0 Then Grh = Miscabezas(HeadIdx).Head(3)
    
    ' Si sigue siendo invalido, salimos
    If Grh <= 0 Then Exit Sub

    ' Source Rect (Todo el GRH)
    With SR
        .Right = GrhData(Grh).pixelWidth
        .Bottom = GrhData(Grh).pixelHeight
    End With

    ' Dest Rect (Centrado en el PictureBox)
    With DR
        .Right = picHead(PicIdx).ScaleWidth
        .Bottom = picHead(PicIdx).ScaleHeight
    End With
    
    ' Dibujamos usando el motor de AoSpain
    Call DrawGrhtoHdc(picHead(PicIdx).hwnd, picHead(PicIdx).Hdc, Grh, SR, DR)
    picHead(PicIdx).Refresh
    
    Exit Sub

ErrorHandler:
    ' Error silencioso para evitar cierres inesperados
End Sub

Private Function CheckCabeza(ByVal HeadIdx As Integer) As Integer
    ' [CODE] - Logica de envoltura robusta para el carrusel de cabezas
    Dim TotalHeads As Integer
    TotalHeads = UltimaCabeza - PrimerCabeza + 1
    
    If TotalHeads <= 0 Then
        CheckCabeza = PrimerCabeza
        Exit Function
    End If
    
    If HeadIdx > UltimaCabeza Then
        ' Envolver hacia adelante
        CheckCabeza = PrimerCabeza + ((HeadIdx - PrimerCabeza) Mod TotalHeads)
    ElseIf HeadIdx < PrimerCabeza Then
        ' Envolver hacia atras
        CheckCabeza = UltimaCabeza - ((UltimaCabeza - HeadIdx) Mod TotalHeads)
    Else
        ' Dentro del rango
        CheckCabeza = HeadIdx
    End If
End Function

Private Sub HeadPJ_Click(Index As Integer)
    If Index = 1 Then 'Siguiente
        UserHead = CheckCabeza(UserHead + 1)
    Else 'Anterior
        UserHead = CheckCabeza(UserHead - 1)
    End If
    
    Call UpdateHeadSelection
End Sub

Private Sub DirPJ_Click(Index As Integer)
    If Index = 0 Then 'Girar derecha (sentido horario)
        UserDir = UserDir + 1
        If UserDir > 4 Then UserDir = 1
    Else 'Girar izquierda
        UserDir = UserDir - 1
        If UserDir < 1 Then UserDir = 4
    End If
    
    Call UpdateHeadSelection
End Sub

Private Sub lstRaza_Click()
    UserRaza = lstRaza.List(lstRaza.ListIndex)
    Call LoadHeadRanges
End Sub

Private Sub lstGenero_Click()
    UserSexo = lstGenero.List(lstGenero.ListIndex)
    Call LoadHeadRanges
End Sub

Private Sub UpdatePJPreview()
    ' [CODE] - Implementacion del sistema de cuerpos y cabezas dinámico para AoSpain
    On Error Resume Next
    
    Dim Body As Integer
    Dim Head As Integer
    Dim Grh As Grh
    Dim Pos As Integer
    Dim YYY As Integer
    Dim XBody As Integer, YBody As Integer, BHead As Integer
    Dim bVisible As Boolean
    Dim i As Integer
    
    ' Obtenemos el cuerpo segun raza y sexo (IDs oficiales de AoSpain)
    Body = GetBodyByRaceSex(lstRaza.List(lstRaza.ListIndex), lstGenero.List(lstGenero.ListIndex))
    Head = UserHead
    
    ' Definimos visibilidad (si hay cuerpo y cabeza validos)
    bVisible = (Body > 0 And Head > 0)
    
    ' Gestionar visibilidad de controles (Botones de cabezas y direccion)
    HeadPJ(0).Visible = bVisible
    HeadPJ(1).Visible = bVisible
    DirPJ(0).Visible = bVisible
    DirPJ(1).Visible = bVisible
    
    ' Gestionar visibilidad de PictureBoxes
    picPJ.Visible = bVisible
    picTemp.Visible = bVisible
    
    For i = 0 To 4
        picHead(i).Visible = bVisible
    Next i
    
    ' Gestionar visibilidad de las lineas decorativas
    For i = 0 To 3
        Line1(i).Visible = bVisible
    Next i
    
    ' Si no hay datos validos, no dibujamos nada mas
    If Not bVisible Then Exit Sub
    
    ' Limpiamos el PictureBox de vista previa
    picPJ.Cls
    
    ' Validar dirección (1=Norte, 2=Este, 3=Sur, 4=Oeste)
    If UserDir < 1 Or UserDir > 4 Then UserDir = 3
    
    ' Coordenadas de dibujado base para picPJ (41x65 pixels)
    XBody = 15
    YBody = 25
    BHead = 21
    
    ' Dibujamos el Cuerpo (usamos el primer frame de la animacion de caminar)
    Grh = BodyData(Body).Walk(UserDir)
    If Grh.GrhIndex > 0 Then
        ' Dibujamos cuerpo
        Call GrhRenderToHdc(GrhData(Grh.GrhIndex).Frames(1), picPJ.Hdc, XBody, YBody, True)
        
        ' Calculamos posicion de la cabeza usando el offset del cuerpo
        YYY = BodyData(Body).HeadOffset.Y
        Pos = YYY + GrhData(GrhData(Grh.GrhIndex).Frames(1)).pixelHeight
        
        ' Dibujamos la Cabeza
        Grh = HeadData(Head).Head(UserDir)
        If Grh.GrhIndex > 0 Then
            Call GrhRenderToHdc(GrhData(Grh.GrhIndex).Frames(1), picPJ.Hdc, BHead, Pos + 7, True)
        End If
    End If
    
    ' Refrescamos el PictureBox
    picPJ.Refresh
End Sub

Private Function GetBodyByRaceSex(ByVal Raza As String, ByVal Sexo As String) As Integer
    ' [CODE] - Mapeo de cuerpos iniciales segun configuracion del servidor AoSpain
    Select Case UCase$(Sexo)
        Case "HOMBRE"
            Select Case UCase$(Raza)
                Case "HUMANO": GetBodyByRaceSex = 1
                Case "ELFO": GetBodyByRaceSex = 2
                Case "ELFO OSCURO": GetBodyByRaceSex = 3
                Case "ENANO", "GNOMO": GetBodyByRaceSex = 52
                Case "ORCO": GetBodyByRaceSex = 232
                Case "HOBBIT": GetBodyByRaceSex = 239
                Case Else: GetBodyByRaceSex = 1
            End Select
        Case "MUJER"
            Select Case UCase$(Raza)
                Case "HUMANO": GetBodyByRaceSex = 1
                Case "ELFO": GetBodyByRaceSex = 2
                Case "ELFO OSCURO": GetBodyByRaceSex = 3
                Case "ENANO", "GNOMO": GetBodyByRaceSex = 52
                Case "ORCO": GetBodyByRaceSex = 234
                Case "HOBBIT": GetBodyByRaceSex = 240
                Case Else: GetBodyByRaceSex = 1
            End Select
        Case Else
            GetBodyByRaceSex = 1
    End Select
End Function

