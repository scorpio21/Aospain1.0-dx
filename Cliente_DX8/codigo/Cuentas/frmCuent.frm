VERSION 5.00
Begin VB.Form frmCuent 
   BackColor       =   &H80000007&
   BorderStyle     =   0  'None
   Caption         =   "Tierras Perdidas AO"
   ClientHeight    =   9015
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12015
   BeginProperty Font 
      Name            =   "Georgia"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCuent.frx":0000
   LinkTopic       =   "Form1"
   MouseIcon       =   "frmCuent.frx":000C
   MousePointer    =   99  'Custom
   Picture         =   "frmCuent.frx":0CD6
   ScaleHeight     =   9015
   ScaleWidth      =   12015
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer AnimationTimer 
      Interval        =   100
      Left            =   120
      Top             =   120
   End
   Begin VB.PictureBox PJ 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1600
      Index           =   9
      Left            =   8345
      MouseIcon       =   "frmCuent.frx":6EB3E
      MousePointer    =   99  'Custom
      Picture         =   "frmCuent.frx":6F808
      ScaleHeight     =   1605
      ScaleWidth      =   1245
      TabIndex        =   47
      Top             =   4920
      Width           =   1250
      Begin VB.Label GM 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "GM"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFC0&
         Height          =   195
         Index           =   9
         Left            =   0
         TabIndex        =   51
         Top             =   720
         Visible         =   0   'False
         Width           =   1268
      End
      Begin VB.Label CP 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Crear Personaje"
         BeginProperty Font 
            Name            =   "Candara"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Index           =   9
         Left            =   0
         TabIndex        =   50
         Top             =   600
         Width           =   1268
      End
      Begin VB.Label nombre 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Nada"
         BeginProperty Font 
            Name            =   "Candara"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   9
         Left            =   0
         TabIndex        =   49
         Top             =   1080
         Visible         =   0   'False
         Width           =   1268
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Nivel: 0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   9
         Left            =   0
         TabIndex        =   48
         Top             =   1320
         Visible         =   0   'False
         Width           =   1268
      End
   End
   Begin VB.PictureBox PJ 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1600
      Index           =   8
      Left            =   6840
      MouseIcon       =   "frmCuent.frx":6FAA3
      MousePointer    =   99  'Custom
      Picture         =   "frmCuent.frx":7076D
      ScaleHeight     =   1605
      ScaleWidth      =   1245
      TabIndex        =   42
      Top             =   4870
      Width           =   1250
      Begin VB.Label GM 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "GM"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFC0&
         Height          =   195
         Index           =   8
         Left            =   0
         TabIndex        =   46
         Top             =   720
         Visible         =   0   'False
         Width           =   1268
      End
      Begin VB.Label CP 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Crear Personaje"
         BeginProperty Font 
            Name            =   "Candara"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Index           =   8
         Left            =   0
         TabIndex        =   45
         Top             =   600
         Width           =   1268
      End
      Begin VB.Label nombre 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Nada"
         BeginProperty Font 
            Name            =   "Candara"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   8
         Left            =   0
         TabIndex        =   44
         Top             =   1080
         Visible         =   0   'False
         Width           =   1268
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Nivel: 0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   8
         Left            =   0
         TabIndex        =   43
         Top             =   1320
         Visible         =   0   'False
         Width           =   1268
      End
   End
   Begin VB.PictureBox PJ 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1600
      Index           =   7
      Left            =   5280
      MouseIcon       =   "frmCuent.frx":70A08
      MousePointer    =   99  'Custom
      Picture         =   "frmCuent.frx":716D2
      ScaleHeight     =   1605
      ScaleWidth      =   1245
      TabIndex        =   37
      Top             =   4870
      Width           =   1250
      Begin VB.Label GM 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "GM"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFC0&
         Height          =   195
         Index           =   7
         Left            =   0
         TabIndex        =   41
         Top             =   720
         Visible         =   0   'False
         Width           =   1268
      End
      Begin VB.Label CP 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Crear Personaje"
         BeginProperty Font 
            Name            =   "Candara"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Index           =   7
         Left            =   0
         TabIndex        =   40
         Top             =   600
         Width           =   1268
      End
      Begin VB.Label nombre 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Nada"
         BeginProperty Font 
            Name            =   "Candara"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   7
         Left            =   0
         TabIndex        =   39
         Top             =   1080
         Visible         =   0   'False
         Width           =   1268
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Nivel: 0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   7
         Left            =   0
         TabIndex        =   38
         Top             =   1320
         Visible         =   0   'False
         Width           =   1268
      End
   End
   Begin VB.PictureBox PJ 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1600
      Index           =   6
      Left            =   3720
      MouseIcon       =   "frmCuent.frx":7196D
      MousePointer    =   99  'Custom
      Picture         =   "frmCuent.frx":72637
      ScaleHeight     =   1605
      ScaleWidth      =   1245
      TabIndex        =   32
      Top             =   4870
      Width           =   1250
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Nivel: 0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   6
         Left            =   0
         TabIndex        =   36
         Top             =   1320
         Visible         =   0   'False
         Width           =   1268
      End
      Begin VB.Label nombre 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Nada"
         BeginProperty Font 
            Name            =   "Candara"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   6
         Left            =   0
         TabIndex        =   35
         Top             =   1080
         Visible         =   0   'False
         Width           =   1268
      End
      Begin VB.Label CP 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Crear Personaje"
         BeginProperty Font 
            Name            =   "Candara"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Index           =   6
         Left            =   0
         TabIndex        =   34
         Top             =   600
         Width           =   1268
      End
      Begin VB.Label GM 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "GM"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFC0&
         Height          =   195
         Index           =   6
         Left            =   0
         TabIndex        =   33
         Top             =   720
         Visible         =   0   'False
         Width           =   1268
      End
   End
   Begin VB.PictureBox PJ 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1600
      Index           =   5
      Left            =   2240
      MouseIcon       =   "frmCuent.frx":728D2
      MousePointer    =   99  'Custom
      Picture         =   "frmCuent.frx":7359C
      ScaleHeight     =   1605
      ScaleWidth      =   1245
      TabIndex        =   27
      Top             =   4870
      Width           =   1250
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Nivel: 0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   5
         Left            =   0
         TabIndex        =   31
         Top             =   1320
         Visible         =   0   'False
         Width           =   1268
      End
      Begin VB.Label nombre 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Nada"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   5
         Left            =   0
         TabIndex        =   30
         Top             =   1080
         Visible         =   0   'False
         Width           =   1268
      End
      Begin VB.Label CP 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Crear Personaje"
         BeginProperty Font 
            Name            =   "Candara"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Index           =   5
         Left            =   0
         TabIndex        =   29
         Top             =   600
         Width           =   1268
      End
      Begin VB.Label GM 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "GM"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFC0&
         Height          =   195
         Index           =   5
         Left            =   0
         TabIndex        =   28
         Top             =   720
         Visible         =   0   'False
         Width           =   1268
      End
   End
   Begin VB.PictureBox PJ 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1600
      Index           =   4
      Left            =   8400
      MouseIcon       =   "frmCuent.frx":73837
      MousePointer    =   99  'Custom
      Picture         =   "frmCuent.frx":74501
      ScaleHeight     =   1605
      ScaleWidth      =   1245
      TabIndex        =   22
      Top             =   3055
      Width           =   1250
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Nivel: 0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   4
         Left            =   0
         TabIndex        =   26
         Top             =   1320
         Visible         =   0   'False
         Width           =   1268
      End
      Begin VB.Label nombre 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Nada"
         BeginProperty Font 
            Name            =   "Candara"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   4
         Left            =   0
         TabIndex        =   25
         Top             =   1080
         Visible         =   0   'False
         Width           =   1268
      End
      Begin VB.Label CP 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Crear Personaje"
         BeginProperty Font 
            Name            =   "Candara"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Index           =   4
         Left            =   0
         TabIndex        =   24
         Top             =   600
         Width           =   1268
      End
      Begin VB.Label GM 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "GM"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFC0&
         Height          =   195
         Index           =   4
         Left            =   0
         TabIndex        =   23
         Top             =   720
         Visible         =   0   'False
         Width           =   1268
      End
   End
   Begin VB.PictureBox PJ 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1600
      Index           =   3
      Left            =   6840
      MouseIcon       =   "frmCuent.frx":7479C
      MousePointer    =   99  'Custom
      Picture         =   "frmCuent.frx":75466
      ScaleHeight     =   1605
      ScaleWidth      =   1245
      TabIndex        =   17
      Top             =   3100
      Width           =   1250
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Nivel: 0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   3
         Left            =   0
         TabIndex        =   21
         Top             =   1320
         Visible         =   0   'False
         Width           =   1238
      End
      Begin VB.Label nombre 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Nada"
         BeginProperty Font 
            Name            =   "Candara"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   3
         Left            =   0
         TabIndex        =   20
         Top             =   1080
         Visible         =   0   'False
         Width           =   1238
      End
      Begin VB.Label CP 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Crear Personaje"
         BeginProperty Font 
            Name            =   "Candara"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Index           =   3
         Left            =   0
         TabIndex        =   19
         Top             =   600
         Width           =   1268
      End
      Begin VB.Label GM 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "GM"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFC0&
         Height          =   195
         Index           =   3
         Left            =   0
         TabIndex        =   18
         Top             =   720
         Visible         =   0   'False
         Width           =   1238
      End
   End
   Begin VB.PictureBox PJ 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1600
      Index           =   2
      Left            =   5280
      MouseIcon       =   "frmCuent.frx":75701
      MousePointer    =   99  'Custom
      Picture         =   "frmCuent.frx":763CB
      ScaleHeight     =   1605
      ScaleWidth      =   1245
      TabIndex        =   12
      Top             =   3055
      Width           =   1250
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Nivel: 0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   2
         Left            =   0
         TabIndex        =   16
         Top             =   1320
         Visible         =   0   'False
         Width           =   1268
      End
      Begin VB.Label nombre 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Nada"
         BeginProperty Font 
            Name            =   "Candara"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   2
         Left            =   0
         TabIndex        =   15
         Top             =   1080
         Visible         =   0   'False
         Width           =   1268
      End
      Begin VB.Label CP 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Crear Personaje"
         BeginProperty Font 
            Name            =   "Candara"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Index           =   2
         Left            =   0
         TabIndex        =   14
         Top             =   600
         Width           =   1268
      End
      Begin VB.Label GM 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "GM"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFC0&
         Height          =   195
         Index           =   2
         Left            =   0
         TabIndex        =   13
         Top             =   720
         Visible         =   0   'False
         Width           =   1268
      End
   End
   Begin VB.PictureBox PJ 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1600
      Index           =   1
      Left            =   3720
      MouseIcon       =   "frmCuent.frx":76666
      MousePointer    =   99  'Custom
      Picture         =   "frmCuent.frx":77330
      ScaleHeight     =   1605
      ScaleWidth      =   1245
      TabIndex        =   7
      Top             =   3055
      Width           =   1250
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Nivel: 0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   0
         TabIndex        =   11
         Top             =   1320
         Visible         =   0   'False
         Width           =   1268
      End
      Begin VB.Label nombre 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Nada"
         BeginProperty Font 
            Name            =   "Candara"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   0
         TabIndex        =   10
         Top             =   1080
         Visible         =   0   'False
         Width           =   1268
      End
      Begin VB.Label CP 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Crear Personaje"
         BeginProperty Font 
            Name            =   "Candara"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Index           =   1
         Left            =   0
         TabIndex        =   9
         Top             =   600
         Width           =   1268
      End
      Begin VB.Label GM 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "GM"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFC0&
         Height          =   195
         Index           =   1
         Left            =   0
         TabIndex        =   8
         Top             =   720
         Visible         =   0   'False
         Width           =   1268
      End
   End
   Begin VB.PictureBox PJ 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1600
      Index           =   0
      Left            =   2240
      MouseIcon       =   "frmCuent.frx":775CB
      MousePointer    =   99  'Custom
      Picture         =   "frmCuent.frx":78295
      ScaleHeight     =   1605
      ScaleWidth      =   1245
      TabIndex        =   2
      Top             =   3055
      Width           =   1250
      Begin VB.Label GM 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "GM"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFC0&
         Height          =   195
         Index           =   0
         Left            =   0
         TabIndex        =   6
         Top             =   720
         Visible         =   0   'False
         Width           =   1268
      End
      Begin VB.Label CP 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Crear Personaje"
         BeginProperty Font 
            Name            =   "Candara"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Index           =   0
         Left            =   0
         TabIndex        =   5
         Top             =   600
         Width           =   1268
      End
      Begin VB.Label nombre 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Nada"
         BeginProperty Font 
            Name            =   "Candara"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   4
         Top             =   1080
         Visible         =   0   'False
         Width           =   1268
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Nivel: 0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   3
         Top             =   1320
         Visible         =   0   'False
         Width           =   1268
      End
   End
   Begin VB.Image imgCambiarPass 
      Height          =   495
      Left            =   5400
      Top             =   8040
      Width           =   1575
   End
   Begin VB.Image imgBorrarCuenta 
      Height          =   495
      Left            =   11400
      Top             =   0
      Width           =   495
   End
   Begin VB.Image imgEntrar 
      Height          =   495
      Left            =   7560
      Top             =   8040
      Width           =   1575
   End
   Begin VB.Image imgBorrarPJ 
      Height          =   375
      Left            =   2880
      Top             =   8040
      Width           =   1695
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Accname"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   255
      Left            =   12840
      TabIndex        =   1
      Top             =   2400
      Width           =   2175
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "PJClick"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   495
      Left            =   3360
      TabIndex        =   0
      Top             =   1320
      Visible         =   0   'False
      Width           =   2895
   End
End
Attribute VB_Name = "frmCuent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private Sub CP_Click(Index As Integer)
Call PlayWaveDS(SND_CLICK)
If CP(Index).Visible = True Then
    Call PlayWaveDS(SND_CLICK)

' Estandarizado a 10 slots (Revision Fase 2)
If nombre(9).Caption <> "Nada" Then
    MsgBox "Tu cuenta ha llegado al maximo de personajes (10).", vbInformation
    Exit Sub
End If

    EstadoLogin = Dados
    frmCrearPersonaje.Show vbModal
    'Audio.StopWave
    Me.MousePointer = 11
    Exit Sub
End If
End Sub

Private Sub Form_Load()
'Me.Icon = LoadPicture(App.Path & "\Graficos\Icono.ico")
Dim cca As Integer
Me.Picture = LoadPicture(App.Path & "\Graficos\Principal\cuentas-copia.jpg")

Dim i As Integer
'Label3.Caption = nombrecuent

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim asd As Integer
    ' Estandarizado a 10 slots (Fase 2 - Revision)
    For asd = 0 To 9
        nombre(asd).ForeColor = vbWhite
        Label2(asd).ForeColor = vbWhite
    Next asd
End Sub

Private Sub imgBorrarPJ_Click()
    ' [CODE] - Fase 2: Refactorizacion
    Call PlayWaveDS(SND_CLICK)
    
    If PJClickeado = "Nada" Or PJClickeado = "" Then
        MsgBox "Seleccione un personaje para borrar.", vbExclamation
        Exit Sub
    End If
    
    If MsgBox("¿Estás seguro que deseas borrar a " & PJClickeado & "?", vbYesNo + vbQuestion) = vbYes Then
        Call SendData("BORR" & PJClickeado & "," & nombrecuent)
    End If
End Sub

Private Sub imgEntrar_Click()
    ' [CODE] - Fase 2: Refactorizacion
    Call PlayWaveDS(SND_CLICK)
    Call EntrarAlMundo
End Sub

Private Sub EntrarAlMundo()
    ' [CODE] - Centralizacion de logica de conexion para AoSpain
    If PJClickeado = "Nada" Or PJClickeado = "" Then
        frmMensaje.Show
        frmMensaje.msg.Caption = "Seleccione un personaje."
        Exit Sub
    End If

    ' Nos aseguramos de tener una conexion limpia
    If frmMain.Socket1.Connected Then
        frmMain.Socket1.Disconnect
        frmMain.Socket1.Cleanup
        DoEvents
    End If

    ' Configuramos el host y puerto
    If CurServer <> 0 Then
        frmMain.Socket1.HostAddress = ServersLst(CurServer).Ip
        frmMain.Socket1.RemotePort = ServersLst(CurServer).Puerto
    Else
        If IPdelServidor <> "" Then
            frmMain.Socket1.HostAddress = IPdelServidor
            frmMain.Socket1.RemotePort = PuertoDelServidor
        Else
            ' Fallback a localhost/config
            frmMain.Socket1.HostAddress = "127.0.0.1"
            frmMain.Socket1.RemotePort = Config_Inicio.Puerto
        End If
    End If

    ' Conectamos
    frmMain.Socket1.Connect

    ' Esperamos un momento a que conecte (max 3 segundos)
    Dim lWait As Long
    lWait = GetTickCount
    Do While Not frmMain.Socket1.Connected
        DoEvents
        If GetTickCount - lWait > 3000 Then
            MsgBox "No se ha podido establecer conexion con el servidor.", vbCritical
            Exit Sub
        End If
    Loop
     
    ' Enviamos paquete de entrada al mundo (Online Login)
    Call SendData("OOLOGI" & PJClickeado & "," & nombrecuent)
End Sub

Private Sub imgBorrarCuenta_Click()
    ' [CODE] - Boton Salir (X)
    Call PlayWaveDS(SND_CLICK)
    
    ' Desconectamos el socket si esta conectado
    If frmMain.Socket1.Connected Then
        frmMain.Socket1.Disconnect
        frmMain.Socket1.Cleanup
    End If
    
    ' Volvemos a la pantalla de conexion
    Unload Me
    frmConnect.Show
End Sub

Private Sub imgCambiarPass_Click()
    ' [CODE] - Fase 2: Refactorizacion
    Call PlayWaveDS(SND_CLICK)
    frmCambiarPass.Visible = True
    Call SendData("PEDPRE" & nombrecuent)
End Sub




Private Sub Label2_Click(Index As Integer)
    If Index > 9 Then Exit Sub ' Límite de 10 slots
    Call PlayWaveDS(SND_CLICK)
    PJClickeado = nombre(Index)
End Sub

Private Sub Label2_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Index > 9 Then Exit Sub
    nombre(Index).ForeColor = vbYellow
    Label2(Index).ForeColor = vbYellow
End Sub

Private Sub nombre_Click(Index As Integer)
    If Index > 9 Then Exit Sub
    Call PlayWaveDS(SND_CLICK)
    PJClickeado = nombre(Index)
End Sub

Private Sub nombre_dblClick(Index As Integer)
    If Index > 9 Then Exit Sub
    If PJClickeado = "Nada" Then Exit Sub
    Call PlayWaveDS(SND_CLICK)
    Call EntrarAlMundo
End Sub

Private Sub nombre_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Index > 9 Then Exit Sub
    nombre(Index).ForeColor = vbYellow
    Label2(Index).ForeColor = vbYellow
End Sub

Private Sub PJ_Click(Index As Integer)
    If Index > 9 Then Exit Sub
    
    If CP(Index).Visible = True Then
        Call PlayWaveDS(SND_CLICK)
        
        ' Verificar si ya tenemos el máximo de personajes (10)
        If nombre(9).Caption <> "Nada" Then
            MsgBox "Tu cuenta ha llegado al máximo de personajes (10).", vbInformation
            Exit Sub
        End If

        EstadoLogin = Dados
        frmCrearPersonaje.Show vbModal
        Me.MousePointer = 11
        Exit Sub
    End If
    
    PJClickeado = nombre(Index)
End Sub

Private Sub PJ_dblClick(Index As Integer)
    If Index > 9 Then Exit Sub
    If PJClickeado = "Nada" Then Exit Sub
    Call PlayWaveDS(SND_CLICK)
    Call EntrarAlMundo
End Sub

Private Sub AnimationTimer_Timer()
    Call RenderizarPJsCuentas
End Sub

Private Sub PJ_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Index > 9 Then Exit Sub
    Dim asd As Integer
    For asd = 0 To 9
        nombre(asd).ForeColor = vbWhite
        Label2(asd).ForeColor = vbWhite
    Next asd
End Sub
