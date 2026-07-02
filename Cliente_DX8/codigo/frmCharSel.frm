VERSION 5.00
Begin VB.Form frmCharSel 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "AoSpain - Selección de Personaje"
   ClientHeight    =   4005
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5415
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4005
   ScaleWidth      =   5415
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstChars 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2100
      Left            =   480
      TabIndex        =   0
      Top             =   720
      Width           =   4455
   End
   Begin VB.CommandButton cmdConectar 
      Caption         =   "ENTRAR AL MUNDO"
      Height          =   495
      Left            =   480
      TabIndex        =   1
      Top             =   3120
      Width           =   2055
   End
   Begin VB.CommandButton cmdCrearPJ 
      Caption         =   "NUEVO PERSONAJE"
      Height          =   495
      Left            =   2880
      TabIndex        =   2
      Top             =   3120
      Width           =   2055
   End
   Begin VB.CommandButton cmdBorrarPJ 
      Caption         =   "BORRAR PERSONAJE"
      Height          =   495
      Left            =   1680
      TabIndex        =   4
      Top             =   3720
      Width           =   2055
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "TUS PERSONAJES"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   480
      TabIndex        =   3
      Top             =   240
      Width           =   4455
   End
End
Attribute VB_Name = "frmCharSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdConectar_Click()
    If lstChars.ListIndex <> -1 Then
        ' Enviamos el slot seleccionado (sumamos 1 porque ListIndex empieza en 0)
        Call EnviarSeleccionarPJ(lstChars.ListIndex + 1)
    Else
        MsgBox "Selecciona un personaje de la lista."
    End If
End Sub

Private Sub cmdCrearPJ_Click()
    ' Lógica para abrir la creación de personaje si el slot está vacío o si se desea uno nuevo
    frmCrearPersonaje.Show
    Unload Me
End Sub

Private Sub cmdBorrarPJ_Click()
    If lstChars.ListIndex <> -1 Then
        If lstChars.Text <> "(Slot Vacío)" Then
            If MsgBox("¿Estás seguro de que deseas borrar a " & lstChars.Text & "? Se enviará un código de confirmación a tu email.", vbYesNo + VBExclamation) = vbYes Then
                ' Solicitamos el borrado al servidor
                Call EnviarSolicitudBorrado(lstChars.ListIndex + 1)
            End If
        End If
    End If
End Sub
