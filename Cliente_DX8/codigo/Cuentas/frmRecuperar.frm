VERSION 5.00
Begin VB.Form frmRecuperar 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4890
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   4890
   ControlBox      =   0   'False
   Icon            =   "frmRecuperar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4890
   ScaleWidth      =   4890
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtNombre 
      BackColor       =   &H00000080&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   315
      Left            =   255
      TabIndex        =   2
      Top             =   1440
      Width           =   4350
   End
   Begin VB.TextBox txtMail 
      BackColor       =   &H00000080&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   315
      Left            =   255
      TabIndex        =   1
      Top             =   2280
      Width           =   4350
   End
   Begin VB.TextBox txtRespuesta 
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   240
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   3750
      Visible         =   0   'False
      Width           =   4455
   End
   Begin VB.Label txtPregunta 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "¿Tamaño de mi pija? "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   255
      TabIndex        =   3
      Top             =   3105
      Visible         =   0   'False
      Width           =   4455
   End
   Begin VB.Image Cancelar 
      Height          =   465
      Left            =   255
      Stretch         =   -1  'True
      Top             =   2640
      Width           =   1710
   End
   Begin VB.Image Siguiente 
      Height          =   465
      Left            =   2880
      Stretch         =   -1  'True
      Top             =   2640
      Width           =   1725
   End
   Begin VB.Image Recuperar 
      Height          =   525
      Left            =   300
      Top             =   4260
      Width           =   4455
   End
End
Attribute VB_Name = "frmRecuperar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Cancelar_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
PlayWaveDS (SND_CLICK)
If frmMain.Socket1.HostAddress = CurServerIp Or frmMain.Socket1.RemotePort = CurServerPort Then
            frmMain.Socket1.Disconnect
            frmMain.Socket1.Cleanup
            End If
            
Cancelar.Picture = LoadPicture(App.Path & "\Graficos\Principal\RecuperarCancelarA.jpg")
Unload Me
End Sub
Private Sub Form_Load()

Me.Picture = LoadPicture(App.Path & "\Graficos\Principal\Recuperar.jpg")
Cancelar.Picture = LoadPicture(App.Path & "\Graficos\Principal\RecuperarCancelar.jpg")
Siguiente.Picture = LoadPicture(App.Path & "\Graficos\Principal\RecuperarSiguiente.jpg")
Recuperar.Picture = LoadPicture(App.Path & "\Graficos\Principal\Recuperar2Fin.jpg")
Me.Height = 3345
txtPregunta.Visible = False


            If frmCuent.Visible = True Then
            Unload frmCuent
            End If
            
End Sub

Private Sub Recuperar_Click()
PlayWaveDS (SND_CLICK)
Call SendData("REECUU" & txtNombre & "," & txtRespuesta)
End Sub

Private Sub Recuperar_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Me.Recuperar.Picture = LoadPicture(App.Path & "\Graficos\Principal\Recuperar2FinA.jpg")
End Sub

Private Sub Recuperar_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Me.Recuperar.Picture = LoadPicture(App.Path & "\Graficos\Principal\Recuperar2Fin.jpg")
End Sub

Private Sub Siguiente_Click()

PlayWaveDS (SND_CLICK)


        
If txtNombre = "" Or txtMail = "" Then
MsgBox "¡Completa todo!"
Exit Sub
End If

Call SendData("RECCUU" & txtNombre & "," & txtMail)
End Sub

Private Sub Siguiente_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Siguiente.Picture = LoadPicture(App.Path & "\Graficos\Principal\RecuperarSiguienteA.jpg")
End Sub

