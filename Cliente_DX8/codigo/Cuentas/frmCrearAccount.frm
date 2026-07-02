VERSION 5.00
Begin VB.Form frmCrearAccount 
   BorderStyle     =   0  'None
   ClientHeight    =   9390
   ClientLeft      =   0
   ClientTop       =   60
   ClientWidth     =   3945
   ControlBox      =   0   'False
   Icon            =   "frmCrearAccount.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9390
   ScaleWidth      =   3945
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      Text            =   "frmCrearAccount.frx":000C
      Top             =   6240
      Width           =   3615
   End
   Begin VB.TextBox respuesta 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000080&
      BeginProperty Font 
         Name            =   "Candara"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   340
      Left            =   480
      TabIndex        =   6
      Top             =   4440
      Width           =   3100
   End
   Begin VB.TextBox pregunta 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000080&
      BeginProperty Font 
         Name            =   "Candara"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   340
      Left            =   480
      TabIndex        =   5
      Top             =   3786
      Width           =   3100
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   195
      Left            =   240
      MaskColor       =   &H00004080&
      TabIndex        =   7
      Top             =   7560
      UseMaskColor    =   -1  'True
      Width           =   170
   End
   Begin VB.TextBox Mail2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000080&
      BeginProperty Font 
         Name            =   "Candara"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   340
      Left            =   480
      TabIndex        =   4
      Top             =   3110
      Width           =   3100
   End
   Begin VB.TextBox Nombre 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000080&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Candara"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   340
      Left            =   480
      MaxLength       =   25
      TabIndex        =   0
      Top             =   426
      Width           =   3100
   End
   Begin VB.TextBox Pass 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000080&
      BeginProperty Font 
         Name            =   "Candara"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   340
      IMEMode         =   3  'DISABLE
      Left            =   480
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1070
      Width           =   3100
   End
   Begin VB.TextBox RePass 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000080&
      BeginProperty Font 
         Name            =   "Candara"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   340
      IMEMode         =   3  'DISABLE
      Left            =   480
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1720
      Width           =   3100
   End
   Begin VB.TextBox Mail 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000080&
      BeginProperty Font 
         Name            =   "Candara"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   340
      Left            =   480
      TabIndex        =   3
      Top             =   2370
      Width           =   3100
   End
   Begin VB.Image Image1 
      Height          =   990
      Left            =   100
      Top             =   8400
      Width           =   1785
   End
   Begin VB.Image Image2 
      Height          =   960
      Left            =   2000
      Top             =   8400
      Width           =   1770
   End
End
Attribute VB_Name = "frmCrearAccount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_Load()
Me.Picture = LoadPicture(App.Path & "\Graficos\Principal\CrearCuenta_Main.jpg")
Image1.Picture = LoadPicture(App.Path & "\Graficos\Principal\CrearCuenta_BAtrasN.jpg")
Image2.Picture = LoadPicture(App.Path & "\Graficos\Principal\CrearCuenta_BCrearN.jpg")
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Image1.Picture = LoadPicture(App.Path & "\Graficos\Principal\CrearCuenta_BAtrasN.jpg")
Image2.Picture = LoadPicture(App.Path & "\Graficos\Principal\CrearCuenta_BCrearN.jpg")
End Sub

Private Sub Image1_Click()
Unload Me

End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
Image1.Picture = LoadPicture(App.Path & "\Graficos\Principal\CrearCuenta_BAtrasA.jpg")
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Image1.Picture = LoadPicture(App.Path & "\Graficos\Principal\CrearCuenta_BAtrasI.jpg")
End Sub

Private Sub Image2_Click()
If Len(nombre.Text) < 5 Then
    MsgBox "El nombre de la cuenta debe de tener mas de 4 caracteres.", vbCritical
    Exit Sub
End If
        
If Len(Pass.Text) < 6 Then
    MsgBox "El password de la cuenta debe de tener mas de 6 caracteres.", vbCritical
    Exit Sub
End If

If Pass <> RePass Then
    MsgBox "Las passwords que tipeo no coinciden", , "Coco rules"
    Exit Sub
End If

If Not CheckMailString(Mail) Then
    MsgBox "Direccion de mail invalida."
    Exit Sub
End If
If Mail.Text <> Mail2.Text Then
MsgBox "Los emails no coinciden"
Exit Sub
End If
If nombre = "" Or Pass = "" Or RePass = "" Or Mail = "" Or pregunta = "" Or respuesta = "" Then
    MsgBox "Completa todo!"
    Exit Sub
End If
If Not Check1.value = vbChecked Then
    MsgBox "Debe Aceptar los términos y Reglamento para poder crear la cuenta.", vbCritical
    Exit Sub
End If
Call SendData("NACCNT" & nombre & "," & Pass & "," & Mail & "," & pregunta & "," & respuesta)

Unload Me
MsgBox "La cuenta fue creada con éxito."
End Sub

Private Sub Image2_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
Image2.Picture = LoadPicture(App.Path & "\Graficos\Principal\CrearCuenta_BCrearA.jpg")
End Sub

Private Sub Image2_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Image2.Picture = LoadPicture(App.Path & "\Graficos\Principal\CrearCuenta_BCrearI.jpg")
End Sub
