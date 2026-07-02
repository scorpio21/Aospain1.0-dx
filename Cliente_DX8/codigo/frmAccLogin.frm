VERSION 5.00
Begin VB.Form frmAccLogin 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "AoSpain - Ingreso a la Cuenta"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtEmail 
      Height          =   285
      Left            =   1560
      TabIndex        =   5
      Top             =   1320
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.CommandButton cmdCrear 
      Caption         =   "CREAR CUENTA"
      Height          =   495
      Left            =   480
      TabIndex        =   4
      Top             =   2400
      Width           =   1695
   End
   Begin VB.CommandButton cmdLogin 
      Caption         =   "INGRESAR"
      Height          =   495
      Left            =   2400
      TabIndex        =   3
      Top             =   2400
      Width           =   1695
   End
   Begin VB.TextBox txtPass 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1560
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   840
      Width           =   2535
   End
   Begin VB.TextBox txtAccName 
      Height          =   285
      Left            =   1560
      TabIndex        =   0
      Top             =   360
      Width           =   2535
   End
   Begin VB.Label lblEmail 
      BackStyle       =   0  'Transparent
      Caption         =   "Email:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   480
      TabIndex        =   6
      Top             =   1320
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   480
      TabIndex        =   2
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Cuenta:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   480
      TabIndex        =   7
      Top             =   360
      Width           =   975
   End
End
Attribute VB_Name = "frmAccLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub cmdLogin_Click()
    If txtAccName.Text <> "" And txtPass.Text <> "" Then
        ' Llamamos a la nueva función de envío
        Call EnviarLoginCuenta(txtAccName.Text, txtPass.Text)
    Else
        MsgBox "Completa todos los campos."
    End If
End Sub

Private Sub cmdCrear_Click()
    ' Lógica simple para mostrar el campo email si quieren crear cuenta
    If Not txtEmail.Visible Then
        txtEmail.Visible = True
        lblEmail.Visible = True
        cmdCrear.Caption = "CONFIRMAR REGISTRO"
    Else
        If txtAccName.Text <> "" And txtPass.Text <> "" And txtEmail.Text <> "" Then
             Call EnviarCrearCuenta(txtAccName.Text, txtPass.Text, txtEmail.Text)
        End If
    End If
End Sub
