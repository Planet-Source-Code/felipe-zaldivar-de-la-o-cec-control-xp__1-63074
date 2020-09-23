VERSION 5.00
Begin VB.Form FrmLogin 
   BackColor       =   &H00FFCEBB&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "::: Acceso Restringido :::"
   ClientHeight    =   1815
   ClientLeft      =   6000
   ClientTop       =   5730
   ClientWidth     =   5535
   ControlBox      =   0   'False
   Icon            =   "FrmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1815
   ScaleWidth      =   5535
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00E89C78&
      BorderStyle     =   0  'None
      Height          =   1575
      Left            =   120
      ScaleHeight     =   1575
      ScaleWidth      =   5295
      TabIndex        =   4
      Top             =   120
      Width           =   5295
      Begin VB.CommandButton CmdCancelar 
         BackColor       =   &H00E89C78&
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   2520
         TabIndex        =   3
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox TxtUsuario 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   960
         MaxLength       =   9
         TabIndex        =   0
         Top             =   120
         Width           =   2775
      End
      Begin VB.PictureBox PictStop 
         BorderStyle     =   0  'None
         Height          =   1335
         Left            =   3840
         Picture         =   "FrmLogin.frx":08CA
         ScaleHeight     =   85
         ScaleMode       =   0  'User
         ScaleWidth      =   85
         TabIndex        =   6
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton CmdAceptar 
         BackColor       =   &H00E89C78&
         Caption         =   "&Aceptar"
         Enabled         =   0   'False
         Height          =   375
         Left            =   960
         TabIndex        =   2
         Top             =   1080
         Width           =   1335
      End
      Begin VB.TextBox TxtPassword 
         BorderStyle     =   0  'None
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   960
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   600
         Width           =   2775
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cuenta:"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   170
         Width           =   555
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Password:"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   650
         Width           =   735
      End
   End
End
Attribute VB_Name = "FrmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdAceptar_Click()
    If VerificarAdmon = False Then Exit Sub
    ConexionLogin
    SqlCLogin = "Select * from Tbl_Acceso where C_Acceso='" _
    & TxtUsuario.Text & "' and Password='" & TxtPassword.Text & "' and " _
    & " Nivel=3 and Usr_Bloqueado=0"
    RsCLogin.Open SqlCLogin, Conecta, adOpenDynamic, adLockBatchOptimistic
    If Not RsCLogin.EOF Then
        Call RegistrarUsuario
        MDIPrincipal.MenuDespleglable 0
        MDIPrincipal.ImgRoll.Enabled = True
        ActivarForm FrmPrincipal
        PosicionInicial FrmPrincipal
        FrmPrincipal.TxtUsuario.Text = TxtUsuario.Text
        Unload Me
    Else
        MsgBox "Administrador no registrado!!!", , "Atención!!!"
        TxtUsuario.SetFocus
    End If
    RsCLogin.Close
End Sub

Private Sub RegistrarUsuario()
    ConexionLogin
    SqlCLogin = "Select * From Tbl_Maquina Where C_Maq='SERVIDOR'"
    RsCLogin.Open SqlCLogin, Conecta, adOpenStatic, adLockPessimistic
    If Not RsCLogin.EOF Then
        RsCLogin!Maq_Ocupada = True
        RsCLogin!Maq_Inicio = True
        RsCLogin!C_Acceso = TxtUsuario.Text
        RsCLogin.Update
    End If
    RsCLogin.Close
    SqlCLogin = "Select * From Tbl_Historial"
    RsCLogin.Open SqlCLogin, Conecta, adOpenStatic, adLockPessimistic
    RsCLogin.AddNew
    RsCLogin!C_Acceso = TxtUsuario.Text
    RsCLogin!C_Maq = "SERVIDOR"
    RsCLogin!Fecha_Entrada = Date
    RsCLogin!Hora_Entrada = Time
    RsCLogin.Update
End Sub

Private Sub CmdCancelar_Click()
    Set FrmLogin = Nothing
    Set MDIPrincipal = Nothing
    Set Conecta = Nothing
    End
End Sub

Private Sub Form_Load()
    MDIPrincipal.AgregarVentana Me, "Login"
    Redondear Me
    PosicionInicial Me
End Sub

Private Function VerificarAdmon() As Boolean
    VerificarAdmon = True
    ConexionLogin
    SqlCLogin = "Select * from Tbl_Acceso where Nivel=3"
    RsCLogin.Open SqlCLogin, Conecta, adOpenDynamic, adLockBatchOptimistic
    If RsCLogin.EOF Then
        VerificarAdmon = False
        ActivarForm FrmUsrAlta
    End If
End Function

Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then
        Visible = False
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MDIPrincipal.RemoverVentana Me, "Login"
End Sub

Private Sub TxtPassword_Change()
    If TxtPassword.Text = "" Or TxtUsuario.Text = "" Then
        CmdAceptar.Enabled = False
    Else
        CmdAceptar.Enabled = True
    End If
End Sub

Private Sub TxtUsuario_Change()
    If TxtUsuario.Text = "" Or TxtPassword.Text = "" Then
        CmdAceptar.Enabled = False
    Else
        CmdAceptar.Enabled = True
    End If
End Sub

Private Sub TxtUsuario_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And TxtUsuario.Text <> "" Then
        TxtPassword.SetFocus
    End If
End Sub

Private Sub TxtPassword_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And TxtPassword.Text <> "" Then
        CmdAceptar.SetFocus
    End If
End Sub
