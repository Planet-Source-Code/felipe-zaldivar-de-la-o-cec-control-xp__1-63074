VERSION 5.00
Begin VB.Form FrmCUsr 
   BackColor       =   &H00FFCEBB&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "::: Cambio de usuario :::"
   ClientHeight    =   1815
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5535
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "FrmCUsr.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1815
   ScaleWidth      =   5535
   Begin VB.PictureBox PicDes 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1335
      Left            =   4440
      Picture         =   "FrmCUsr.frx":08CA
      ScaleHeight     =   85
      ScaleMode       =   0  'User
      ScaleWidth      =   85
      TabIndex        =   11
      Top             =   1920
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.PictureBox PicBloq 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1335
      Left            =   3000
      Picture         =   "FrmCUsr.frx":8074
      ScaleHeight     =   85
      ScaleMode       =   0  'User
      ScaleWidth      =   85
      TabIndex        =   10
      Top             =   1920
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00E89C78&
      BorderStyle     =   0  'None
      Height          =   1575
      Left            =   120
      ScaleHeight     =   1575
      ScaleWidth      =   5295
      TabIndex        =   6
      Top             =   120
      Width           =   5295
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
      Begin VB.PictureBox PicMain 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   1335
         Left            =   3840
         Picture         =   "FrmCUsr.frx":10210
         ScaleHeight     =   85
         ScaleMode       =   0  'User
         ScaleWidth      =   85
         TabIndex        =   7
         Top             =   120
         Width           =   1335
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
      Begin VB.CommandButton CmdCancelar 
         BackColor       =   &H00E89C78&
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   2520
         TabIndex        =   3
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Password:"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   650
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cuenta:"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   170
         Width           =   555
      End
   End
   Begin VB.PictureBox PicCU 
      BorderStyle     =   0  'None
      Height          =   1335
      Left            =   1560
      Picture         =   "FrmCUsr.frx":17643
      ScaleHeight     =   85
      ScaleMode       =   0  'User
      ScaleWidth      =   85
      TabIndex        =   5
      Top             =   1920
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.PictureBox PicApagar 
      BorderStyle     =   0  'None
      Height          =   1335
      Left            =   120
      Picture         =   "FrmCUsr.frx":1E838
      ScaleHeight     =   85
      ScaleMode       =   0  'User
      ScaleWidth      =   85
      TabIndex        =   4
      Top             =   1920
      Visible         =   0   'False
      Width           =   1335
   End
End
Attribute VB_Name = "FrmCUsr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public OpcionAcceso As Integer
Public ServidorB As Boolean

Private Sub CmdAceptar_Click()
    ConexionLogin
    SqlCLogin = "Select * from Tbl_Acceso where C_Acceso='" _
    & TxtUsuario.Text & "' and Password='" & TxtPassword.Text & "' and " _
    & " Nivel=3 and Usr_Bloqueado=0"
    RsCLogin.Open SqlCLogin, Conecta, adOpenDynamic, adLockBatchOptimistic
    If Not RsCLogin.EOF Then
        If OpcionAcceso = 3 Then
                'Bloquear terminal
                Dim FrmB As Form
                If ServidorB = False Then
                    MDIPrincipal.ImgRoll.Enabled = False
                    MDIPrincipal.PicContenedor.Enabled = False
                    MDIPrincipal.PicVentanas.Enabled = False
                    
                    For Each FrmB In Forms
                    FrmB.Enabled = False
                    Next
                    
                    MDIPrincipal.Enabled = True
                    Me.Enabled = True
                    ServidorB = True
                    
                    TxtUsuario.Text = ""
                    TxtPassword = ""
                    TxtUsuario.SetFocus
                    Me.Caption = "::: Desloquear servidor :::"
                    PicMain.Picture = PicDes.Picture
                Else
                    MDIPrincipal.ImgRoll.Enabled = True
                    MDIPrincipal.PicVentanas.Enabled = True
                    MDIPrincipal.PicContenedor.Enabled = True
                    For Each FrmB In Forms
                        FrmB.Enabled = True
                    Next
                    ServidorB = False
                    Unload Me
                End If
                Set FrmB = Nothing
            Exit Sub
        End If
        RsCLogin.Close
        SqlCLogin = "Select * From Tbl_Maquina Where C_Acceso='" & TxtUsuario.Text & "'"
        RsCLogin.Open SqlCLogin, Conecta, adOpenDynamic, adLockBatchOptimistic
        If RsCLogin.EOF Then
            Dim frm As Form
            If OpcionAcceso = 1 Then
                'Cambiar de usuario
                Call RegCambiarUsuario
                FrmPrincipal.TxtUsuario.Text = TxtUsuario.Text
                Unload Me
            ElseIf OpcionAcceso = 2 Then
                'Terminar el Servidor
                
                For Each frm In Forms
                    Unload frm
                    Set frm = Nothing
                    DoEvents
                Next
                End
            End If
        Else
            If OpcionAcceso = 2 Then
                'Terminar el Servidor
                For Each frm In Forms
                    Unload frm
                    Set frm = Nothing
                    DoEvents
                Next
                End
            End If
            MsgBox "Esta cuenta ya esta siendo utilizada... (" & RsCLogin!C_Maq & ")", , "Atención!!!"
            TxtUsuario.SetFocus
        End If
    Else
        MsgBox "Administrador no registrado...", , "Atención!!!"
        TxtUsuario.SetFocus
    End If
End Sub

Private Sub RegCambiarUsuario()
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
    SqlCLogin = "Select * From Tbl_Historial Where C_Acceso='" & FrmPrincipal.TxtUsuario.Text & "' " _
                & "and C_Maq='SERVIDOR' and ((Hora_Entrada)Is Not Null) and ((Hora_Salida)Is  Null)"
    RsCLogin.Open SqlCLogin, Conecta, adOpenStatic, adLockPessimistic
    If Not RsCLogin.EOF Then
        RsCLogin!Hora_Salida = Time
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
    If ServidorB = True Then Exit Sub
    Unload Me
End Sub

Private Sub Form_Load()
    MDIPrincipal.AgregarVentana Me, "Opciones"
    Redondear Me
    PosicionInicial Me
    DoEvents
End Sub

Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then
        Visible = False
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MDIPrincipal.RemoverVentana Me, "Opciones"
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


