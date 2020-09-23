VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FrmUsrAlta 
   BackColor       =   &H00FFCEBB&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "::: Registro de usuarios :::"
   ClientHeight    =   4575
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4695
   Icon            =   "FrmUsrAlta.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4575
   ScaleWidth      =   4695
   Begin VB.PictureBox P1 
      BackColor       =   &H00800000&
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   0
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   26
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Data DataGrupo 
      Caption         =   "DataGrupo"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   4680
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   120
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Frame FrmOpt 
      BackColor       =   &H00E89C78&
      BorderStyle     =   0  'None
      Height          =   4335
      Left            =   120
      TabIndex        =   15
      Top             =   120
      Width           =   4455
      Begin VB.CommandButton CmdNuevo 
         Appearance      =   0  'Flat
         BackColor       =   &H00E89C78&
         Caption         =   "&Nuevo"
         Height          =   315
         Left            =   2520
         TabIndex        =   13
         Top             =   3840
         Width           =   855
      End
      Begin VB.CommandButton CmdCT 
         BackColor       =   &H00FFCEBB&
         Caption         =   "&C. Temporal"
         Height          =   255
         Left            =   3240
         TabIndex        =   4
         ToolTipText     =   "Crear cuenta temporal..."
         Top             =   930
         Width           =   1095
      End
      Begin VB.TextBox TxtCuenta 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1560
         MaxLength       =   9
         TabIndex        =   3
         ToolTipText     =   "Cuenta de acceso al sistema"
         Top             =   915
         Width           =   1575
      End
      Begin VB.TextBox TxtPassword 
         BorderStyle     =   0  'None
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1560
         MaxLength       =   10
         TabIndex        =   5
         ToolTipText     =   "Clave o password..."
         Top             =   1254
         Width           =   2775
      End
      Begin VB.CheckBox ChkBloqueo 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&No"
         Height          =   255
         Left            =   1560
         TabIndex        =   7
         ToolTipText     =   "Bloquear cuenta de usuario..."
         Top             =   1980
         Width           =   2775
      End
      Begin VB.ComboBox CmbGrupo 
         Height          =   315
         ItemData        =   "FrmUsrAlta.frx":0CCA
         Left            =   1560
         List            =   "FrmUsrAlta.frx":0CCC
         Style           =   2  'Dropdown List
         TabIndex        =   0
         ToolTipText     =   "Clave de grupo del usuario"
         Top             =   180
         Width           =   1575
      End
      Begin VB.ComboBox CmbNivel 
         Height          =   315
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   6
         ToolTipText     =   "Nivel de usuario (1: Alumno, 2: Profesor/Otro, 3: Administrador)"
         Top             =   1602
         Width           =   2775
      End
      Begin VB.TextBox TxtGrupo 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   3084
         Width           =   2775
      End
      Begin VB.TextBox TxtCarrera 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   3435
         Width           =   2775
      End
      Begin VB.TextBox TxtGrado 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   2736
         Width           =   2775
      End
      Begin VB.CommandButton CmdAgregar 
         Appearance      =   0  'Flat
         BackColor       =   &H00E89C78&
         Caption         =   "&Registrar"
         Height          =   315
         Left            =   1560
         TabIndex        =   12
         Top             =   3840
         Width           =   855
      End
      Begin VB.CommandButton CmdGrupo 
         BackColor       =   &H00FFCEBB&
         Caption         =   "&Grupos"
         Height          =   255
         Left            =   3240
         TabIndex        =   1
         ToolTipText     =   "Grupos..."
         Top             =   210
         Width           =   1095
      End
      Begin VB.CommandButton CmdSalir 
         Caption         =   "&Cancelar"
         Height          =   315
         Left            =   3480
         TabIndex        =   14
         Top             =   3840
         Width           =   855
      End
      Begin VB.TextBox TxtNombre 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1560
         MaxLength       =   99
         TabIndex        =   2
         ToolTipText     =   "Nombre del usuario (Nombre's'_Apellido Paterno_Apellido Materno)"
         Top             =   558
         Width           =   2775
      End
      Begin MSComCtl2.DTPicker DPRegistro 
         Height          =   375
         Left            =   1560
         TabIndex        =   8
         ToolTipText     =   "Fecha de registro del usuario"
         Top             =   2295
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   661
         _Version        =   393216
         Format          =   49872897
         CurrentDate     =   38586
         MinDate         =   36526
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FF0000&
         BackStyle       =   0  'Transparent
         Caption         =   "Clave de grupo:"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   120
         TabIndex        =   25
         Top             =   240
         Width           =   1305
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FF0000&
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre:"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   120
         TabIndex        =   24
         Top             =   600
         Width           =   1305
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FF0000&
         BackStyle       =   0  'Transparent
         Caption         =   "Cuenta:"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   120
         TabIndex        =   23
         Top             =   960
         Width           =   1305
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FF0000&
         BackStyle       =   0  'Transparent
         Caption         =   "Password:"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   120
         TabIndex        =   22
         Top             =   1320
         Width           =   1305
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FF0000&
         BackStyle       =   0  'Transparent
         Caption         =   "Nivel:"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   120
         TabIndex        =   21
         Top             =   1680
         Width           =   1305
      End
      Begin VB.Label Label9 
         BackColor       =   &H00FF0000&
         BackStyle       =   0  'Transparent
         Caption         =   "Bloqueo:"
         Height          =   195
         Left            =   120
         TabIndex        =   20
         Top             =   2040
         Width           =   1305
      End
      Begin VB.Label Label12 
         BackColor       =   &H00FF0000&
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha de registro:"
         Height          =   195
         Left            =   120
         TabIndex        =   19
         Top             =   2400
         Width           =   1305
      End
      Begin VB.Label Label14 
         BackColor       =   &H00FF0000&
         BackStyle       =   0  'Transparent
         Caption         =   "Carrera.:"
         Height          =   195
         Left            =   120
         TabIndex        =   18
         Top             =   3480
         Width           =   1305
      End
      Begin VB.Label Label15 
         BackColor       =   &H00FF0000&
         BackStyle       =   0  'Transparent
         Caption         =   "Grupo.:"
         Height          =   195
         Left            =   120
         TabIndex        =   17
         Top             =   3120
         Width           =   1305
      End
      Begin VB.Label Label16 
         BackColor       =   &H00FF0000&
         BackStyle       =   0  'Transparent
         Caption         =   "Grado:"
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   2760
         Width           =   1305
      End
   End
End
Attribute VB_Name = "FrmUsrAlta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private CGrupo As String
Private CFecha As Date

Private Sub ChkBloqueo_Click()
    If ChkBloqueo.Value = 0 Then
        ChkBloqueo.Caption = "No"
    Else
        ChkBloqueo.Caption = "Sí"
    End If
End Sub

Private Sub ChkBloqueo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        DPRegistro.SetFocus
    End If
End Sub

Private Sub CmbGrupo_Click()
    Call BuscarDG 'buscamos los campos del grupo
End Sub

Private Sub BuscarDG()
    DataGrupo.DatabaseName = Direccion
    DataGrupo.RecordSource = "select * from Tbl_Grupo where C_Grupo='" & CmbGrupo.Text & "'"
    DataGrupo.Connect = ";Pwd=" & C_BD
    Call DataGrupo.Refresh
    If Not DataGrupo.Recordset.EOF Then
        TxtGrado.Text = DataGrupo.Recordset("Grado")
        TxtGrupo.Text = DataGrupo.Recordset("Grupo")
        TxtCarrera.Text = DataGrupo.Recordset("Carrera")
    Else
        TxtGrado.Text = ""
        TxtGrupo.Text = ""
        TxtCarrera.Text = ""
    End If
End Sub

Private Sub CmbGrupo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And CmbGrupo.Text <> "" Then
        TxtNombre.SetFocus
    End If
End Sub

Private Sub CmbNivel_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And CmbNivel.Text <> "" Then
        ChkBloqueo.SetFocus
    End If
End Sub

Private Sub CmdAgregar_Click()
    If VerificarCampos = False Then Exit Sub
    Dim MensajeUsr2 As String
    MensajeUsr2 = "Los datos son correctos?"
    If Preguntar(MensajeUsr2) = False Then Exit Sub
    
    ConexionPrincipal
    Sql = "select * from Tbl_Acceso where C_Acceso= '" & TxtCuenta.Text & "'"
    Rs.Open Sql, Conecta, adOpenDynamic, adLockBatchOptimistic
    
    If Not Rs.EOF Then
        MsgBox "Usuario ya registrado!!!", , "Atención!!!"
    Else
        Rs.Close
        Sql = "insert into Tbl_Acceso ([C_Grupo],[Nombre]" & _
        ",[C_Acceso],[Password],[Nivel],[Usr_Bloqueado],[Fecha_Reg],[C_U_Registro])" & _
        " VALUES ('" & CGrupo & "','" & TxtNombre.Text & "','" & _
        TxtCuenta.Text & "','" & TxtPassword.Text & "','" _
        & Val(CmbNivel.Text) & "','" & ChkBloqueo.Value & "','" & CFecha & "','" & _
        FrmPrincipal.TxtUsuario.Text & "')"
        Rs.Open Sql, Conecta, adOpenDynamic, adLockBatchOptimistic
        MsgBox "Usuario registrado satisfactoriamente!!!", , "Atención!!!": DoEvents
    End If
End Sub

Private Function VerificarCampos() As Boolean
    VerificarCampos = False
    CGrupo = ""
    If Val(CmbNivel.Text) = 1 Then
        If CmbGrupo.ListCount > 0 Then
            CGrupo = CmbGrupo.Text
        Else
            MsgBox "Seleciona una clave de grupo...", , "Atención!!!": DoEvents
            VerificarCampos = False
            Exit Function
        End If
    ElseIf Val(CmbNivel.Text) = 2 Then
        CGrupo = "Profesor"
    ElseIf Val(CmbNivel.Text) = 3 Then
        CGrupo = "Administrador"
    End If
    If Left(TxtCuenta.Text, 2) = "T-" Then
        CFecha = CDate("01/01/1900")
    Else
        CFecha = Date
    End If
    
    If TxtNombre.Text = "" Or TxtCuenta.Text = "" Or TxtPassword.Text = "" Then
        MsgBox "No se pueden dejar campos en blanco...", , "Atención!!!": DoEvents
        VerificarCampos = False
        Exit Function
    End If
    If CmbNivel.Text = "1" Then
        If CmbGrupo.ListCount = 0 Or TxtGrado.Text = "" Or TxtGrupo.Text = "" Or TxtCarrera.Text = "" Then
            MsgBox "No se pueden dejar campos en blanco...", , "Atención!!!": DoEvents
            VerificarCampos = False
            Exit Function
        End If
    End If
    VerificarCampos = True
End Function

Private Sub CmdCT_Click()
    TxtCuenta.Text = CrearCT
    TxtPassword.Text = TxtCuenta.Text
    CmbNivel.SetFocus
End Sub

Private Sub CmdGrupo_Click()
    ActivarForm FrmGrupos
End Sub

Private Sub CmdNuevo_Click()
    Call VaciarCampos
End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Sub DPRegistro_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        TxtGrado.SetFocus
    End If
End Sub

Private Sub Form_Load()
    MDIPrincipal.AgregarVentana Me, "N. Usuario", "Nuevo usuario..."
    Call CargarGrupos
    Call CargarNiveles
    Call VaciarCampos
    InsertarPicture Me
    PosicionInicial Me
End Sub

Private Sub CargarNiveles()
    CmbNivel.Clear
    CmbNivel.AddItem "1"
    CmbNivel.AddItem "2"
    CmbNivel.AddItem "3"
    CmbNivel.Text = "1"
End Sub

Private Sub CargarGrupos()
    DataGrupo.DatabaseName = Direccion
    DataGrupo.RecordSource = "SELECT * FROM [Tbl_Grupo]"
    DataGrupo.Connect = ";Pwd=" & C_BD
    Call DataGrupo.Refresh
    CmbGrupo.Clear
    Do While Not DataGrupo.Recordset.EOF
        CmbGrupo.AddItem DataGrupo.Recordset("C_Grupo")
        DataGrupo.Recordset.MoveNext
    Loop
    If CmbGrupo.ListCount > 0 Then CmbGrupo.Text = CmbGrupo.List(0)
End Sub

Private Sub VaciarCampos()
    Call CargarGrupos
    CmbNivel.Text = "1"
    DPRegistro.Value = Date
    TxtNombre.Text = ""
    TxtCuenta.Text = ""
    TxtPassword.Text = ""
    TxtGrado.Text = ""
    TxtGrupo.Text = ""
    TxtCarrera.Text = ""
    Call CmbGrupo_Click
End Sub

Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then
        Visible = False
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MDIPrincipal.RemoverVentana Me, "N. Usuario"
End Sub

Private Sub TxtCarrera_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And TxtCarrera.Text <> "" Then
        CmdAgregar.SetFocus
    End If
End Sub

Private Sub TxtCuenta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And TxtCuenta.Text <> "" Then
        TxtPassword.SetFocus
    End If
End Sub

Private Sub TxtGrado_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And TxtGrado.Text <> "" Then
        TxtGrupo.SetFocus
    End If
End Sub

Private Sub TxtGrupo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And TxtGrupo.Text <> "" Then
        TxtCarrera.SetFocus
    End If
End Sub

Private Sub TxtNombre_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And TxtNombre.Text <> "" Then
        TxtCuenta.SetFocus
    End If
End Sub

Private Sub TxtPassword_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And TxtPassword.Text <> "" Then
        CmbNivel.SetFocus
    End If
End Sub
