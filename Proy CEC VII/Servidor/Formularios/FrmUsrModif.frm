VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FrmUsrModif 
   BackColor       =   &H00FFCEBB&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "::: Modificaciones a usuario :::"
   ClientHeight    =   5310
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4695
   Icon            =   "FrmUsrModif.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5310
   ScaleWidth      =   4695
   Begin VB.PictureBox P1 
      BackColor       =   &H00800000&
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   0
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   30
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Data DataMU 
      Caption         =   "DataMU"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   420
      Left            =   4680
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   600
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.TextBox TxtCT 
      Height          =   375
      Left            =   4680
      TabIndex        =   16
      Text            =   "Text1"
      Top             =   120
      Width           =   2175
   End
   Begin VB.Frame FrmOpt 
      BackColor       =   &H00E89C78&
      BorderStyle     =   0  'None
      Height          =   5055
      Left            =   120
      TabIndex        =   15
      Top             =   120
      Width           =   4455
      Begin VB.TextBox TxtBloqueos 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1560
         MaxLength       =   3
         TabIndex        =   8
         Top             =   2660
         Width           =   2775
      End
      Begin VB.TextBox TxtAmonestaciones 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1560
         MaxLength       =   3
         TabIndex        =   6
         Top             =   1990
         Width           =   2775
      End
      Begin VB.TextBox TxtNombre 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1560
         MaxLength       =   99
         TabIndex        =   2
         ToolTipText     =   "Nombre del usuario (Nombre's'_Apellido Paterno_Apellido Materno)"
         Top             =   560
         Width           =   2775
      End
      Begin VB.CommandButton CmdSalir 
         Caption         =   "&Cancelar"
         Height          =   315
         Left            =   3000
         TabIndex        =   14
         Top             =   4560
         Width           =   1335
      End
      Begin VB.CommandButton CmdGrupo 
         BackColor       =   &H00FFCEBB&
         Caption         =   "&Grupos"
         Height          =   255
         Left            =   3360
         TabIndex        =   1
         ToolTipText     =   "Grupos..."
         Top             =   210
         Width           =   975
      End
      Begin VB.CommandButton CmdModificar 
         Appearance      =   0  'Flat
         BackColor       =   &H00E89C78&
         Caption         =   "&Modificar"
         Height          =   315
         Left            =   1560
         TabIndex        =   13
         Top             =   4560
         Width           =   1335
      End
      Begin VB.TextBox TxtGrado 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   10
         Top             =   3450
         Width           =   2775
      End
      Begin VB.TextBox TxtCarrera 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   12
         Top             =   4155
         Width           =   2775
      End
      Begin VB.TextBox TxtGrupo 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   11
         Top             =   3800
         Width           =   2775
      End
      Begin VB.ComboBox CmbNivel 
         Height          =   315
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   5
         ToolTipText     =   "Nivel de usuario (1: Alumno, 2: Profesor/Otro, 3: Administrador)"
         Top             =   1610
         Width           =   2775
      End
      Begin VB.ComboBox CmbGrupo 
         Height          =   315
         ItemData        =   "FrmUsrModif.frx":0CCA
         Left            =   1560
         List            =   "FrmUsrModif.frx":0CCC
         Style           =   2  'Dropdown List
         TabIndex        =   0
         ToolTipText     =   "Clave de grupo del usuario"
         Top             =   180
         Width           =   1695
      End
      Begin VB.CheckBox ChkBloqueo 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&No"
         Height          =   255
         Left            =   1560
         TabIndex        =   7
         ToolTipText     =   "Bloquear cuenta de usuario..."
         Top             =   2340
         Width           =   2775
      End
      Begin VB.TextBox TxtPassword 
         BorderStyle     =   0  'None
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1560
         MaxLength       =   10
         TabIndex        =   4
         ToolTipText     =   "Clave o password..."
         Top             =   1260
         Width           =   2775
      End
      Begin VB.TextBox TxtCuenta 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1560
         MaxLength       =   9
         TabIndex        =   3
         ToolTipText     =   "Cuenta de acceso al sistema"
         Top             =   910
         Width           =   2775
      End
      Begin MSComCtl2.DTPicker DPRegistro 
         Height          =   375
         Left            =   1560
         TabIndex        =   9
         ToolTipText     =   "Fecha de registro del usuario"
         Top             =   3010
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   661
         _Version        =   393216
         Format          =   49872897
         CurrentDate     =   38586
         MinDate         =   36526
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FF0000&
         BackStyle       =   0  'Transparent
         Caption         =   "Amonestaciones:"
         Height          =   195
         Left            =   120
         TabIndex        =   29
         Top             =   2040
         Width           =   1305
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FF0000&
         BackStyle       =   0  'Transparent
         Caption         =   "Bloqueo:"
         Height          =   195
         Left            =   120
         TabIndex        =   28
         Top             =   2400
         Width           =   1305
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FF0000&
         BackStyle       =   0  'Transparent
         Caption         =   "Clave de grupo:"
         Height          =   195
         Left            =   120
         TabIndex        =   27
         Top             =   240
         Width           =   1305
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FF0000&
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre:"
         Height          =   195
         Left            =   120
         TabIndex        =   26
         Top             =   600
         Width           =   1305
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FF0000&
         BackStyle       =   0  'Transparent
         Caption         =   "Cuenta:"
         Height          =   195
         Left            =   120
         TabIndex        =   25
         Top             =   960
         Width           =   1305
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FF0000&
         BackStyle       =   0  'Transparent
         Caption         =   "Password:"
         Height          =   195
         Left            =   120
         TabIndex        =   24
         Top             =   1320
         Width           =   1305
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FF0000&
         BackStyle       =   0  'Transparent
         Caption         =   "Nivel:"
         Height          =   195
         Left            =   120
         TabIndex        =   23
         Top             =   1680
         Width           =   1305
      End
      Begin VB.Label Label9 
         BackColor       =   &H00FF0000&
         BackStyle       =   0  'Transparent
         Caption         =   "Bloqueos:"
         Height          =   195
         Left            =   120
         TabIndex        =   22
         Top             =   2760
         Width           =   1305
      End
      Begin VB.Label Label12 
         BackColor       =   &H00FF0000&
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha de registro:"
         Height          =   195
         Left            =   120
         TabIndex        =   21
         Top             =   3120
         Width           =   1305
      End
      Begin VB.Label Label14 
         BackColor       =   &H00FF0000&
         BackStyle       =   0  'Transparent
         Caption         =   "Carrera.:"
         Height          =   195
         Left            =   120
         TabIndex        =   20
         Top             =   4200
         Width           =   1305
      End
      Begin VB.Label Label15 
         BackColor       =   &H00FF0000&
         BackStyle       =   0  'Transparent
         Caption         =   "Grupo.:"
         Height          =   195
         Left            =   120
         TabIndex        =   19
         Top             =   3840
         Width           =   1305
      End
      Begin VB.Label Label16 
         BackColor       =   &H00FF0000&
         BackStyle       =   0  'Transparent
         Caption         =   "Grado:"
         Height          =   195
         Left            =   120
         TabIndex        =   18
         Top             =   3480
         Width           =   1305
      End
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FF0000&
      BackStyle       =   0  'Transparent
      Caption         =   "Nivel:"
      Height          =   195
      Left            =   240
      TabIndex        =   17
      Top             =   2595
      Width           =   1305
   End
End
Attribute VB_Name = "FrmUsrModif"
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

Private Sub CmbGrupo_Click()
    Call BuscarDG 'buscamos los campos del grupo
End Sub

Private Sub BuscarDG()
    DataMU.DatabaseName = Direccion
    DataMU.RecordSource = "select * from Tbl_Grupo where C_Grupo='" & CmbGrupo.Text & "'"
    DataMU.Connect = ";Pwd=" & C_BD
    Call DataMU.Refresh
    If Not DataMU.Recordset.EOF Then
        TxtGrado.Text = DataMU.Recordset("Grado")
        TxtGrupo.Text = DataMU.Recordset("Grupo")
        TxtCarrera.Text = DataMU.Recordset("Carrera")
    Else
        TxtGrado.Text = ""
        TxtGrupo.Text = ""
        TxtCarrera.Text = ""
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
        If DPRegistro.Value > CDate("01/01/1900") Then
            CFecha = DPRegistro.Value
        Else
            DPRegistro.Value = Date
        End If
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

Private Sub CmdGrupo_Click()
    ActivarForm FrmGrupos
End Sub

Private Sub CmdModificar_Click()
    If VerificarCampos = False Then Exit Sub
    Dim MensajeUsr2 As String
    MensajeUsr2 = "Los datos son correctos?"
    If Preguntar(MensajeUsr2) = False Then Exit Sub
    
    ConexionPrincipal
    Sql = "select * from Tbl_Acceso where C_Acceso= '" & TxtCuenta.Text & "'"
    Rs.Open Sql, Conecta, adOpenDynamic, adLockBatchOptimistic
    
    If Rs.EOF Then
        Sql = "Update Tbl_Acceso set Nombre='" & TxtNombre.Text & "'," _
        & " C_Acceso='" & TxtCuenta.Text & "', [Password]='" & TxtPassword.Text & "'," _
        & " C_grupo='" & CmbGrupo.Text & "', Nivel='" & Val(CmbNivel.Text) & "'," _
        & " Amonestaciones='" & Val(TxtAmonestaciones.Text) & "', N_Bloqueos='" & Val(TxtBloqueos.Text) & "'," _
        & " Usr_Bloqueado='" & ChkBloqueo.Value & "' where C_Acceso='" & TxtCT.Text & "'"
            Rs.Close
    Rs.Open Sql, Conecta, adOpenDynamic, adLockBatchOptimistic
    MsgBox "Usuario modificado satisfactoriamente!!!", , "Atención!!!": DoEvents
    Else
        Sql = "Update Tbl_Acceso set Nombre='" & TxtNombre.Text & "'," _
        & "  [Password]='" & TxtPassword.Text & "'," _
        & " C_grupo='" & CmbGrupo.Text & "', Nivel='" & Val(CmbNivel.Text) & "'," _
        & " Amonestaciones='" & Val(TxtAmonestaciones.Text) & "', N_Bloqueos='" & Val(TxtBloqueos.Text) & "'," _
        & " Usr_Bloqueado='" & ChkBloqueo.Value & "' where C_Acceso='" & TxtCT.Text & "'"
            Rs.Close
    Rs.Open Sql, Conecta, adOpenDynamic, adLockBatchOptimistic
    MsgBox "Usuario modificado satisfactoriamente!!!", , "Atención!!!": DoEvents
    End If
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
    MDIPrincipal.AgregarVentana Me, "M. Usuario", "Modificaciones a usuario..."
    PosicionInicial Me
    InsertarPicture Me
    Call CargarNiveles ' carga el combo de niveles
    Call CargarGrupos ' carga los grupos disponibles
End Sub

Private Sub CargarNiveles()
    CmbNivel.Clear
    CmbNivel.AddItem "1"
    CmbNivel.AddItem "2"
    CmbNivel.AddItem "3"
    CmbNivel.Text = "1"
End Sub

Private Sub CargarGrupos()

    DataMU.DatabaseName = Direccion
    DataMU.RecordSource = "SELECT * FROM [Tbl_Grupo]"
    DataMU.Connect = ";Pwd=" & C_BD
    Call DataMU.Refresh
    
    CmbGrupo.Clear
    
    Do While Not DataMU.Recordset.EOF
        CmbGrupo.AddItem DataMU.Recordset("C_Grupo")
        DataMU.Recordset.MoveNext
    Loop
    If CmbGrupo.ListCount > 0 Then CmbGrupo.Text = CmbGrupo.List(0)
End Sub

Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then
        Visible = False
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MDIPrincipal.RemoverVentana Me, "M. Usuario"
End Sub

Private Sub TxtAmonestaciones_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And TxtAmonestaciones.Text <> "" Then
        ChkBloqueo.SetFocus
    End If
End Sub

Private Sub TxtBloqueos_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And TxtBloqueos.Text <> "" Then
        DPRegistro.SetFocus
    End If
End Sub

Private Sub TxtCarrera_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And TxtCarrera.Text <> "" Then
        CmdModificar.SetFocus
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

Private Sub CmbGrupo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And CmbGrupo.Text <> "" Then
        TxtNombre.SetFocus
    End If
End Sub

Private Sub CmbNivel_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And CmbNivel.Text <> "" Then
        TxtAmonestaciones.SetFocus
    End If
End Sub

Private Sub ChkBloqueo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        TxtBloqueos.SetFocus
    End If
End Sub
