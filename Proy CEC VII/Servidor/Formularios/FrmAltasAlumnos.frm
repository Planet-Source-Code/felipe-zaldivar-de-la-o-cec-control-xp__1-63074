VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form FrmAltasAlumnos 
   BackColor       =   &H00FFCEBB&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "::: Usuarios :::"
   ClientHeight    =   6570
   ClientLeft      =   1875
   ClientTop       =   2685
   ClientWidth     =   9705
   Icon            =   "FrmAltasAlumnos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6570
   ScaleWidth      =   9705
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox P1 
      BackColor       =   &H00800000&
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   0
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   33
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
      Left            =   6240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   240
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Frame FrmOpt 
      BackColor       =   &H00FFCEBB&
      BorderStyle     =   0  'None
      Height          =   3015
      Left            =   120
      TabIndex        =   18
      Top             =   3480
      Width           =   9135
      Begin VB.CommandButton CmdTodos 
         BackColor       =   &H00E89C78&
         Caption         =   "M&ostrar Todos"
         Height          =   315
         Left            =   7800
         TabIndex        =   37
         Top             =   1905
         Width           =   1215
      End
      Begin VB.CommandButton CmdModificar 
         Appearance      =   0  'Flat
         BackColor       =   &H00E89C78&
         Caption         =   "&Modificar"
         Height          =   315
         Left            =   7800
         TabIndex        =   36
         Top             =   1410
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Cancelar"
         Height          =   315
         Left            =   6480
         TabIndex        =   35
         Top             =   2400
         Width           =   2535
      End
      Begin VB.TextBox TxtFecha 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   5040
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   1506
         Width           =   1215
      End
      Begin VB.TextBox TxtBloqueos 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   5040
         Locked          =   -1  'True
         MaxLength       =   3
         TabIndex        =   9
         Top             =   1143
         Width           =   1215
      End
      Begin VB.TextBox TxtResponsable 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   5040
         Locked          =   -1  'True
         MaxLength       =   9
         TabIndex        =   11
         Top             =   1869
         Width           =   1215
      End
      Begin VB.CheckBox ChkBloqueo 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&No"
         Height          =   255
         Left            =   5040
         TabIndex        =   8
         Top             =   810
         Width           =   1215
      End
      Begin VB.TextBox TxtGrupo 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   5040
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   13
         Top             =   2595
         Width           =   1215
      End
      Begin VB.TextBox TxtGrado 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   5040
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   12
         Top             =   2232
         Width           =   1215
      End
      Begin VB.TextBox TxtNivel 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   2244
         Width           =   1935
      End
      Begin VB.TextBox TxtCGrupo 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   840
         Width           =   1935
      End
      Begin VB.TextBox TxtNombre 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1440
         Locked          =   -1  'True
         MaxLength       =   99
         TabIndex        =   3
         Top             =   1191
         Width           =   1935
      End
      Begin VB.CommandButton CmdEliminar 
         Appearance      =   0  'Flat
         BackColor       =   &H00E89C78&
         Caption         =   "&Eliminar"
         Height          =   315
         Left            =   6480
         TabIndex        =   16
         Top             =   1905
         Width           =   1215
      End
      Begin VB.CommandButton CmdAgregar 
         Appearance      =   0  'Flat
         BackColor       =   &H00E89C78&
         Caption         =   "&Agregar"
         Height          =   315
         Left            =   6480
         TabIndex        =   15
         Top             =   1410
         Width           =   1215
      End
      Begin VB.TextBox TxtCarrera 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   7200
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   14
         Top             =   795
         Width           =   1815
      End
      Begin VB.TextBox TxtAmo 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1440
         Locked          =   -1  'True
         MaxLength       =   3
         TabIndex        =   7
         Top             =   2595
         Width           =   1935
      End
      Begin VB.TextBox TxtPassword 
         BorderStyle     =   0  'None
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1440
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   5
         Top             =   1893
         Width           =   1935
      End
      Begin VB.TextBox TxtCuenta 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1440
         Locked          =   -1  'True
         MaxLength       =   9
         TabIndex        =   4
         Top             =   1542
         Width           =   1935
      End
      Begin VB.TextBox TxtBusqueda 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   3840
         TabIndex        =   1
         Text            =   "TxtBusqueda"
         Top             =   255
         Width           =   2175
      End
      Begin VB.ComboBox CmbOpcion 
         Height          =   315
         Left            =   1440
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Like"
         Height          =   255
         Left            =   3360
         TabIndex        =   34
         Top             =   270
         Width           =   375
      End
      Begin VB.Label Label16 
         BackColor       =   &H00FF0000&
         BackStyle       =   0  'Transparent
         Caption         =   "Grado:"
         Height          =   195
         Left            =   3600
         TabIndex        =   32
         Top             =   2280
         Width           =   1305
      End
      Begin VB.Label Label15 
         BackColor       =   &H00FF0000&
         BackStyle       =   0  'Transparent
         Caption         =   "Grupo.:"
         Height          =   195
         Left            =   3600
         TabIndex        =   31
         Top             =   2640
         Width           =   1305
      End
      Begin VB.Label Label14 
         BackColor       =   &H00FF0000&
         BackStyle       =   0  'Transparent
         Caption         =   "Carrera.:"
         Height          =   195
         Left            =   6480
         TabIndex        =   30
         Top             =   840
         Width           =   1305
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FF0000&
         BackStyle       =   0  'Transparent
         Caption         =   "Responsable:"
         Height          =   195
         Left            =   3600
         TabIndex        =   29
         Top             =   1920
         Width           =   1305
      End
      Begin VB.Label Label12 
         BackColor       =   &H00FF0000&
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha de registro:"
         Height          =   195
         Left            =   3600
         TabIndex        =   28
         Top             =   1560
         Width           =   1305
      End
      Begin VB.Label Label11 
         BackColor       =   &H00FF0000&
         BackStyle       =   0  'Transparent
         Caption         =   "Bloqueos:"
         Height          =   195
         Left            =   3600
         TabIndex        =   27
         Top             =   1200
         Width           =   1305
      End
      Begin VB.Label Label9 
         BackColor       =   &H00FF0000&
         BackStyle       =   0  'Transparent
         Caption         =   "Bloqueo:"
         Height          =   195
         Left            =   3600
         TabIndex        =   26
         Top             =   840
         Width           =   1305
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FF0000&
         BackStyle       =   0  'Transparent
         Caption         =   "Amonestaciones:"
         Height          =   195
         Left            =   120
         TabIndex        =   25
         Top             =   2640
         Width           =   1305
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FF0000&
         BackStyle       =   0  'Transparent
         Caption         =   "Nivel:"
         Height          =   195
         Left            =   120
         TabIndex        =   24
         Top             =   2280
         Width           =   1305
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FF0000&
         BackStyle       =   0  'Transparent
         Caption         =   "Password:"
         Height          =   195
         Left            =   120
         TabIndex        =   23
         Top             =   1920
         Width           =   1305
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FF0000&
         BackStyle       =   0  'Transparent
         Caption         =   "Cuenta:"
         Height          =   195
         Left            =   120
         TabIndex        =   22
         Top             =   1560
         Width           =   1305
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FF0000&
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre:"
         Height          =   195
         Left            =   120
         TabIndex        =   21
         Top             =   1200
         Width           =   1305
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FF0000&
         BackStyle       =   0  'Transparent
         Caption         =   "Clave de grupo:"
         Height          =   195
         Left            =   120
         TabIndex        =   20
         Top             =   840
         Width           =   1305
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Busqueda:"
         Height          =   195
         Left            =   120
         TabIndex        =   19
         Top             =   300
         Width           =   765
      End
   End
   Begin MSFlexGridLib.MSFlexGrid FG 
      Height          =   3495
      Left            =   120
      TabIndex        =   17
      Tag             =   "1"
      Top             =   120
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   6165
      _Version        =   393216
      Cols            =   3
      FixedCols       =   0
      RowHeightMin    =   100
      BackColor       =   16442835
      ForeColor       =   0
      BackColorFixed  =   15244408
      ForeColorFixed  =   16777215
      BackColorSel    =   16764603
      ForeColorSel    =   12582912
      BackColorBkg    =   16777215
      GridColor       =   16777215
      GridColorFixed  =   16777215
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      HighLight       =   2
      GridLinesFixed  =   1
      BorderStyle     =   0
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "FrmAltasAlumnos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const Modulo As String = "FrmAltasAlumnos"

Dim MensajeUsr1 As String
Dim RespuestaUsr1 As String

Private Sub CmdAgregar_Click()
    ActivarForm FrmUsrAlta
End Sub

Private Sub CmdEliminar_Click()
    If TxtCuenta.Text = "" Then MsgBox "Selecciona un usuario...", , "Atención!!!": DoEvents: Exit Sub
    MensajeUsr1 = "Estas seguro de eliminar a: " & TxtCuenta.Text
    If Preguntar(MensajeUsr1) = False Then Exit Sub
    ConexionPrincipal
    Sql = "select * from Tbl_Acceso where C_Acceso= '" & TxtCuenta.Text & "'"
    Rs.Open Sql, Conecta, adOpenDynamic, adLockBatchOptimistic
    
    If Rs.BOF = False And Rs.EOF = False Then
        Rs.Close
        Sql = "delete from Tbl_Acceso where C_Acceso= '" & TxtCuenta.Text & "'"
        Rs.Open Sql, Conecta, adOpenDynamic, adLockBatchOptimistic
        LoadFG 1
    Else
        MsgBox "Usuario no encontrado!!!", , "Atención!!!"
    End If
End Sub

Private Sub CmdModificar_Click()
    If TxtCuenta.Text = "" Then MsgBox "Selecciona un usuario...", , "Atención!!!": DoEvents: Exit Sub
    On Error Resume Next
    FrmUsrModif.TxtCT.Text = TxtCuenta.Text
    FrmUsrModif.TxtCuenta.Text = TxtCuenta.Text
    FrmUsrModif.TxtNombre.Text = TxtNombre.Text
    FrmUsrModif.TxtPassword.Text = TxtPassword.Text
    FrmUsrModif.TxtGrado.Text = TxtGrado.Text
    FrmUsrModif.TxtGrupo.Text = TxtGrado.Text
    FrmUsrModif.TxtCarrera.Text = TxtGrado.Text
    ActivarForm FrmUsrModif
    Call FrmUsrModif.Refresh
    FrmUsrModif.ChkBloqueo.Value = ChkBloqueo.Value
    FrmUsrModif.CmbNivel.Text = TxtNivel.Text
    FrmUsrModif.TxtAmonestaciones.Text = TxtAmo.Text
    FrmUsrModif.TxtBloqueos.Text = TxtBloqueos.Text
    FrmUsrModif.CmbGrupo.Text = TxtCGrupo.Text
End Sub

Private Sub CmdTodos_Click()
    LoadFG 1
End Sub

Private Sub CamposGrupo()
    DataGrupo.DatabaseName = Direccion
    DataGrupo.RecordSource = "select * from Tbl_Grupo where C_Grupo='" & TxtCGrupo.Text & "'"
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

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Command3_Click()
    If FrmGrupos.WindowState <> 2 Then
        FrmGrupos.Show
    End If
End Sub

Private Sub FG_Click()
    'If FG.MouseRow = 0 Then Exit Sub
    'Aux = FG.TextMatrix(FG.MouseRow, 0)
    'TxtBusqueda.Text = FG.TextMatrix(FG.MouseRow, FG.MouseCol)
End Sub

Private Sub Form_Activate()
    TxtResponsable.Text = FrmPrincipal.TxtUsuario.Text
End Sub

Private Sub Form_Load()
    MDIPrincipal.AgregarVentana Me, "Usuarios", "Consultas, altas, bajas, modificaciones a usuarios..."
    LLenarCMOOPT
    TxtBusqueda.Text = ""
    InsertarPicture Me
    PosicionInicial Me
End Sub

Private Sub LoadFG(Opcion As Integer, Optional StrBusqueda As String, Optional Busqueda As String)
    Dim Ancho, TamañoCol As Long
    Dim Titulos As Variant
    Dim i As Integer
    Set Rs = Nothing
    Ancho = 0
    TamañoCol = 0
    If Opcion = 1 Then
        Sql = "select * from Tbl_Acceso"
    Else
        Sql = "select * from Tbl_Acceso where " & StrBusqueda & " like " & "'%" & Busqueda & "%'"
    End If
    ConexionPrincipal
    Rs.Open Sql, Conecta, adOpenDynamic, adLockBatchOptimistic
    FG.AllowUserResizing = flexResizeBoth
    FG.Cols = Rs.Fields.Count
    FG.Rows = 1
    Titulos = Array("Grupo", "Nombre" _
              , "Cuenta", "Password", "Nivel", "Amonestaciones", "Bloqueo" _
              , "Bloqueos", "Fecha de registro", "Responsable")
              
    FG.Row = 0
    For i = 0 To Rs.Fields.Count - 1
        FG.Col = i
        FG.ColAlignment(i) = flexAlignLeftCenter
        FG.Text = Titulos(i)
        FG.ColWidth(i) = CInt(TextWidth(FG.Text) + 100)
        Ancho = Ancho + FG.ColWidth(i)
    Next
    Do While Not Rs.EOF
        FG.Rows = FG.Rows + 1
        FG.Row = FG.Rows - 1
        FG.Col = 0
        Ancho = 0
        For i = 0 To Rs.Fields.Count - 1
            FG.Col = i
            FG.Text = Rs(i).Value & ""
            TamañoCol = FG.ColWidth(i)
            If CInt(TextWidth(FG.Text) + 100) > TamañoCol Then
                FG.ColWidth(i) = CInt(TextWidth(FG.Text) + 100)
            End If
            If FG.Row / 2 <> Int(FG.Row / 2) Then
                FG.CellBackColor = RGB(194, 208, 252)
            End If
        Next
        Rs.MoveNext
    Loop
    FG.Width = Me.Width - 340
    FrmOpt.Top = FG.Top + FG.Height
    FrmOpt.Left = (Me.Width / 2) - FrmOpt.Width / 2
    Me.Height = (FrmOpt.Height + FrmOpt.Top + FG.Top + 340)
    If Not Rs.EOF Or Not Rs.BOF Then
        Rs.MoveFirst
        LlenarCampos
        FG.FixedRows = 1
    Else
        FG.FixedRows = 0
        VaciarCampos
    End If
End Sub

Private Sub VaciarCampos()
    TxtCGrupo.Text = ""
    TxtGrupo.Text = ""
    TxtGrado.Text = ""
    TxtGrupo.Text = ""
    TxtCarrera.Text = ""
    TxtNombre.Text = ""
    TxtCuenta.Text = ""
    TxtPassword.Text = ""
    TxtNivel.Text = ""
    ChkBloqueo.Value = 0
    TxtFecha.Text = ""
    TxtAmo.Text = ""
    ChkBloqueo.Caption = "No"
    TxtBloqueos.Text = ""
    TxtResponsable.Text = ""
End Sub

Private Sub LlenarCampos()
    'On Error GoTo Error
    Const Evento As String = "LlenarCampos"
    On Error Resume Next
    TxtCGrupo.Text = Rs!C_Grupo
    TxtNombre.Text = Rs!Nombre
    TxtCuenta.Text = Rs!C_Acceso
    TxtPassword.Text = Rs!Password
    TxtNivel.Text = CStr(Rs!Nivel)
    TxtAmo.Text = Rs!Amonestaciones
    If Rs!Usr_Bloqueado = True Then
        ChkBloqueo.Value = 1
        ChkBloqueo.Caption = "Sí"
    Else
        ChkBloqueo.Value = 0
        ChkBloqueo.Caption = "No"
    End If
    TxtBloqueos.Text = Rs!N_Bloqueos
    TxtFecha.Text = CDate(Rs!Fecha_Reg)
    TxtResponsable.Text = Rs!C_U_Registro
    CamposGrupo
End Sub

Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then
        Visible = False
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set Rs = Nothing
    Set FrmAltasAlumnos = Nothing
    MDIPrincipal.RemoverVentana Me, "Usuarios"
End Sub

Private Sub TxtAmo_LostFocus()
    If Not IsNumeric(TxtAmo.Text) Then
        TxtAmo.Text = "0"
    End If
End Sub

Private Sub TxtBloqueos_LostFocus()
    If Not IsNumeric(TxtBloqueos.Text) Then
        TxtAmo.Text = "0"
    End If
End Sub

Private Sub TxtBusqueda_Change()
    Dim StrBusqueda As String
    StrBusqueda = CmbOpcion.Text
    If CmbOpcion.Text = "Clave de Grupo" Then
        StrBusqueda = "C_Grupo"
    ElseIf CmbOpcion.Text = "Cuenta" Then
        StrBusqueda = "C_Acceso"
    ElseIf CmbOpcion.Text = "Nombre" Then
        StrBusqueda = "Nombre"
    End If
    If TxtBusqueda.Text = "" Or Len(TxtBusqueda.Text) = 0 Then
        LoadFG 1 'busque todos
    Else
        Call LoadFG(2, StrBusqueda, TxtBusqueda.Text) 'busque solo los especificados
    End If
End Sub

Private Sub LLenarCMOOPT()
    CmbOpcion.AddItem "Clave de Grupo"
    CmbOpcion.AddItem "Cuenta"
    CmbOpcion.AddItem "Nombre"
    CmbOpcion.Text = "Cuenta"
End Sub
