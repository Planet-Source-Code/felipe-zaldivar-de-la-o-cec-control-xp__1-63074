VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form FrmConsMaquina 
   BackColor       =   &H00FFCEBB&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "::: Maquinas  :::"
   ClientHeight    =   5535
   ClientLeft      =   3255
   ClientTop       =   2910
   ClientWidth     =   4965
   Icon            =   "FrmConsMaquina.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5535
   ScaleWidth      =   4965
   Begin VB.PictureBox P1 
      BackColor       =   &H00800000&
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   0
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   17
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H00FFCEBB&
      BorderStyle     =   0  'None
      Height          =   2055
      Left            =   120
      TabIndex        =   13
      Top             =   3360
      Width           =   4730
      Begin VB.CommandButton CmdTodos 
         Caption         =   "&Todos"
         Height          =   255
         Left            =   3480
         TabIndex        =   11
         Top             =   1680
         Width           =   1080
      End
      Begin VB.CommandButton CmdAgregar 
         Caption         =   "&Agregar"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   1680
         Width           =   1080
      End
      Begin VB.CommandButton CmdEliminar 
         Caption         =   "&Eliminar"
         Height          =   255
         Left            =   2360
         TabIndex        =   10
         Top             =   1680
         Width           =   1080
      End
      Begin VB.CommandButton CmdModificar 
         Caption         =   "&Modificar"
         Height          =   255
         Left            =   1240
         TabIndex        =   9
         Top             =   1680
         Width           =   1080
      End
      Begin VB.CheckBox ChkMB 
         BackColor       =   &H00FFCEBB&
         Caption         =   "&Bloqueada"
         Height          =   255
         Left            =   3480
         TabIndex        =   6
         Top             =   735
         Width           =   1095
      End
      Begin VB.CheckBox ChkMO 
         BackColor       =   &H00FFCEBB&
         Caption         =   "&Ocupada"
         Height          =   255
         Left            =   3480
         TabIndex        =   7
         Top             =   1155
         Width           =   1095
      End
      Begin VB.TextBox TxtMM 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   885
         MaxLength       =   50
         MultiLine       =   -1  'True
         TabIndex        =   5
         Top             =   1140
         Width           =   2550
      End
      Begin VB.TextBox TxtMU 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   885
         MaxLength       =   9
         MultiLine       =   -1  'True
         TabIndex        =   4
         Top             =   720
         Width           =   2550
      End
      Begin VB.CheckBox ChkB 
         BackColor       =   &H00FFCEBB&
         Caption         =   "Falso"
         Height          =   255
         Left            =   2880
         TabIndex        =   3
         Top             =   270
         Width           =   1575
      End
      Begin VB.ComboBox CboCampo 
         Height          =   315
         Left            =   120
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   240
         Width           =   1215
      End
      Begin VB.ComboBox CboOperador 
         Height          =   315
         Left            =   1440
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox TxtBM 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   2760
         TabIndex        =   2
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Maquina:"
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   1185
         Width           =   660
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Usuario:"
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   765
         Width           =   585
      End
      Begin VB.Label LblCampo 
         BackColor       =   &H0080FFFF&
         Caption         =   "LblCampo"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   960
         Visible         =   0   'False
         Width           =   855
      End
   End
   Begin MSFlexGridLib.MSFlexGrid FGR 
      Height          =   3255
      Left            =   120
      TabIndex        =   12
      Tag             =   "1"
      Top             =   120
      Width           =   4730
      _ExtentX        =   8334
      _ExtentY        =   5741
      _Version        =   393216
      Rows            =   1
      Cols            =   4
      FixedRows       =   0
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
      Redraw          =   -1  'True
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      GridLinesFixed  =   1
      AllowUserResizing=   3
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
Attribute VB_Name = "FrmConsMaquina"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim MensajeUsr1 As String
Dim RespuestaUsr1 As String

Private Sub CargarMaquinas(Opcion As Integer, Optional Campo As String, Optional Operador As String, Optional StrBusqueda As String)
    Dim Ancho, TamañoCol As Long
    Dim Titulos As Variant
    Dim i As Integer
    
    Ancho = 0
    TamañoCol = 0
    
    If Opcion = 1 Then
        SqlCM = "select * from Tbl_Maquina Order by C_Maq asc"
    ElseIf Opcion = 2 Then
        SqlCM = "select * from Tbl_Maquina where " & Campo & " " & Operador & "'%" & StrBusqueda & "%' Order by C_Maq asc"
    ElseIf Opcion = 3 Then
        SqlCM = "select * from Tbl_Maquina where " & Campo & " " & Operador & "'" & StrBusqueda & "' Order by C_Maq asc"
    ElseIf Opcion = 4 Then
        If StrBusqueda = 1 Then
            StrBusqueda = -1
        End If
        SqlCM = "select * from Tbl_Maquina where " & Campo & " " & Operador & " " & StrBusqueda & " Order by C_Maq asc"
    End If
    ConexionMaquinas
    RsCM.Open SqlCM, Conecta, adOpenDynamic, adLockBatchOptimistic
    
    FGR.AllowUserResizing = flexResizeBoth
    FGR.Rows = 1
    
    Titulos = Array("Maquina", "Bloqueo", "Ocupada", "Usuario")
              
    FGR.Row = 0
    For i = 0 To 3
        FGR.Col = i
        FGR.ColAlignment(i) = flexAlignLeftCenter
        FGR.Text = Titulos(i)
        FGR.ColWidth(i) = CInt(TextWidth(FGR.Text) + 300)
        Ancho = Ancho + FGR.ColWidth(i)
    Next
    
    Do While Not RsCM.EOF
        FGR.Rows = FGR.Rows + 1
        FGR.Row = FGR.Rows - 1
        FGR.Col = 0
        Ancho = 0
        For i = 0 To 3
            FGR.Col = i
            FGR.Text = RsCM(i).Value & ""
            TamañoCol = FGR.ColWidth(i)
            If CInt(TextWidth(FGR.Text) + 100) > TamañoCol Then
                FGR.ColWidth(i) = CInt(TextWidth(FGR.Text) + 150)
            End If
            If FGR.Row / 2 <> Int(FGR.Row / 2) Then
                FGR.CellBackColor = RGB(194, 208, 252)
            End If
        Next
        RsCM.MoveNext
    Loop

    If Not RsCM.EOF Or Not RsCM.BOF Then
        RsCM.MoveFirst
        FGR.FixedRows = 1
        LlenarCamposProceso
    Else
        FGR.FixedRows = 0
        VaciarCamposProceso
    End If
End Sub

Private Sub LlenarCamposProceso()
On Error Resume Next
    Dim CMB As Integer
    Dim CMO As Integer
    TxtMU.Text = RsCM!C_Acceso
    TxtMM.Text = RsCM!C_Maq
    CMB = RsCM!Maq_Bloqueada
    CMO = RsCM!Maq_Ocupada
    If CMB = -1 Then CMB = 1
    If CMO = -1 Then CMO = 1
    ChkMB.Value = CMB
    ChkMO.Value = CMO
End Sub

Private Sub VaciarCamposProceso()
    TxtMU.Text = ""
    TxtMM.Text = ""
    ChkMB.Value = 0
    ChkMO.Value = 0
End Sub

Private Sub CargarCR()
    CboCampo.Clear
    CboCampo.AddItem ("Maquina")
    CboCampo.AddItem ("Usuario")
    CboCampo.AddItem ("Ocupada")
    CboCampo.AddItem ("Bloqueada")
    CboCampo.Text = CboCampo.List(0)
    CboCampo_Click
End Sub

Private Sub CboCampo_Click()
    CboOperador.Clear
    If CboCampo.Text = "Maquina" Then
        LblCampo.Caption = "C_Maq"
        CboOperador.AddItem ("Like")
        ChkB.Visible = False
        TxtBM.Visible = True
    ElseIf CboCampo.Text = "Usuario" Then
        LblCampo.Caption = "C_Acceso"
        CboOperador.AddItem ("Like")
        ChkB.Visible = False
        TxtBM.Visible = True
    ElseIf CboCampo.Text = "Bloqueada" Then
        LblCampo.Caption = "Maq_Bloqueada"
        ChkB.Visible = True
        TxtBM.Visible = False
    ElseIf CboCampo.Text = "Ocupada" Then
        LblCampo.Caption = "Maq_Ocupada"
        ChkB.Visible = True
        TxtBM.Visible = False
    End If
    CboOperador.AddItem ("=")
    CboOperador.AddItem ("<>")
    CboOperador.Text = CboOperador.List(0)
End Sub

Private Sub ChkB_Click()
    If ChkB.Value = 1 Then
        ChkB.Caption = "Verdadero"
    Else
        ChkB.Caption = "Falso"
    End If
    Call CargarMaquinas(4, LblCampo.Caption, CboOperador.Text, ChkB.Value)
End Sub

Private Sub CmdAgregar_Click()
    If TxtMM.Text = "" Then MsgBox "Debes escribir el nombre de la Maquina...", , "Atención!!!": Exit Sub
    MensajeUsr1 = "Los datos son correctos?"
    If Preguntar(MensajeUsr1) = False Then Exit Sub
    
    ConexionMaquinas
    SqlCM = "select * from Tbl_Maquina where C_Maq= '" & TxtMM.Text & "'"
    RsCM.Open SqlCM, Conecta, adOpenDynamic, adLockBatchOptimistic
    
    If RsCM.BOF = False And RsCM.EOF = False Then
        MsgBox "Maquina ya registrada!!!", , "Atención!!!"
    Else
        RsCM.Close
        SqlCM = "insert into Tbl_Maquina ([C_Maq],[Maq_Bloqueada]) " _
        & "VALUES ('" & UCase(TxtMM.Text) & "','" & ChkMB.Value & "')"
        RsCM.Open SqlCM, Conecta, adOpenDynamic, adLockBatchOptimistic
        CargarMaquinas 1
        Call FrmPrincipal.AgregarMaquinasLVM
        'Call FrmNetwork.ObtenerCompus
    End If
End Sub

Private Sub CmdEliminar_Click()
    MensajeUsr1 = "Estas seguro de eliminar la Maquina: " & TxtMM.Text
    If Preguntar(MensajeUsr1) = False Then Exit Sub
 
    ConexionMaquinas
    SqlCM = "select * from Tbl_Maquina where C_Maq= '" & TxtMM.Text & "'"
    RsCM.Open SqlCM, Conecta, adOpenDynamic, adLockBatchOptimistic
    
    If RsCM.BOF = False And RsCM.EOF = False Then
        RsCM.Close
        SqlCM = "delete from Tbl_Maquina where C_Maq= '" & TxtMM.Text & "'"
        RsCM.Open SqlCM, Conecta, adOpenDynamic, adLockBatchOptimistic
        CargarMaquinas 1
        Call FrmPrincipal.AgregarMaquinasLVM
        'Call FrmNetwork.ObtenerCompus
    Else
        MsgBox "Maquina no encontrada!!!", , "Atención!!!"
    End If
End Sub

Private Sub CmdModificar_Click()
    If TxtMM.Text = "" Then MsgBox "Debes escribir el nombre de la Maquina...", , "Atención!!!": Exit Sub
    MensajeUsr1 = "Los datos son correctos?"
    If Preguntar(MensajeUsr1) = False Then Exit Sub
    
    ConexionMaquinas
    SqlCM = "select * from Tbl_Maquina where C_Maq= '" & TxtMM.Text & "'"
    RsCM.Open SqlCM, Conecta, adOpenDynamic, adLockBatchOptimistic
    
    If RsCM.BOF = False And RsCM.EOF = False Then
        If RsCM!Maq_Ocupada = True Then
            MsgBox "No se puede modificar esta máquina mientras este ocupada...", , "Atención!!!": Exit Sub
        Else
            RsCM.Close
            SqlCM = "Update Tbl_Maquina set C_Maq='" & TxtMM.Text & "',Maq_Bloqueada='" _
            & ChkMB.Value & "' where C_Maq='" & TxtMM.Text & "' and Maq_Ocupada=False"
            RsCM.Open SqlCM, Conecta, adOpenDynamic, adLockBatchOptimistic
            CargarMaquinas 1
            Call FrmPrincipal.AgregarMaquinasLVM
            DoEvents
        End If
    Else
        MsgBox "Maquina no encontrada...", , "Atención!!!"
    End If
End Sub

Private Sub CmdTodos_Click()
    CargarMaquinas 1
End Sub

Private Sub Form_Activate()
    Call CargarMaquinas(1)
End Sub

Private Sub Form_Load()
    MDIPrincipal.AgregarVentana Me, "C. Maquinas", "Altas, bajas, modificaciones a maquinas..."
    Call CargarCR
    InsertarPicture Me
    PosicionInicial Me
End Sub

Private Sub FGR_Click()
    Dim CMB As Boolean
    Dim CMO As Boolean
    If FGR.MouseRow = 0 Then Exit Sub
    TxtMM.Text = FGR.TextMatrix(FGR.MouseRow, 0)
    TxtMU.Text = FGR.TextMatrix(FGR.MouseRow, 3)
    CMB = FGR.TextMatrix(FGR.MouseRow, 1)
    CMO = FGR.TextMatrix(FGR.MouseRow, 2)
    If CMB = True Then
        ChkMB.Value = 1
    Else
        ChkMB.Value = 0
    End If
    If CMO = True Then
        ChkMO.Value = 1
    Else
        ChkMO.Value = 0
    End If
End Sub

Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then
        Visible = False
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MDIPrincipal.RemoverVentana Me, "C. Maquinas"
End Sub

Private Sub TxtBM_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If TxtBM.Text = "" Then CargarMaquinas 1: Exit Sub
        If CboOperador.Text = "Like" Then
            Call CargarMaquinas(2, LblCampo.Caption, CboOperador.Text, TxtBM.Text)
        Else
            Call CargarMaquinas(3, LblCampo.Caption, CboOperador.Text, TxtBM.Text)
        End If
    End If
End Sub
