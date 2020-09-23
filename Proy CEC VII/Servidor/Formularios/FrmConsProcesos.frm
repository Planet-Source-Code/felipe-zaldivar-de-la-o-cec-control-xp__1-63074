VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FrmConsProcesos 
   BackColor       =   &H00FFCEBB&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "::: Procesos :::"
   ClientHeight    =   6810
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7545
   Icon            =   "FrmConsProcesos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6810
   ScaleWidth      =   7545
   Begin VB.PictureBox PicEscuela 
      Height          =   495
      Left            =   7560
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   23
      Top             =   120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox P1 
      BackColor       =   &H00800000&
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   0
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   22
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox TxtPP 
      BorderStyle     =   0  'None
      Height          =   885
      Left            =   885
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   11
      Top             =   5805
      Width           =   6562
   End
   Begin VB.TextBox TxtPH 
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   5985
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   10
      Top             =   5415
      Width           =   1455
   End
   Begin VB.TextBox TxtPF 
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   5985
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   8
      Top             =   5040
      Width           =   1455
   End
   Begin VB.TextBox TxtPU 
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   885
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   7
      Top             =   5040
      Width           =   4355
   End
   Begin VB.TextBox TxtPM 
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   885
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   9
      Top             =   5415
      Width           =   4355
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   120
      TabIndex        =   13
      Top             =   3840
      Width           =   7320
      Begin VB.CommandButton CmdPB 
         Caption         =   "Procesos Bloqueados"
         Height          =   255
         Left            =   3960
         TabIndex        =   5
         Top             =   720
         Width           =   2175
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
      Begin VB.TextBox TxtBR 
         Height          =   285
         Left            =   2760
         TabIndex        =   2
         Top             =   240
         Width           =   4455
      End
      Begin VB.CommandButton CmdTodos 
         Caption         =   "&Todos"
         Height          =   255
         Left            =   6240
         TabIndex        =   6
         Top             =   720
         Width           =   975
      End
      Begin MSComCtl2.DTPicker DT2 
         CausesValidation=   0   'False
         Height          =   285
         Left            =   5760
         TabIndex        =   4
         Top             =   255
         Visible         =   0   'False
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
         _Version        =   393216
         Format          =   47710211
         CurrentDate     =   38601.9583333333
      End
      Begin MSComCtl2.DTPicker DT1 
         CausesValidation=   0   'False
         Height          =   285
         Left            =   3480
         TabIndex        =   3
         Top             =   255
         Visible         =   0   'False
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
         _Version        =   393216
         Format          =   47710211
         CurrentDate     =   38601
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha 1:"
         Height          =   195
         Left            =   2760
         TabIndex        =   16
         Top             =   300
         Visible         =   0   'False
         Width           =   630
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha 2:"
         Height          =   195
         Left            =   5040
         TabIndex        =   15
         Top             =   300
         Visible         =   0   'False
         Width           =   630
      End
      Begin VB.Label LblCampo 
         BackColor       =   &H0080FFFF&
         Caption         =   "LblCampo"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   720
         Visible         =   0   'False
         Width           =   855
      End
   End
   Begin MSFlexGridLib.MSFlexGrid FGR 
      Height          =   3735
      Left            =   120
      TabIndex        =   12
      Tag             =   "1"
      Top             =   120
      Width           =   7320
      _ExtentX        =   12912
      _ExtentY        =   6588
      _Version        =   393216
      Rows            =   1
      Cols            =   5
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
   Begin VB.Label Label25 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hora:"
      Height          =   195
      Left            =   5340
      TabIndex        =   21
      Top             =   5460
      Width           =   390
   End
   Begin VB.Label Label26 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha:"
      Height          =   195
      Left            =   5340
      TabIndex        =   20
      Top             =   5085
      Width           =   495
   End
   Begin VB.Label Label27 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Usuario:"
      Height          =   195
      Left            =   120
      TabIndex        =   19
      Top             =   5085
      Width           =   585
   End
   Begin VB.Label Label28 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Proceso:"
      Height          =   195
      Left            =   120
      TabIndex        =   18
      Top             =   5850
      Width           =   630
   End
   Begin VB.Label Label29 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Maquina:"
      Height          =   195
      Left            =   120
      TabIndex        =   17
      Top             =   5460
      Width           =   660
   End
End
Attribute VB_Name = "FrmConsProcesos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim NEscuela As String
Dim CEscuela As String

Public Sub CargarReportes(Opcion As Integer, Optional Campo As String, Optional Operador As String, Optional StrBusqueda As String, Optional Fecha1 As Date, Optional Fecha2 As Date)
    Dim Ancho, TamañoCol As Long
    Dim Titulos As Variant
    Dim i As Integer
    
    Ancho = 0
    TamañoCol = 0
    
    If Opcion = 1 Then
        SqlRep = "select * from Tbl_Procesos_Reg ORDER BY Fecha desc, Hora desc"
    ElseIf Opcion = 2 Then
        SqlRep = "select * from Tbl_Procesos_Reg where " & Campo & " " & Operador & " " & "'%" & StrBusqueda & "%'"
    ElseIf Opcion = 3 Then
        SqlRep = "select * from Tbl_Procesos_Reg where " & Campo & " " & Operador & " " & "#" & Format(Fecha1, "mm/dd/yy") & "#"
    ElseIf Opcion = 4 Then
        SqlRep = "select * from Tbl_Procesos_Reg where " & Campo & " Between #" & Format(Fecha1, "mm/dd/yy") & "# and #" & Format(Fecha2, "mm/dd/yy") & "#"
     ElseIf Opcion = 5 Then
        SqlRep = "select * from Tbl_Procesos_Reg where " & Campo & " " & Operador & " " & "#" & Format(Fecha1, "hh:mm:ss am/pm") & "#"
    ElseIf Opcion = 6 Then
        SqlRep = "select * from Tbl_Procesos_Reg where " & Campo & " Between #" & Format(Fecha1, "hh:mm:ss am/pm") & "# and #" & Format(Fecha2, "hh:mm:ss am/pm") & "# "
    ElseIf Opcion = 7 Then
        SqlRep = "select * from Tbl_Procesos_Reg where " & Campo & " " & Operador & " '" & StrBusqueda & "'"
    End If

    ConexionProcesos
    RsCP.Open SqlRep, Conecta, adOpenDynamic, adLockBatchOptimistic
    
    FGR.AllowUserResizing = flexResizeBoth
    FGR.Rows = 1
    
    Titulos = Array("Usuario", "Maquina", "Fecha", "Hora", "Proceso")
              
    FGR.Row = 0
    For i = 0 To RsCP.Fields.Count - 1
        FGR.Col = i
        FGR.ColAlignment(i) = flexAlignLeftCenter
        FGR.Text = Titulos(i)
        FGR.ColWidth(i) = CInt(TextWidth(FGR.Text) + 300)
        Ancho = Ancho + FGR.ColWidth(i)
    Next
    
    Do While Not RsCP.EOF
        FGR.Rows = FGR.Rows + 1
        FGR.Row = FGR.Rows - 1
        FGR.Col = 0
        Ancho = 0
        For i = 0 To RsCP.Fields.Count - 1
            FGR.Col = i
            FGR.Text = RsCP(i).Value & ""
            TamañoCol = FGR.ColWidth(i)
            If CInt(TextWidth(FGR.Text) + 100) > TamañoCol Then
                FGR.ColWidth(i) = CInt(TextWidth(FGR.Text) + 150)
            End If
            If FGR.Row / 2 <> Int(FGR.Row / 2) Then
                FGR.CellBackColor = RGB(194, 208, 252)
            End If
        Next
        RsCP.MoveNext
    Loop
    
    If Not RsCP.EOF Or Not RsCP.BOF Then
        RsCP.MoveFirst
        LlenarCamposProceso
        FGR.FixedRows = 1
    Else
        FGR.FixedRows = 0
        VaciarCamposProceso
    End If
End Sub

Private Sub LlenarCamposProceso()
    TxtPU.Text = RsCP!C_Acceso
    TxtPM.Text = RsCP!C_Maq
    TxtPF.Text = RsCP!Fecha
    TxtPH.Text = RsCP!Hora
    TxtPP.Text = RsCP!Proceso
End Sub

Private Sub VaciarCamposProceso()
    TxtPU.Text = ""
    TxtPM.Text = ""
    TxtPF.Text = ""
    TxtPH.Text = ""
    TxtPP.Text = ""
End Sub

Private Sub CargarCR()
    CboCampo.Clear
    CboCampo.AddItem ("Maquina")
    CboCampo.AddItem ("Usuario")
    CboCampo.AddItem ("Fecha")
    CboCampo.AddItem ("Hora")
    CboCampo.AddItem ("Proceso")
    CboCampo.Text = CboCampo.List(0)
    
    CboCampo_Click
End Sub

Private Sub CboCampo_Click()
    CboOperador.Clear
    If CboCampo.Text = "Maquina" Then
        LblCampo.Caption = "C_Maq"
        AgregarItem 1
        Exit Sub
    ElseIf CboCampo.Text = "Usuario" Then
        LblCampo.Caption = "C_Acceso"
        AgregarItem 1
        Exit Sub
    ElseIf CboCampo.Text = "Fecha" Then
        LblCampo.Caption = "Fecha"
        AgregarItem 2
        DT1.Format = dtpCustom
        DT2.Format = dtpCustom
        DT1.Value = Date
        DT2.Value = Date
        DT2.Visible = True
        DT2.Visible = False
        Label30.Caption = "Fecha 1:"
        Label31.Caption = "Fecha 2:"
        Label30.Visible = True
        Label31.Visible = False
        Exit Sub
    ElseIf CboCampo.Text = "Hora" Then
        LblCampo.Caption = "Hora"
        AgregarItem 2
        DT1.Format = dtpTime
        DT2.Format = dtpTime
        DT1.Value = Time
        DT2.Value = Time
        DT2.Visible = True
        DT2.Visible = False
        Label30.Caption = "Hora 1:"
        Label31.Caption = "Hora 2:"
        Label30.Visible = True
        Label31.Visible = False
        Exit Sub
    Else
        LblCampo.Caption = "Proceso"
        AgregarItem 1
        Exit Sub
    End If
End Sub

Private Sub AgregarItem(Opcion As Integer)
    CboOperador.Clear
    If Opcion = 1 Then
        TxtBR.Visible = True
        DT1.Visible = False
        DT2.Visible = False
        CboOperador.AddItem ("=")
        CboOperador.AddItem ("<>")
        CboOperador.AddItem ("Like")
        CboOperador.Text = CboOperador.List(0)
        DT2.Visible = False
        DT2.Visible = False
        Label30.Visible = False
        Label31.Visible = False
    Else
        TxtBR.Visible = False
        CboOperador.AddItem ("=")
        CboOperador.AddItem ("<>")
        CboOperador.AddItem ("<")
        CboOperador.AddItem (">")
        CboOperador.AddItem ("<=")
        CboOperador.AddItem (">=")
        CboOperador.AddItem ("Between")
        CboOperador.AddItem ("Like")
        CboOperador.Text = CboOperador.List(0)
        DT1.Format = dtpTime: DT1.Visible = True
    End If
End Sub

Private Sub CmdPB_Click()
    On Error Resume Next
    
    If ValidarImpresora = True Then
        POLandscape Me
        With DE_Procesos
            .ConPrincipal.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & (App.Path & "\BD\Bdcec1.mdb") & ";Mode=Read;Persist Security Info=False;Jet OLEDB:Database Password=control"
            .Connections("ConPrincipal").Open
            .rsTbl_Acceso.Open
            .rsTbl_Acceso.Requery
        End With
        With DR_PB
            Set .Sections(1).Controls("ImgLogo").Picture = PicEscuela.Picture
            .Sections(1).Controls("EtqEscuela").Caption = NEscuela
            .Sections(1).Controls("EtqCEscuela").Caption = CEscuela
            .WindowState = 2
            .Show
        End With
    End If
End Sub

Private Function ValidarImpresora() As Boolean
    If InStr(1, LCase(Printer.DeviceName), "microsoft") <> 0 Then
        MsgBox Printer.DeviceName & " no es una impresora válida...", , "Atención!!!": DoEvents
        ValidarImpresora = False
    Else
        ValidarImpresora = True
    End If
End Function

Private Sub CmdTodos_Click()
    CargarReportes 1
End Sub

Private Sub DT1_Click()
    DT1Event
End Sub

Private Sub DT1_Change()
    DT1Event
End Sub

Private Sub DT1Event()
    If LblCampo.Caption = "Hora" Then
        If Not IsDate(DT1.Value) Then MsgBox "Hora incorrecta!!!": Exit Sub
        If CboOperador.Text = "Between" Then
            Call CargarReportes(6, LblCampo.Caption, CboOperador.Text, , DT1.Value, DT2.Value)
        Else
            Call CargarReportes(5, LblCampo.Caption, CboOperador.Text, , DT1.Value)
        End If
        Exit Sub
    End If
    If Not IsDate(DT2.Value) Then MsgBox "Fecha incorrecta!!!": Exit Sub
    If CboOperador.Text = "Between" Then
        Call CargarReportes(4, LblCampo.Caption, CboOperador.Text, , DT1.Value, DT2.Value)
    Else
        Call CargarReportes(3, LblCampo.Caption, CboOperador.Text, , DT1.Value)
    End If
End Sub

Private Sub DT2_Click()
    DT2Event
End Sub

Private Sub DT2_Change()
    DT1Event
End Sub

Private Sub DT2Event()
    If LblCampo.Caption = "Hora" Then
        If Not IsDate(DT2.Value) Then MsgBox "Hora incorrecta!!!": Exit Sub
        If CboOperador.Text = "Between" Then
            Call CargarReportes(6, LblCampo.Caption, CboOperador.Text, , DT1.Value, DT2.Value)
        Else
            Call CargarReportes(5, LblCampo.Caption, CboOperador.Text, , DT2.Value)
        End If
        Exit Sub
    End If
    If Not IsDate(DT2.Value) Then MsgBox "Fecha incorrecta!!!": Exit Sub
    If CboOperador.Text = "Between" Then
        Call CargarReportes(4, LblCampo.Caption, CboOperador.Text, , DT1.Value, DT2.Value)
    Else
        Call CargarReportes(3, LblCampo.Caption, CboOperador.Text, , DT1.Value)
    End If
End Sub

Private Sub Form_Load()
    MDIPrincipal.AgregarVentana Me, "C. Procesos", "Procesos registrados..."
    Call CargarReportes(1)
    Call CargarInf
    Call CargarCR
    InsertarPicture Me
    PosicionInicial Me
End Sub

Public Sub CargarInf()
On Error Resume Next
    Dim logo1() As Byte
    Dim PBag As PropertyBag
    ConexionEscuelas
    SQLCE = "Select * from Tbl_Escuela where IDESC='CBT2'"
    RsCE.Open SQLCE, Conecta, adOpenDynamic, adLockBatchOptimistic
    If Not RsCE.EOF Then
        On Error Resume Next
        If Not IsNull(RsCE.Fields("logotipo1").Value) Then
            logo1 = RsCE.Fields("logotipo1").Value
            Set PBag = New PropertyBag
            PBag.Contents = logo1
            Set PicEscuela = PBag.ReadProperty("CBT2")
        End If
        NEscuela = RsCE.Fields("Escuela").Value
        CEscuela = RsCE.Fields("C_Escuela").Value
    End If
End Sub

Private Sub CboOperador_Click()
    If CboOperador.Text = "Between" Then
        DT2.Visible = True
        Label31.Visible = True
    Else
        DT2.Visible = False
        Label31.Visible = False
    End If
End Sub

Private Sub FGR_Click()
    If FGR.MouseRow = 0 Then Exit Sub
    TxtPU.Text = FGR.TextMatrix(FGR.MouseRow, 0)
    TxtPM.Text = FGR.TextMatrix(FGR.MouseRow, 1)
    TxtPF.Text = FGR.TextMatrix(FGR.MouseRow, 2)
    TxtPH.Text = FGR.TextMatrix(FGR.MouseRow, 3)
    TxtPP.Text = FGR.TextMatrix(FGR.MouseRow, 4)
End Sub

Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then
        Visible = False
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MDIPrincipal.RemoverVentana Me, "C. Procesos"
End Sub


Private Sub TxtBR_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If TxtBR.Text = "" Then CargarReportes 1: Exit Sub
        If CboOperador.Text = "Like" Then
            Call CargarReportes(2, LblCampo.Caption, CboOperador.Text, TxtBR.Text)
        Else
            Call CargarReportes(7, LblCampo.Caption, CboOperador.Text, TxtBR.Text)
        End If
    End If
End Sub
