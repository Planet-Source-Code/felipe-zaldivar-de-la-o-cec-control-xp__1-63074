VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FrmHistorial 
   BackColor       =   &H00FFCEBB&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "::: Historial de usuarios :::"
   ClientHeight    =   5775
   ClientLeft      =   45
   ClientTop       =   795
   ClientWidth     =   7530
   Icon            =   "FrmHistorial.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5775
   ScaleWidth      =   7530
   Begin VB.TextBox TxtHHS 
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   6240
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   5415
      Width           =   1200
   End
   Begin VB.TextBox TxtHHE 
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   3600
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   5415
      Width           =   1200
   End
   Begin VB.TextBox TxtHM 
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4440
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   7
      Top             =   5040
      Width           =   3000
   End
   Begin VB.TextBox TxtHU 
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   885
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   6
      Top             =   5040
      Width           =   2535
   End
   Begin VB.TextBox TxtHF 
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   885
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   5415
      Width           =   1200
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   120
      TabIndex        =   13
      Top             =   3840
      Width           =   7320
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
         TabIndex        =   5
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
   Begin VB.PictureBox P1 
      BackColor       =   &H00800000&
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   0
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   12
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin MSFlexGridLib.MSFlexGrid FGR 
      Height          =   3735
      Left            =   120
      TabIndex        =   11
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
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Maquina:"
      Height          =   195
      Left            =   3600
      TabIndex        =   21
      Top             =   5085
      Width           =   660
   End
   Begin VB.Label Label25 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hora de salida:"
      Height          =   195
      Left            =   4920
      TabIndex        =   20
      Top             =   5460
      Width           =   1065
   End
   Begin VB.Label Label26 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hora de entrada:"
      Height          =   195
      Left            =   2160
      TabIndex        =   19
      Top             =   5460
      Width           =   1200
   End
   Begin VB.Label Label27 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Usuario:"
      Height          =   195
      Left            =   120
      TabIndex        =   18
      Top             =   5085
      Width           =   585
   End
   Begin VB.Label Label29 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha:"
      Height          =   195
      Left            =   120
      TabIndex        =   17
      Top             =   5460
      Width           =   495
   End
End
Attribute VB_Name = "FrmHistorial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub CargarReportes(Opcion As Integer, Optional Campo As String, Optional Operador As String, Optional StrBusqueda As String, Optional Fecha1 As Date, Optional Fecha2 As Date)
    Dim Ancho, TamañoCol As Long
    Dim Titulos As Variant
    Dim i As Integer

    Ancho = 0
    TamañoCol = 0
    
    If Opcion = 1 Then
        SqlCH = "select * from Tbl_Historial ORDER BY Fecha_Entrada desc, Hora_Entrada asc"
    ElseIf Opcion = 2 Then
        SqlCH = "select * from Tbl_Historial where " & Campo & " " & Operador & " " & "'%" & StrBusqueda & "%' ORDER BY Fecha_Entrada desc, Hora_Entrada asc"
    ElseIf Opcion = 3 Then
        SqlCH = "select * from Tbl_Historial where " & Campo & " " & Operador & " " & "#" & Format(Fecha1, "mm/dd/yy") & "# ORDER BY Fecha_Entrada desc, Hora_Entrada asc"
    ElseIf Opcion = 4 Then
        SqlCH = "select * from Tbl_Historial where " & Campo & " Between #" & Format(Fecha1, "mm/dd/yy") & "# and #" & Format(Fecha2, "mm/dd/yy") & "#ORDER BY Fecha_Entrada desc, Hora_Entrada asc"
     ElseIf Opcion = 5 Then
        SqlCH = "select * from Tbl_Historial where " & Campo & " " & Operador & " " & "#" & Format(Fecha1, "hh:mm:ss am/pm") & "# ORDER BY Fecha_Entrada desc, Hora_Entrada asc"
    ElseIf Opcion = 6 Then
        SqlCH = "select * from Tbl_Historial where " & Campo & " Between #" & Format(Fecha1, "hh:mm:ss am/pm") & "# and #" & Format(Fecha2, "hh:mm:ss am/pm") & "# ORDER BY Fecha_Entrada desc, Hora_Entrada asc"
    Else
        SqlCH = "select * from Tbl_Historial where " & Campo & " " & Operador & " " & "'" & StrBusqueda & "' ORDER BY Fecha_Entrada desc, Hora_Entrada asc"
    End If
    
    ConexionHistorial
    RsCH.Open SqlCH, Conecta, adOpenDynamic, adLockBatchOptimistic
    
    FGR.AllowUserResizing = flexResizeBoth
    FGR.Rows = 1
    
    Titulos = Array("Usuario", "Maquina", "Fecha", "Hora entrada", "Hora Salida")
              
    FGR.Row = 0
    For i = 0 To RsCH.Fields.Count - 1
        FGR.Col = i
        FGR.ColAlignment(i) = flexAlignLeftCenter
        FGR.Text = Titulos(i)
        FGR.ColWidth(i) = CInt(TextWidth(FGR.Text) + 300)
        Ancho = Ancho + FGR.ColWidth(i)
    Next
    
    Do While Not RsCH.EOF
        FGR.Rows = FGR.Rows + 1
        FGR.Row = FGR.Rows - 1
        FGR.Col = 0
        Ancho = 0
        For i = 0 To RsCH.Fields.Count - 1
            FGR.Col = i
            FGR.Text = RsCH(i).Value & ""
            TamañoCol = FGR.ColWidth(i)
            If CInt(TextWidth(FGR.Text) + 100) > TamañoCol Then
                FGR.ColWidth(i) = CInt(TextWidth(FGR.Text) + 150)
            End If
            If FGR.Row / 2 <> Int(FGR.Row / 2) Then
                FGR.CellBackColor = RGB(194, 208, 252)
            End If
        Next
        RsCH.MoveNext
    Loop
    
    If Not RsCH.EOF Or Not RsCH.BOF Then
        RsCH.MoveFirst
        LlenarCamposReporte
        FGR.FixedRows = 1
    Else
        FGR.FixedRows = 0
        VaciarCamposReporte
    End If
    
    Opcion = 1
    Operador = ""
    Campo = ""
    StrBusqueda = ""
End Sub

Private Sub LlenarCamposReporte()
On Error Resume Next

    TxtHU.Text = RsCH!C_Acceso
    TxtHM.Text = RsCH!C_Maq
    TxtHF.Text = RsCH!Fecha_Entrada
    TxtHHE.Text = RsCH!Hora_Entrada
    TxtHHS.Text = RsCH!Hora_Salida
End Sub

Private Sub VaciarCamposReporte()
    TxtHU.Text = ""
    TxtHM.Text = ""
    TxtHF.Text = ""
    TxtHHE.Text = ""
    TxtHHS.Text = ""
End Sub

Private Sub CargarCR()
    CboCampo.Clear
    CboCampo.AddItem ("Maquina")
    CboCampo.AddItem ("Usuario")
    CboCampo.AddItem ("Fecha")
    CboCampo.AddItem ("Hora entrada")
    CboCampo.AddItem ("Hora salida")
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
        LblCampo.Caption = "Fecha_Entrada"
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
    ElseIf CboCampo.Text = "Hora salida" Then
        LblCampo.Caption = "Hora_Salida"
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
    ElseIf CboCampo.Text = "Hora entrada" Then
        LblCampo.Caption = "Hora_Entrada"
        AgregarItem 2
        DT2.Format = dtpTime
        DT2.Value = DateTime.Now
        DT1.Value = DateTime.Now
        DT2.Visible = True
        DT2.Visible = False
        Label30.Caption = "Hora 1:"
        Label31.Caption = "Hora 2:"
        Label30.Visible = True
        Label31.Visible = False
        Exit Sub
    End If
End Sub

Private Sub AgregarItem(Opcion As Integer)
    CboOperador.Clear
    If Opcion = 1 Then
        TxtBR.Visible = True
        CboOperador.AddItem ("=")
        CboOperador.AddItem ("<>")
        CboOperador.AddItem ("Like")
        CboOperador.Text = CboOperador.List(0)
        DT1.Visible = False
        DT2.Visible = False
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
        CboOperador.Text = CboOperador.List(0)
        DT1.Format = dtpTime: DT1.Visible = True
    End If
End Sub

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
    If LblCampo.Caption = "Hora_Entrada" Then
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
    If LblCampo.Caption = "Hora_Reporte" Then
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
    MDIPrincipal.AgregarVentana Me, "Historial", "Historial de usuarios..."
    Call CargarReportes(1)
    Call CargarCR
    InsertarPicture Me
    PosicionInicial Me
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
    TxtHU.Text = FGR.TextMatrix(FGR.MouseRow, 0)
    TxtHM.Text = FGR.TextMatrix(FGR.MouseRow, 1)
    TxtHF.Text = FGR.TextMatrix(FGR.MouseRow, 2)
    TxtHHE.Text = FGR.TextMatrix(FGR.MouseRow, 3)
    TxtHHS.Text = FGR.TextMatrix(FGR.MouseRow, 4)
End Sub

Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then
        Visible = False
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MDIPrincipal.RemoverVentana Me, "Historial"
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
