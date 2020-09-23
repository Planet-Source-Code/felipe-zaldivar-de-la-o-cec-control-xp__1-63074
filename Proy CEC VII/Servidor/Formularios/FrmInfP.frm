VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FrmInfP 
   BackColor       =   &H00FFCEBB&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "::: Información Personal :::"
   ClientHeight    =   7305
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10770
   Icon            =   "FrmInfP.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7305
   ScaleWidth      =   10770
   Begin VB.PictureBox Picture5 
      BackColor       =   &H00E89C78&
      BorderStyle     =   0  'None
      Height          =   4575
      Left            =   6360
      ScaleHeight     =   4575
      ScaleWidth      =   4335
      TabIndex        =   41
      Top             =   2640
      Width           =   4335
      Begin VB.PictureBox Picture7 
         BackColor       =   &H00D05C28&
         BorderStyle     =   0  'None
         Height          =   2055
         Left            =   120
         ScaleHeight     =   2055
         ScaleWidth      =   4095
         TabIndex        =   43
         Top             =   2400
         Width           =   4095
         Begin VB.TextBox TxtProcesos 
            Height          =   1815
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   22
            Top             =   120
            Width           =   3855
         End
      End
      Begin VB.PictureBox Picture6 
         BackColor       =   &H00D05C28&
         BorderStyle     =   0  'None
         Height          =   2055
         Left            =   120
         ScaleHeight     =   2055
         ScaleWidth      =   4095
         TabIndex        =   42
         Top             =   120
         Width           =   4095
         Begin VB.TextBox TxtHistorial 
            Height          =   1815
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   20
            Top             =   120
            Width           =   3855
         End
      End
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00E89C78&
      BorderStyle     =   0  'None
      Height          =   2415
      Left            =   6360
      ScaleHeight     =   2415
      ScaleWidth      =   4335
      TabIndex        =   37
      Top             =   120
      Width           =   4335
      Begin VB.ComboBox CboTablas 
         Height          =   315
         Left            =   715
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   180
         Width           =   1095
      End
      Begin VB.ComboBox CboCampo 
         Height          =   315
         Left            =   2640
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   180
         Width           =   1575
      End
      Begin VB.ComboBox CboOperador 
         Height          =   315
         Left            =   960
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox TxtBR 
         Height          =   285
         Left            =   120
         TabIndex        =   15
         Top             =   1320
         Width           =   4095
      End
      Begin VB.CommandButton CmdTodos 
         BackColor       =   &H00E89C78&
         Caption         =   "&Mostrar Todos"
         Height          =   375
         Left            =   3000
         TabIndex        =   18
         Top             =   1920
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker DT2 
         CausesValidation=   0   'False
         Height          =   285
         Left            =   3000
         TabIndex        =   17
         Top             =   1320
         Visible         =   0   'False
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         Format          =   47710211
         CurrentDate     =   38601.9583333333
      End
      Begin MSComCtl2.DTPicker DT1 
         CausesValidation=   0   'False
         Height          =   285
         Left            =   885
         TabIndex        =   16
         Top             =   1320
         Visible         =   0   'False
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         Format          =   47710211
         CurrentDate     =   38601
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Operador:"
         Height          =   195
         Left            =   120
         TabIndex        =   48
         Top             =   780
         Width           =   705
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Campo:"
         Height          =   195
         Left            =   1920
         TabIndex        =   47
         Top             =   240
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tabla:"
         Height          =   195
         Left            =   120
         TabIndex        =   46
         Top             =   240
         Width           =   450
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha 2:"
         Height          =   195
         Left            =   2235
         TabIndex        =   45
         Top             =   1365
         Visible         =   0   'False
         Width           =   630
      End
      Begin VB.Label LblTabla 
         BackColor       =   &H0080FFFF&
         Caption         =   "LblTabla"
         Height          =   255
         Left            =   960
         TabIndex        =   44
         Top             =   720
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha 1:"
         Height          =   195
         Left            =   120
         TabIndex        =   40
         Top             =   1365
         Visible         =   0   'False
         Width           =   630
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " "
         Height          =   195
         Left            =   120
         TabIndex        =   39
         Top             =   1320
         Visible         =   0   'False
         Width           =   45
      End
      Begin VB.Label LblCampo 
         BackColor       =   &H0080FFFF&
         Caption         =   "LblCampo"
         Height          =   255
         Left            =   2040
         TabIndex        =   38
         Top             =   720
         Visible         =   0   'False
         Width           =   855
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00E89C78&
      BorderStyle     =   0  'None
      Height          =   2415
      Left            =   120
      ScaleHeight     =   2415
      ScaleWidth      =   6135
      TabIndex        =   24
      Top             =   120
      Width           =   6135
      Begin VB.TextBox TxtCuenta 
         Height          =   285
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   120
         Width           =   1575
      End
      Begin VB.TextBox TxtNombre 
         Height          =   285
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   480
         Width           =   1575
      End
      Begin VB.TextBox TxtNivel 
         Height          =   285
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   840
         Width           =   1575
      End
      Begin VB.TextBox TxtAmo 
         Height          =   285
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   1200
         Width           =   1575
      End
      Begin VB.TextBox TxtBloq 
         Height          =   285
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   1560
         Width           =   1575
      End
      Begin VB.TextBox TxtNBloq 
         Height          =   285
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   1920
         Width           =   1575
      End
      Begin VB.TextBox TxtFecha 
         Height          =   285
         Left            =   4440
         Locked          =   -1  'True
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   120
         Width           =   1575
      End
      Begin VB.TextBox TxtResp 
         Height          =   285
         Left            =   4440
         Locked          =   -1  'True
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   480
         Width           =   1575
      End
      Begin VB.TextBox TxtCGrupo 
         Height          =   285
         Left            =   4440
         Locked          =   -1  'True
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   840
         Width           =   1575
      End
      Begin VB.TextBox TxtGrupo 
         Height          =   285
         Left            =   4440
         Locked          =   -1  'True
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   1200
         Width           =   1575
      End
      Begin VB.TextBox TxtCarrera 
         Height          =   285
         Left            =   4440
         Locked          =   -1  'True
         TabIndex        =   10
         Text            =   "Text1"
         Top             =   1560
         Width           =   1575
      End
      Begin VB.TextBox TxtGrado 
         Height          =   285
         Left            =   4440
         Locked          =   -1  'True
         TabIndex        =   11
         Text            =   "Text1"
         Top             =   1920
         Width           =   1575
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cuenta:"
         Height          =   195
         Left            =   120
         TabIndex        =   36
         Top             =   165
         Width           =   555
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre:"
         Height          =   195
         Left            =   120
         TabIndex        =   35
         Top             =   525
         Width           =   600
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nivel:"
         Height          =   195
         Left            =   120
         TabIndex        =   34
         Top             =   885
         Width           =   405
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Amonestaciones:"
         Height          =   195
         Left            =   120
         TabIndex        =   33
         Top             =   1245
         Width           =   1215
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bloqueado."
         Height          =   195
         Left            =   120
         TabIndex        =   32
         Top             =   1605
         Width           =   810
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bloqueos:"
         Height          =   195
         Left            =   120
         TabIndex        =   31
         Top             =   1965
         Width           =   705
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha de registro:"
         Height          =   195
         Left            =   3120
         TabIndex        =   30
         Top             =   165
         Width           =   1275
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Responsable:"
         Height          =   195
         Left            =   3120
         TabIndex        =   29
         Top             =   525
         Width           =   975
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Clave de grupo:"
         Height          =   195
         Left            =   3120
         TabIndex        =   28
         Top             =   885
         Width           =   1125
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Grupo:"
         Height          =   195
         Left            =   3120
         TabIndex        =   27
         Top             =   1245
         Width           =   480
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Carrera"
         Height          =   195
         Left            =   3120
         TabIndex        =   26
         Top             =   1605
         Width           =   510
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Grado"
         Height          =   195
         Left            =   3120
         TabIndex        =   25
         Top             =   1965
         Width           =   435
      End
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00E89C78&
      BorderStyle     =   0  'None
      Height          =   4575
      Left            =   120
      ScaleHeight     =   4575
      ScaleWidth      =   6135
      TabIndex        =   23
      Top             =   2640
      Width           =   6135
      Begin MSFlexGridLib.MSFlexGrid FGP 
         Height          =   1815
         Left            =   120
         TabIndex        =   21
         Tag             =   "1"
         Top             =   2640
         Width           =   5880
         _ExtentX        =   10372
         _ExtentY        =   3201
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
      Begin MSFlexGridLib.MSFlexGrid FGR 
         Height          =   1935
         Left            =   120
         TabIndex        =   19
         Tag             =   "1"
         Top             =   360
         Width           =   5880
         _ExtentX        =   10372
         _ExtentY        =   3413
         _Version        =   393216
         Rows            =   1
         Cols            =   6
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
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "Procesos"
         Height          =   255
         Left            =   120
         TabIndex        =   50
         Top             =   2350
         Width           =   855
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "Reportes"
         Height          =   255
         Left            =   120
         TabIndex        =   49
         Top             =   50
         Width           =   1215
      End
   End
End
Attribute VB_Name = "FrmInfP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public CuentaInfP As String
Public MaquinaInfP As String
Dim TitulosI() As Variant
Dim TitulosII() As Variant

Private Sub CboCampo_Click()
    If LblTabla.Caption = "Tbl_Procesos_Reg" Then
        Select Case CboCampo.Text
            Case Is = "Maquina"
                LblCampo.Caption = "C_Maq"
                AgregarItemII 1
            Case Is = "Fecha"
                LblCampo.Caption = "Fecha"
                AgregarItemII 2
                HoraO 1
            Case Is = "Hora"
                LblCampo.Caption = "Hora"
                AgregarItemII 2
                HoraO 2
            Case Is = "Proceso"
                LblCampo.Caption = "Proceso"
                AgregarItemII 1
        End Select
    Else
        Select Case CboCampo.Text
            Case Is = "Maquina"
                LblCampo.Caption = "C_Maq"
                AgregarItemII 1
            Case Is = "Fecha reporte"
                LblCampo.Caption = "Fecha_Reporte"
                AgregarItemII 2
                HoraO 1
            Case Is = "Hora reporte"
                LblCampo.Caption = "Hora_Reporte"
                AgregarItemII 2
                HoraO 2
            Case Is = "Titulo"
                LblCampo.Caption = "Titulo"
                AgregarItemII 1
            Case "Reporte"
                LblCampo.Caption = "Reporte"
                AgregarItemII 1
        End Select
    End If
End Sub

Private Sub HoraO(OptHO As Integer)
    If OptHO = 1 Then
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
    Else
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

Private Sub CboTablas_Click()
    If CboTablas.Text = "Procesos" Then
        LblTabla.Caption = "Tbl_Procesos_Reg"
        AgregarItem 1
    Else
        LblTabla.Caption = "Tbl_Reportes"
        AgregarItem 2
    End If
End Sub

Private Sub AgregarItemII(Opcion As Integer)
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

Private Sub AgregarItem(OptItem As Integer)
    If OptItem = 1 Then
        CboCampo.Clear
        CboCampo.AddItem "Fecha"
        CboCampo.AddItem "Hora"
        CboCampo.AddItem "Proceso"
    Else
        CboCampo.Clear
        CboCampo.AddItem "Fecha reporte"
        CboCampo.AddItem "Hora reporte"
        CboCampo.AddItem "Titulo"
        CboCampo.AddItem "Reporte"
    End If
    CboCampo.AddItem "Maquina"
    CboCampo.Text = CboCampo.List(0)
End Sub

Private Sub CmdTodos_Click()
    FGInicial
End Sub

Private Sub DT1_Click()
   Call DTEvent
End Sub

Private Sub DTEvent()
    If LblTabla.Caption = "Tbl_Procesos_Reg" Then
        If LblCampo.Caption = "Hora" Then
            If Not IsDate(DT1.Value) Or Not IsDate(DT1.Value) Then MsgBox "Hora incorrecta!!!": Exit Sub
            If CboOperador.Text = "Between" Then
                Call CargarOpcionesInf("Tbl_Procesos_Reg", 7, TitulosII, FGP, 4, "Fecha" _
                        , "Hora", LblCampo.Caption, CboOperador.Text, , DT1.Value, DT2.Value)
            Else
                Call CargarOpcionesInf("Tbl_Procesos_Reg", 6, TitulosII, FGP, 4, "Fecha" _
                        , "Hora", LblCampo.Caption, CboOperador.Text, , DT1.Value)
            End If
            Exit Sub
        End If
    
        If Not IsDate(DT2.Value) Or Not IsDate(DT1.Value) Then MsgBox "Fecha incorrecta!!!": Exit Sub
        If CboOperador.Text = "Between" Then
            Call CargarOpcionesInf("Tbl_Procesos_Reg", 5, TitulosII, FGP, 4, "Fecha" _
                        , "Hora", LblCampo.Caption, CboOperador.Text, , DT1.Value, DT2.Value)
        Else
            Call CargarOpcionesInf("Tbl_Procesos_Reg", 4, TitulosII, FGP, 4, "Fecha" _
                            , "Hora", LblCampo.Caption, CboOperador.Text, , DT1.Value)
        End If
    Else
        If LblCampo.Caption = "Hora_Reporte" Then
            If Not IsDate(DT1.Value) Or Not IsDate(DT1.Value) Then MsgBox "Hora incorrecta!!!": Exit Sub
            If CboOperador.Text = "Between" Then
                Call CargarOpcionesInf("Tbl_Reportes", 7, TitulosI, FGR, 5, "Fecha" _
                        , "Hora", LblCampo.Caption, CboOperador.Text, , DT1.Value, DT2.Value)
            Else
                Call CargarOpcionesInf("Tbl_Reportes", 6, TitulosI, FGP, 5, "Fecha" _
                        , "Hora", LblCampo.Caption, CboOperador.Text, , DT1.Value)
            End If
            Exit Sub
        End If
    
        If Not IsDate(DT1.Value) Or Not Not IsDate(DT2.Value) = False Then MsgBox "Fecha incorrecta!!!": Exit Sub
        If CboOperador.Text = "Between" Then
            Call CargarOpcionesInf("Tbl_Reportes", 5, TitulosI, FGR, 5, "Fecha_Reporte" _
                        , "Hora_Reporte", LblCampo.Caption, CboOperador.Text, , DT1.Value, DT2.Value)
        Else
            Call CargarOpcionesInf("Tbl_Reportes", 4, TitulosI, FGR, 5, "Fecha_Reporte" _
                            , "Hora_Reporte", LblCampo.Caption, CboOperador.Text, , DT1.Value)
        End If
    End If
End Sub

Private Sub DT2_Click()
    Call DTEvent
End Sub

Private Sub FGR_Click()
    If FGR.MouseRow = 0 Then Exit Sub
    TxtHistorial.Text = ""
    TxtHistorial.Text = "Usuario: " & FGR.TextMatrix(FGR.MouseRow, 0) & vbNewLine
    TxtHistorial.Text = TxtHistorial.Text & "Maquina: " & FGR.TextMatrix(FGR.MouseRow, 1) & vbNewLine
    TxtHistorial.Text = TxtHistorial.Text & "Fecha: " & FGR.TextMatrix(FGR.MouseRow, 2) & vbNewLine
    TxtHistorial.Text = TxtHistorial.Text & "Hora: " & FGR.TextMatrix(FGR.MouseRow, 3) & vbNewLine
    TxtHistorial.Text = TxtHistorial.Text & "Titulo: " & FGR.TextMatrix(FGR.MouseRow, 4) & vbNewLine
    TxtHistorial.Text = TxtHistorial.Text & "Reporte: " & FGR.TextMatrix(FGR.MouseRow, 5) & vbNewLine
End Sub

Private Sub FGP_Click()
    If FGP.MouseRow = 0 Then Exit Sub
    TxtProcesos.Text = ""
    TxtProcesos.Text = "Usuario: " & FGP.TextMatrix(FGP.MouseRow, 0) & vbNewLine
    TxtProcesos.Text = TxtProcesos.Text & "Maquina: " & FGP.TextMatrix(FGP.MouseRow, 1) & vbNewLine
    TxtProcesos.Text = TxtProcesos.Text & "Fecha: " & FGP.TextMatrix(FGP.MouseRow, 2) & vbNewLine
    TxtProcesos.Text = TxtProcesos.Text & "Hora " & FGP.TextMatrix(FGP.MouseRow, 3) & vbNewLine
    TxtProcesos.Text = TxtProcesos.Text & "Proceso: " & FGP.TextMatrix(FGP.MouseRow, 4)
End Sub

Private Sub Form_Load()
    MDIPrincipal.AgregarVentana Me, "Inf. Personal", "Información personal de los usuarios..."
    VaciarControles
    ConexionConsultasIP
    SqlCCIP = "Select * from Tbl_Acceso where C_Acceso='" & CuentaInfP & "'"
    RsCCIP.Open SqlCCIP, Conecta, adOpenDynamic, adLockBatchOptimistic
    If Not RsCCIP.EOF Then
        LlenarControles 1
    End If
    RsCCIP.Close
    SqlCCIP = "Select * from Tbl_Grupo where C_Grupo='" & TxtCGrupo.Text & "'"
    RsCCIP.Open SqlCCIP, Conecta, adOpenDynamic, adLockBatchOptimistic
    If Not RsCCIP.EOF Then
        LlenarControles 2
    End If
    TitulosI = Array("Usuario", "Maquina", "Fecha", "Hora", "Titulo", "Reporte")
    TitulosII = Array("Usuario", "Maquina", "Fecha", "Hora", "Proceso")
    FGInicial
    Call CargarCT
    PosicionInicial Me
End Sub

Public Sub CargarOpcionesInf(Tabla As String, Opcion As Integer _
           , Titulos() As Variant, FG As MSFlexGrid, NCols As Integer _
           , FO As String, HO As String, Campo As String _
           , Operador As String, Optional StrBusqueda As String _
           , Optional Fecha1 As Date, Optional Fecha2 As Date)
           
    Dim Ancho, TamañoCol As Long
    Dim i As Integer
    
    Ancho = 0
    TamañoCol = 0
    
    If Opcion = 1 Then
        SqlCCIP = "select * from " & Tabla & " where C_Acceso='" & TxtCuenta.Text & "' Order by " & FO & " desc, " & HO & " desc"
    ElseIf Opcion = 2 Then
        SqlCCIP = "select * from " & Tabla & " where C_Acceso='" & TxtCuenta.Text & "' and " & Campo & " " & Operador & " " & "'" & StrBusqueda & "' Order by " & FO & " desc, " & HO & " desc"
    ElseIf Opcion = 3 Then
        SqlCCIP = "select * from " & Tabla & " where C_Acceso='" & TxtCuenta.Text & "' and " & Campo & " " & Operador & " " & "'%" & StrBusqueda & "%' Order by " & FO & " desc, " & HO & " desc"
    ElseIf Opcion = 4 Then
        SqlCCIP = "select * from " & Tabla & " where  C_Acceso='" & TxtCuenta.Text & "' and " & Campo & " " & Operador & " " & "#" & Format(Fecha1, "mm/dd/yy") & "# Order by " & FO & " desc, " & HO & " desc"
    ElseIf Opcion = 5 Then
        SqlCCIP = "select * from " & Tabla & " where C_Acceso='" & TxtCuenta.Text & "' and " & Campo & " Between #" & Format(Fecha1, "mm/dd/yy") & "# and #" & Format(Fecha2, "mm/dd/yy") & "# Order by " & FO & " desc, " & HO & " desc"
     ElseIf Opcion = 6 Then
        SqlCCIP = "select * from " & Tabla & " where C_Acceso='" & TxtCuenta.Text & "' and " & Campo & " " & Operador & " " & "#" & Format(Fecha1, "hh:mm:ss am/pm") & "# Order by " & FO & " desc, " & HO & " desc"
    ElseIf Opcion = 7 Then
        SqlCCIP = "select * from " & Tabla & " where C_Acceso='" & TxtCuenta.Text & "' and " & Campo & " Between #" & Format(Fecha1, "hh:mm:ss am/pm") & "# and #" & Format(Fecha2, "hh:mm:ss am/pm") & "# Order by " & FO & " desc, " & HO & " desc"
    End If
    'MsgBox SqlCCIP
    ConexionConsultasIP
    RsCCIP.Open SqlCCIP, Conecta, adOpenDynamic, adLockBatchOptimistic
    FG.AllowUserResizing = flexResizeBoth
    FG.Rows = 1
    FG.Row = 0
    FG.Cols = NCols + 1

    For i = 0 To NCols
        FG.Col = i
        FG.ColAlignment(i) = flexAlignLeftCenter
        FG.Text = Titulos(i)
        FG.ColWidth(i) = CInt(TextWidth(FG.Text) + 300)
        Ancho = Ancho + FG.ColWidth(i)
    Next
    
    Do While Not RsCCIP.EOF
        FG.Rows = FG.Rows + 1
        FG.Row = FG.Rows - 1
        FG.Col = 0
        Ancho = 0
        For i = 0 To RsCCIP.Fields.Count - 1
            FG.Col = i
            FG.Text = RsCCIP(i).Value & ""
            TamañoCol = FG.ColWidth(i)
            If CInt(TextWidth(FG.Text) + 100) > TamañoCol Then
                FG.ColWidth(i) = CInt(TextWidth(FG.Text) + 100)
            End If
            If FG.Row / 2 <> Int(FG.Row / 2) Then
                FG.CellBackColor = RGB(194, 208, 252)
            End If
        Next
        RsCCIP.MoveNext
    Loop
    
    If Not RsCCIP.EOF Or Not RsCCIP.BOF Then
        RsCCIP.MoveFirst
        'LlenarCamposReporte
        FG.FixedRows = 1
    Else
        FG.FixedRows = 0
        'VaciarCamposReporte
    End If
End Sub

Private Sub LlenarControles(OptRS As Integer)
    If OptRS = 1 Then
        TxtCGrupo.Text = RsCCIP!C_Grupo
        TxtResp.Text = RsCCIP!C_U_Registro
        TxtFecha.Text = RsCCIP!Fecha_Reg
        TxtNBloq.Text = RsCCIP!N_Bloqueos
        TxtBloq.Text = RsCCIP!Usr_Bloqueado
        TxtAmo.Text = RsCCIP!Amonestaciones
        TxtNivel.Text = RsCCIP!Nivel
        TxtNombre.Text = RsCCIP!Nombre
        TxtCuenta.Text = RsCCIP!C_Acceso
    Else
        TxtGrado.Text = RsCCIP!Grado
        TxtCarrera.Text = RsCCIP!Carrera
        TxtGrupo.Text = RsCCIP!Grupo
    End If
    If TxtNivel.Text = "1" Then
        TxtGrado.Visible = True
        TxtCarrera.Visible = True
        TxtGrupo.Visible = True
        Label13.Visible = True
        Label14.Visible = True
        Label15.Visible = True
    Else
        TxtGrado.Visible = False
        TxtCarrera.Visible = False
        TxtGrupo.Visible = False
        Label13.Visible = False
        Label14.Visible = False
        Label15.Visible = False
    End If
End Sub

Private Sub VaciarControles()
    TxtGrado.Text = ""
    TxtCarrera.Text = ""
    TxtGrupo.Text = ""
    TxtCGrupo.Text = ""
    TxtResp.Text = ""
    TxtFecha.Text = ""
    TxtNBloq.Text = ""
    TxtBloq.Text = ""
    TxtAmo.Text = ""
    TxtNivel.Text = ""
    TxtNombre.Text = ""
    TxtCuenta.Text = ""
End Sub

Private Sub CargarCT()
    CboTablas.Clear
    CboTablas.AddItem "Procesos"
    CboTablas.AddItem "Reportes"
    CboTablas.Text = CboTablas.List(0)
End Sub

Private Sub FGInicial()
    CargarOpcionesInf "Tbl_Reportes", 1, TitulosI, FGR, 5, "Fecha_Reporte" _
                    , "Hora_Reporte", "C_Acceso", "=", TxtCuenta.Text
    CargarOpcionesInf "Tbl_Procesos_Reg", 1, TitulosII, FGP, 4, "Fecha" _
                    , "Hora", "C_Acceso", "=", TxtCuenta.Text
End Sub


Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then
        Visible = False
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set FrmInfP = Nothing
    MDIPrincipal.RemoverVentana Me, "Inf. Personal"
End Sub

Private Sub TxtBR_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If TxtBR.Text = "" Then FGInicial: Exit Sub
        If LblTabla.Caption = "Tbl_Procesos_Reg" Then
            If CboOperador.Text = "Like" Then
                Call CargarOpcionesInf("Tbl_Procesos_Reg", 3, TitulosII, FGP, 4, "Fecha" _
                        , "Hora", LblCampo.Caption, CboOperador.Text, TxtBR.Text)
            Else
                Call CargarOpcionesInf("Tbl_Procesos_Reg", 2, TitulosII, FGP, 4, "Fecha" _
                        , "Hora", LblCampo.Caption, CboOperador.Text, TxtBR.Text)
            End If
        Else
            If CboOperador.Text = "Like" Then
                Call CargarOpcionesInf("Tbl_Reportes", 3, TitulosI, FGR, 5, "Fecha_Reporte" _
                    , "Hora_Reporte", LblCampo.Caption, CboOperador.Text, TxtBR.Text)
            Else
                Call CargarOpcionesInf("Tbl_Reportes", 2, TitulosI, FGR, 5, "Fecha_Reporte" _
                        , "Hora_Reporte", LblCampo.Caption, CboOperador.Text, TxtBR.Text)
            End If
        End If
    End If
End Sub
