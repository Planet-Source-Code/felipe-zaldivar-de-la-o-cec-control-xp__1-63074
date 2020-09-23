VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmProcesos 
   BackColor       =   &H00FFCEBB&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "::: Procesos Registrados :::"
   ClientHeight    =   5895
   ClientLeft      =   4635
   ClientTop       =   3135
   ClientWidth     =   6375
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5895
   ScaleWidth      =   6375
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFCEBB&
      Caption         =   "Todos"
      Height          =   255
      Left            =   5520
      TabIndex        =   8
      Top             =   2850
      Width           =   735
   End
   Begin MSComctlLib.ProgressBar PB 
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   5520
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.TextBox TxtC_Acceso 
      Height          =   285
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   2835
      Width           =   4455
   End
   Begin VB.TextBox TxtFecha 
      Height          =   285
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   3195
      Width           =   5295
   End
   Begin VB.TextBox TxtProceso 
      Height          =   1845
      Left            =   960
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   3555
      Width           =   5295
   End
   Begin MSFlexGridLib.MSFlexGrid FG 
      Height          =   2535
      Left            =   120
      TabIndex        =   3
      Tag             =   "1"
      Top             =   120
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   4471
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
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cuenta:"
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   2880
      Width           =   555
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha:"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   3240
      Width           =   495
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Procesos:"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   3600
      Width           =   705
   End
End
Attribute VB_Name = "FrmProcesos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub LoadFG(Opcion As Integer, Optional Clave As String, Optional Fecha As Date)
    Dim Titulos As Variant
    Dim i As Integer
    Set Rs = Nothing
    Me.Enabled = False
    
    TxtC_Acceso.Text = ""
    TxtFecha.Text = ""
    TxtProceso.Text = ""
    
    Conexion
    
    If Opcion = 1 Then
        Sql = "select * from Tbl_RPP"
    Else
        Sql = "select * from Tbl_RPP where C_Acceso = '" & Clave & "' and fecha=#" & Format(Fecha, "mm/dd/yy") & "#"
    End If
    
    Rs.Open Sql, Conecta, adOpenDynamic, adLockBatchOptimistic
    
    FG.AllowUserResizing = flexResizeBoth
    FG.Cols = Rs.Fields.Count
    FG.Rows = 1
    FG.Row = 0
    
    Titulos = Array("Cuenta", "Fecha", "Procesos")
    
    For i = 0 To Rs.Fields.Count - 1
        FG.Col = i
        FG.ColAlignment(i) = flexAlignCenterCenter
        FG.Text = Titulos(i)
        FG.ColWidth(i) = CInt(TextWidth(FG.Text) + 100)
    Next
    
    PB.Min = 0
    PB.Value = 0
    If Not Rs.EOF Then
        Rs.MoveLast
        Rs.MoveFirst
    End If
    If Rs.RecordCount > 0 Then
        PB.Max = Rs.RecordCount
    Else
        PB.Max = 1
    End If
    
    Dim TamañoCol As Long
    
    Do While Not Rs.EOF
        
        FG.Rows = FG.Rows + 1
        FG.Row = FG.Rows - 1
        For i = 0 To Rs.Fields.Count - 1
            FG.Col = i
            FG.Text = Rs(i).Value & ""
            TamañoCol = FG.ColWidth(i)
            If (TextWidth(FG.Text) + 100) > TamañoCol And (TextWidth(FG.Text) + 100) < (TamañoCol * 5) Then
                FG.ColWidth(i) = CInt(TextWidth(FG.Text) + 100)
            End If
            If FG.Row / 2 <> Int(FG.Row / 2) Then
                FG.CellBackColor = RGB(194, 208, 252)
            End If
        Next
        If PB.Value < PB.Max Then PB.Value = PB.Value + 1: PB.Refresh
        Rs.MoveNext
        
    Loop
    
    If Not Rs.BOF Then
        Dim CadenaP As String
        Dim PosP As Integer
        CadenaP = ""
        FG.FixedRows = 1
        Rs.MoveFirst
        TxtC_Acceso.Text = Rs!C_Acceso
        TxtFecha.Text = Rs!Fecha
        CadenaP = Rs!Procesos
        Do While InStr(1, CadenaP, "+++++")
            PosP = InStr(1, CadenaP, "+++++")
            TxtProceso.Text = TxtProceso.Text & Mid(CadenaP, 1, PosP - 1) & vbNewLine
            CadenaP = Mid(CadenaP, PosP + 5)
        Loop
    Else
        FG.FixedRows = 0
        TxtC_Acceso.Text = ""
        TxtFecha.Text = ""
        TxtProceso.Text = ""
    End If
    Me.Enabled = True
End Sub

Private Sub CmdCancelar_Click()
    Set Rs = Nothing
    Set FrmProcesos = Nothing
    Unload Me
End Sub


Private Sub Command1_Click()
    LoadFG 1
End Sub

Private Sub FG_Click()
    If FG.MouseRow = 0 Then Exit Sub
    TxtC_Acceso.Text = FG.TextMatrix(FG.MouseRow, 0)
    TxtFecha.Text = FG.TextMatrix(FG.MouseRow, 1)
    LoadFG 2, TxtC_Acceso.Text, CDate(TxtFecha.Text)
End Sub

Private Sub Form_Load()
    LoadFG 1
End Sub

