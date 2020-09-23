VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form FrmGrupos 
   BackColor       =   &H00FFCEBB&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "::: Grupos :::"
   ClientHeight    =   4710
   ClientLeft      =   5325
   ClientTop       =   3600
   ClientWidth     =   5055
   Icon            =   "FrmGrupos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4710
   ScaleWidth      =   5055
   Begin VB.CommandButton CmdAgregar 
      BackColor       =   &H00FFCEBB&
      Caption         =   "&Agregar"
      Height          =   255
      Left            =   1320
      TabIndex        =   4
      Top             =   4320
      Width           =   1095
   End
   Begin VB.CommandButton CmdEliminar 
      BackColor       =   &H00FFCEBB&
      Caption         =   "&Eliminar"
      Height          =   255
      Left            =   2580
      TabIndex        =   5
      Top             =   4320
      Width           =   1095
   End
   Begin VB.CommandButton CmdCancelar 
      BackColor       =   &H00FFCEBB&
      Caption         =   "&Cancelar"
      Height          =   255
      Left            =   3840
      TabIndex        =   6
      Top             =   4320
      Width           =   1095
   End
   Begin VB.TextBox TxtGrado 
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1320
      MaxLength       =   3
      TabIndex        =   3
      Top             =   3915
      Width           =   3615
   End
   Begin VB.TextBox TxtCarrera 
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1320
      MaxLength       =   50
      TabIndex        =   2
      Top             =   3555
      Width           =   3615
   End
   Begin VB.TextBox TxtGrupo 
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1320
      MaxLength       =   50
      TabIndex        =   1
      Top             =   3195
      Width           =   3615
   End
   Begin VB.TextBox TxtCGrupo 
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1320
      MaxLength       =   50
      TabIndex        =   0
      Top             =   2835
      Width           =   3615
   End
   Begin MSFlexGridLib.MSFlexGrid FG 
      Height          =   2535
      Left            =   120
      TabIndex        =   7
      Tag             =   "1"
      Top             =   120
      Width           =   4815
      _ExtentX        =   8493
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
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Grado:"
      Height          =   195
      Left            =   120
      TabIndex        =   11
      Top             =   3960
      Width           =   480
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Carrera:"
      Height          =   195
      Left            =   120
      TabIndex        =   10
      Top             =   3600
      Width           =   555
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Grupo:"
      Height          =   195
      Left            =   120
      TabIndex        =   9
      Top             =   3240
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Clave de grupo:"
      Height          =   195
      Left            =   120
      TabIndex        =   8
      Top             =   2880
      Width           =   1125
   End
End
Attribute VB_Name = "FrmGrupos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim MensajeUsr1 As String
Dim RespuestaUsr1 As String

Private Sub CmdAgregar_Click()
    If TxtCGrupo.Text = "" Or TxtGrupo.Text = "" Or TxtCarrera.Text = "" Or TxtGrado.Text = "" Then MsgBox _
    "No debes dejar campos vacios...", , "Atención!!!": Exit Sub
    
    MensajeUsr1 = "Los datos son correctos?"
    If Preguntar(MensajeUsr1) = False Then Exit Sub
    
    ConexionPrincipal
    Sql = "select * from Tbl_Grupo where C_Grupo= '" & TxtCGrupo.Text & "'"
    Rs.Open Sql, Conecta, adOpenDynamic, adLockBatchOptimistic
    
    If Rs.BOF = False And Rs.EOF = False Then
        MsgBox "Proceso ya registrado!!!", , "Atención!!!"
    Else
        Rs.Close
        Sql = "insert into Tbl_Grupo ([C_Grupo],[Grupo],[Carrera],[Grado]) VALUES ('" & TxtCGrupo.Text & "','" & TxtGrupo.Text & "','" & TxtCarrera.Text & "','" & Val(TxtGrado.Text) & "')"
        Rs.Open Sql, Conecta, adOpenDynamic, adLockBatchOptimistic
        LoadFG
    End If
End Sub

Private Sub LoadFG()

    Dim Titulos As Variant
    Dim i As Integer
    Set Rs = Nothing
    
    ConexionPrincipal
    Sql = "select * from Tbl_Grupo ORDER BY C_Grupo"
    Rs.Open Sql, Conecta, adOpenDynamic, adLockBatchOptimistic
    
    
    FG.AllowUserResizing = flexResizeBoth
    FG.Cols = Rs.Fields.Count
    FG.Rows = 1
    FG.Row = 0
    
    
    Titulos = Array("Clave de Grupo", "Carrera", "Grupo", "Grado")
    
    For i = 0 To Rs.Fields.Count - 1
        FG.Col = i
        FG.ColAlignment(i) = flexAlignCenterCenter
        FG.Text = Titulos(i)
        FG.ColWidth(i) = CInt(TextWidth(FG.Text) + 100)
    Next
    
    Dim TamañoCol As Long
    
    Do While Not Rs.EOF
        FG.Rows = FG.Rows + 1
        FG.Row = FG.Rows - 1
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
    
    If Not Rs.EOF Or Not Rs.BOF Then
        FG.FixedRows = 1
        Rs.MoveFirst
        TxtCGrupo.Text = Rs!C_Grupo
        TxtGrupo.Text = Rs!Grupo
        TxtCarrera.Text = Rs!Carrera
        TxtGrado.Text = Rs!Grado
    Else
        FG.FixedRows = 0
        TxtCGrupo.Text = ""
        TxtGrupo.Text = ""
        TxtCarrera.Text = ""
        TxtGrado.Text = ""
    End If
    
End Sub

Private Sub CmdCancelar_Click()
    Unload Me
End Sub

Private Sub CmdEliminar_Click()
    MensajeUsr1 = "Estas seguro de eliminar el grupo: " & TxtCGrupo.Text
    If Preguntar(MensajeUsr1) = False Then Exit Sub

    ConexionPrincipal
    Sql = "select * from Tbl_Grupo where C_Grupo= '" & TxtCGrupo.Text & "'"
    Rs.Open Sql, Conecta, adOpenDynamic, adLockBatchOptimistic
    
    If Rs.BOF = False And Rs.EOF = False Then
        Rs.Close
        Sql = "delete from Tbl_Grupo where C_Grupo= '" & TxtCGrupo.Text & "'"
        Rs.Open Sql, Conecta, adOpenDynamic, adLockBatchOptimistic
        LoadFG
    Else
        MsgBox "Grupo no encontrado!!!", , "Atención!!!"
    End If
End Sub

Private Sub FG_Click()
    If FG.MouseRow = 0 Then Exit Sub
    TxtCGrupo.Text = FG.TextMatrix(FG.MouseRow, 0)
End Sub

Private Sub Form_Activate()
    LoadFG
End Sub

Private Sub Form_Load()
    MDIPrincipal.AgregarVentana Me, "Grupos", "Grupo(s), grado(s), carrera(s)..."
    InsertarPicture Me
    PosicionInicial Me
End Sub

Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then
        Visible = False
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MDIPrincipal.RemoverVentana Me, "Grupos"
End Sub

Private Sub TxtCarrera_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And TxtCarrera.Text <> "" Then
        TxtGrado.SetFocus
    End If
End Sub

Private Sub TxtCGrupo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And TxtCGrupo.Text <> "" Then
        TxtGrupo.SetFocus
    End If
End Sub

Private Sub TxtGrado_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And TxtGrado.Text <> "" Then
        CmdAgregar.SetFocus
    End If
End Sub

Private Sub TxtGrupo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And TxtGrupo.Text <> "" Then
        TxtCarrera.SetFocus
    End If
End Sub
