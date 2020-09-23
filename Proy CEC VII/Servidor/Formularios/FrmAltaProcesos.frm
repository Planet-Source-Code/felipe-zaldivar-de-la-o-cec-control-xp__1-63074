VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form FrmAltaProcesos 
   BackColor       =   &H00FFCEBB&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "::: Procesos Restringidos :::"
   ClientHeight    =   3135
   ClientLeft      =   4635
   ClientTop       =   4755
   ClientWidth     =   6255
   Icon            =   "FrmAltaProcesos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3135
   ScaleWidth      =   6255
   Begin VB.CommandButton CmdCancelar 
      BackColor       =   &H00FFCEBB&
      Caption         =   "&Cancelar"
      Height          =   255
      Left            =   5280
      TabIndex        =   3
      Top             =   2760
      Width           =   855
   End
   Begin VB.CommandButton CmdEliminar 
      BackColor       =   &H00FFCEBB&
      Caption         =   "&Eliminar"
      Height          =   255
      Left            =   4320
      TabIndex        =   2
      Top             =   2760
      Width           =   855
   End
   Begin VB.CommandButton CmdAgregar 
      BackColor       =   &H00FFCEBB&
      Caption         =   "&Agregar"
      Height          =   255
      Left            =   3360
      TabIndex        =   1
      Top             =   2760
      Width           =   855
   End
   Begin VB.TextBox TxtProceso 
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   120
      MaxLength       =   50
      TabIndex        =   0
      Top             =   2745
      Width           =   3135
   End
   Begin MSFlexGridLib.MSFlexGrid FG 
      Height          =   2535
      Left            =   120
      TabIndex        =   4
      Tag             =   "1"
      Top             =   120
      Width           =   6015
      _ExtentX        =   10610
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
End
Attribute VB_Name = "FrmAltaProcesos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim MensajeUsr1 As String
Dim RespuestaUsr1 As String

Private Sub CmdAgregar_Click()
    If TxtProceso.Text = "" Then MsgBox "Debes escribir el nombre del proceso...", , "Atención!!!": Exit Sub
    MensajeUsr1 = "Los datos son correctos?"
    If Preguntar(MensajeUsr1) = False Then Exit Sub
    
    ConexionPrincipal
    Sql = "select * from Tbl_Procesos where Programa= '" & TxtProceso.Text & "'"
    Rs.Open Sql, Conecta, adOpenDynamic, adLockBatchOptimistic
    
    If Rs.BOF = False And Rs.EOF = False Then
        MsgBox "Proceso ya registrado!!!", , "Atención!!!"
    Else
        Rs.Close
        Sql = "insert into Tbl_Procesos ([Programa]) VALUES ('" & TxtProceso.Text & "')"
        Rs.Open Sql, Conecta, adOpenDynamic, adLockBatchOptimistic
        LoadFG
    End If
End Sub

Private Sub LoadFG()
    Dim Ancho As Long
    Dim Titulos As Variant
    Dim i As Integer
    Set Rs = Nothing
    ConexionPrincipal
    Sql = "select * from Tbl_Procesos"
    Rs.Open Sql, Conecta, adOpenDynamic, adLockBatchOptimistic
    FG.AllowUserResizing = flexResizeBoth
    FG.Cols = Rs.Fields.Count
    FG.Rows = 1
    FG.Row = 0
    FG.ColAlignment(0) = flexAlignCenterCenter
    FG.ColWidth(0) = FG.Width
    FG.Text = "Proceso"
    Do While Not Rs.EOF
        FG.Rows = FG.Rows + 1
        FG.Row = FG.Rows - 1
        FG.Col = 0
        For i = 0 To Rs.Fields.Count - 1
            FG.Text = Rs(i).Value & ""
            If FG.Row / 2 <> Int(FG.Row / 2) Then
                FG.CellBackColor = RGB(194, 208, 252)
            End If
        Next
        Rs.MoveNext
    Loop
    If Not Rs.EOF Or Not Rs.BOF Then
        Rs.MoveFirst
        TxtProceso.Text = Rs!Programa
        FG.FixedRows = 1
    Else
        FG.FixedRows = 0
        TxtProceso.Text = ""
    End If
End Sub

Private Sub CmdCancelar_Click()
    Set Rs = Nothing
    Set FrmAltaProcesos = Nothing
    Unload Me
End Sub

Private Sub CmdEliminar_Click()
    MensajeUsr1 = "Estas seguro de eliminar el proceso: " & TxtProceso.Text
    If Preguntar(MensajeUsr1) = False Then Exit Sub
    ConexionPrincipal
    Sql = "select * from Tbl_Procesos where Programa= '" & TxtProceso.Text & "'"
    Rs.Open Sql, Conecta, adOpenDynamic, adLockBatchOptimistic
    
    If Rs.BOF = False And Rs.EOF = False Then
        Rs.Close
        Sql = "delete from Tbl_Procesos where Programa= '" & TxtProceso.Text & "'"
        Rs.Open Sql, Conecta, adOpenDynamic, adLockBatchOptimistic
        LoadFG
    Else
        MsgBox "Proceso no encontrado!!!", , "Atención!!!"
    End If
End Sub

Private Sub FG_Click()
    If FG.MouseRow = 0 Then Exit Sub
    TxtProceso.Text = FG.TextMatrix(FG.MouseRow, FG.MouseCol)
End Sub

Private Sub Form_Load()
    MDIPrincipal.AgregarVentana Me, "Procesos", "Procesos restringidos..."
    LoadFG
    PosicionInicial Me
End Sub

Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then
        Visible = False
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MDIPrincipal.RemoverVentana Me, "Procesos"
End Sub

Private Sub TxtProceso_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And TxtProceso.Text <> "" Then
        CmdAgregar.SetFocus
    End If
End Sub
