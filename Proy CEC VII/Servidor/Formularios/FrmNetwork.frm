VERSION 5.00
Begin VB.Form FrmNetwork 
   BackColor       =   &H00FFCEBB&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "::: CEC CONTROL :::"
   ClientHeight    =   6015
   ClientLeft      =   3030
   ClientTop       =   2910
   ClientWidth     =   5655
   Icon            =   "FrmNetwork.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6015
   ScaleWidth      =   5655
   Begin VB.PictureBox P1 
      BackColor       =   &H00800000&
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   0
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   7
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFCEBB&
      Caption         =   "Computadoras en la Red"
      Height          =   5775
      Left            =   2880
      TabIndex        =   6
      Top             =   120
      Width           =   2655
      Begin VB.CommandButton CmdAgregar 
         BackColor       =   &H00FFCEBB&
         Caption         =   "&Agregar"
         Height          =   375
         Left            =   600
         TabIndex        =   3
         Top             =   5160
         Width           =   1455
      End
      Begin VB.ListBox LstNetwork 
         Height          =   4560
         Left            =   120
         Sorted          =   -1  'True
         Style           =   1  'Checkbox
         TabIndex        =   1
         Top             =   360
         Width           =   2415
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFCEBB&
      Caption         =   "Computadoras Registradas"
      Height          =   5775
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   2655
      Begin VB.CommandButton CmdSalir 
         BackColor       =   &H00FFCEBB&
         Caption         =   "&Cerrar"
         Height          =   375
         Left            =   670
         TabIndex        =   2
         Top             =   5160
         Width           =   1455
      End
      Begin VB.ListBox LstCompus 
         Height          =   4560
         Left            =   120
         Sorted          =   -1  'True
         Style           =   1  'Checkbox
         TabIndex        =   0
         Top             =   360
         Width           =   2415
      End
   End
   Begin VB.TextBox TxtNet 
      Height          =   195
      Left            =   3360
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   6120
      Visible         =   0   'False
      Width           =   1575
   End
End
Attribute VB_Name = "FrmNetwork"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Function ObtenerRed() As Boolean
On Error GoTo Salir
    MDIPrincipal.Enabled = False
    Dim ST As String
    Dim SW As ClsNetwork
    Set SW = New ClsNetwork
    Dim i As Long
    'Call SW.SetResourceType(0)
    Call SW.Reset
    TxtNet.Text = ""
    TxtNet.Text = SW.GetServerList
    ST = Replace(TxtNet.Text, "\\", "")
    TxtNet.Text = ST
    LstNetwork.Clear
    For i = 1 To Len(ST)
        If Mid(ST, i, 1) = "," Then
        LstNetwork.AddItem (Left(ST, i - 1))
        ST = Right(ST, Len(ST) - i)
        i = 1
        End If
    Next i
    MDIPrincipal.Enabled = True
    Exit Function
Salir:
    MDIPrincipal.Enabled = True
End Function

Public Sub ObtenerCompus()
    
    ConexionPrincipal
    Sql = "select * from Tbl_Maquina"
    Rs.Open Sql, Conecta, adOpenDynamic, adLockBatchOptimistic
    LstCompus.Clear
    Do While Not Rs.EOF
        LstCompus.AddItem Rs!C_Maq
        Rs.MoveNext
    Loop
    
End Sub

Private Function Existe(Pc As String) As Boolean
    ConexionPrincipal
    Sql = "select * from Tbl_Maquina where C_Maq='" & UCase(Pc) & "'"
    Rs.Open Sql, Conecta, adOpenDynamic, adLockBatchOptimistic
    
    If Not Rs.EOF Then
        MsgBox "Esta maquina ya esta dada de alta!!!"
        Existe = True
    Else
        ConexionPrincipal
        Sql = "insert into Tbl_Maquina(C_Maq) values('" & UCase(Pc) & "')"
        Rs.Open Sql, Conecta, adOpenDynamic, adLockBatchOptimistic
        FrmPrincipal.AgregarMaquinasLVM
        Existe = False
    End If
End Function

Private Sub CmdActualizar_Click()
    ObtenerCompus
    ObtenerRed
End Sub

Private Sub CmdAgregar_Click()
    Dim MSelec As Boolean
    Dim CantNet As Integer
    Dim i As Integer

    CantNet = LstNetwork.ListCount - 1

    For i = 0 To CantNet
        MSelec = LstNetwork.Selected(i)
        If MSelec = True Then
            If Existe(LstNetwork.List(i)) = False Then
                LstCompus.AddItem LstNetwork.List(i): Exit For
            End If
        End If
        CantNet = LstNetwork.ListCount - 1
    Next
End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    ObtenerCompus
End Sub

Private Sub Form_Load()
    ObtenerRed
    MDIPrincipal.AgregarVentana Me, "Red", "Computadoras en red..."
    InsertarPicture Me
    PosicionInicial Me
End Sub

Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then
        Visible = False
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MDIPrincipal.RemoverVentana Me, "Red"
End Sub
