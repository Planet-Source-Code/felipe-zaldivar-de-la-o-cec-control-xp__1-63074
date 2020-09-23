VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form FrmEscuela 
   BackColor       =   &H00FFCEBB&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "::: Datos de la institución :::"
   ClientHeight    =   2550
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6435
   Icon            =   "FrmEscuela.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2550
   ScaleWidth      =   6435
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00E89C78&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   120
      ScaleHeight     =   615
      ScaleWidth      =   6180
      TabIndex        =   12
      Top             =   1800
      Width           =   6180
      Begin VB.CommandButton CmdModificar 
         BackColor       =   &H00E89C78&
         Caption         =   "&Modificar"
         Height          =   375
         Left            =   1200
         TabIndex        =   4
         Top             =   120
         Width           =   1695
      End
      Begin VB.CommandButton CmdCambiar 
         BackColor       =   &H00E89C78&
         Caption         =   "&Cambiar  Imagen"
         Height          =   375
         Left            =   3360
         TabIndex        =   5
         Top             =   120
         Width           =   1695
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00E89C78&
      BorderStyle     =   0  'None
      Height          =   1605
      Left            =   120
      ScaleHeight     =   1605
      ScaleWidth      =   6180
      TabIndex        =   6
      Top             =   120
      Width           =   6180
      Begin VB.TextBox TxtTelefono 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   960
         MaxLength       =   20
         TabIndex        =   3
         Top             =   1200
         Width           =   3495
      End
      Begin VB.TextBox TxtDomicilio 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   960
         MaxLength       =   100
         TabIndex        =   2
         Top             =   840
         Width           =   3495
      End
      Begin VB.TextBox TxtNombre 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   960
         MaxLength       =   100
         TabIndex        =   1
         Top             =   480
         Width           =   3495
      End
      Begin VB.PictureBox PicEscuela 
         BackColor       =   &H00D05C28&
         BorderStyle     =   0  'None
         Height          =   1395
         Left            =   4680
         Picture         =   "FrmEscuela.frx":08CA
         ScaleHeight     =   93
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   91
         TabIndex        =   7
         Top             =   120
         Width           =   1360
      End
      Begin VB.TextBox TxtClave 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   960
         MaxLength       =   50
         TabIndex        =   0
         Top             =   120
         Width           =   3495
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Telefono:"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   1200
         Width           =   675
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Domicilio"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   840
         Width           =   630
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre:"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   480
         Width           =   600
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Clave:"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   120
         Width           =   450
      End
   End
   Begin MSComDlg.CommonDialog CDE 
      Left            =   5520
      Top             =   2280
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "FrmEscuela"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim PBag As PropertyBag
Dim logo1() As Byte

Public Sub CargarInf()
On Error Resume Next
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
        CargarControles
    End If
End Sub

Private Sub CmdModificar_Click()
    If Preguntar("Los datos son correctos?") = False Then Exit Sub
    Set PBag = New PropertyBag
    PBag.WriteProperty "CBT2", PicEscuela.Picture
    logo1 = PBag.Contents
    RsCE.Close
    ConexionEscuelas
    SQLCE = "Select * from Tbl_Escuela where IDESC='CBT2'"
    RsCE.Open SQLCE, Conecta, adOpenStatic, adLockOptimistic
    If Not RsCE.EOF = True Then
        RsCE!C_Escuela = TxtClave.Text
        RsCE!Escuela = TxtNombre.Text
        RsCE!Domicilio = TxtDomicilio.Text
        RsCE!Telefono = TxtTelefono.Text
        RsCE!Logotipo1 = logo1
        RsCE.Update
    End If
    If FrmConsProcesos.Visible = True Then FrmConsProcesos.CargarInf
End Sub

Private Sub CmdCambiar_Click()
On Error GoTo a:
    Dim FileName As String
    CDE.CancelError = True
    CDE.DialogTitle = "::: Selecciona una imagen para tu institución ( 91 x 93 pixeles ) :::"
    CDE.InitDir = App.Path
    CDE.Flags = &H4
    CDE.DefaultExt = "jpg"
    CDE.Filter = "Archivos de Imagen (*.bmp,*.gif,*.jpg,*.jpeg,*.jpe,*.jfif)|*.bmp;*.gif;*.jpg;*.jpeg;*.jpe;*.jfif)"
    CDE.ShowOpen
    FileName = (CDE.FileName)
    PicEscuela.Picture = LoadPicture(FileName)
a:
End Sub

Private Sub Form_Load()
    MDIPrincipal.AgregarVentana Me, "Institución", "Datos de la institución..."
    Redondear Me
    PosicionInicial Me
    CargarInf
End Sub

Private Sub CargarControles()
    TxtClave.Text = RsCE!C_Escuela
    TxtNombre.Text = RsCE!Escuela
    TxtDomicilio.Text = RsCE!Domicilio
    TxtTelefono.Text = RsCE!Telefono
End Sub

Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then
        Visible = False
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MDIPrincipal.RemoverVentana Me, "Institución"
End Sub

Private Sub TxtClave_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And TxtClave.Text <> "" Then
        TxtNombre.SetFocus
    End If
End Sub

Private Sub TxtDomicilio_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And TxtDomicilio.Text <> "" Then
        TxtTelefono.SetFocus
    End If
End Sub

Private Sub TxtNombre_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And TxtNombre.Text <> "" Then
        TxtDomicilio.SetFocus
    End If
End Sub

Private Sub TxtTelefono_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And TxtTelefono.Text <> "" Then
        CmdModificar.SetFocus
    End If
End Sub
