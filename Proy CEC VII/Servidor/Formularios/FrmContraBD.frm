VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmContraBD 
   BackColor       =   &H00FFCEBB&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "::: Selecciona la base de datos :::"
   ClientHeight    =   1530
   ClientLeft      =   5550
   ClientTop       =   4980
   ClientWidth     =   4920
   Icon            =   "FrmContraBD.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1530
   ScaleWidth      =   4920
   Begin VB.Data DataChec 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   5040
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1080
      Width           =   1140
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFCEBB&
      Caption         =   "..."
      Height          =   375
      Left            =   4440
      TabIndex        =   4
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   2880
      TabIndex        =   3
      Top             =   1080
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   1320
      TabIndex        =   2
      Top             =   1080
      Width           =   1455
   End
   Begin VB.TextBox Txtclave 
      BorderStyle     =   0  'None
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1320
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   120
      Width           =   3015
   End
   Begin VB.TextBox txtbd 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   600
      Width           =   3015
   End
   Begin MSComDlg.CommonDialog CDBD 
      Left            =   5040
      Top             =   600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Contraseña:"
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   240
      Width           =   855
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Base de Datos:"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   720
      Width           =   1095
   End
End
Attribute VB_Name = "FrmContraBD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim respuesta As String

Private Sub GuardarDir()
inicio:
Dim D1 As String
D1 = GetSetting(App.Title, "DireccionBDCEC1", "Ruta", InitDireccion)
If D1 <> "" Then
    'Direccion = GetSetting(App.Title, "DireccionBDCEC1", "Ruta", InitDireccion) 'asignamos la
    C_BD = GetSetting(App.Title, "DireccionBDCEC1", "C_BDs", C_BD)
    If ValidarBD(Direccion) = True Then
        DeleteSetting App.Title, "DireccionBDCEC1", "Ruta"
        DeleteSetting App.Title, "DireccionBDCEC1", "C_BDs"
        Me.Visible = True
        GoTo inicio
    End If
    C_BD1 = C_BD
    MDIPrincipal.CmdMenu.Enabled = True
    MDIPrincipal.MenuDespleglable 0
    FrmPrincipal.Top = 0
    FrmPrincipal.Left = 0
    FrmPrincipal.Show
    Direccion = App.Path & "\BD\Bdcec1.mdb"
End If
End Sub

Private Sub AbrirBD()
'Obtenemos el path de la base de datos
On Error GoTo Problema
    CDBD.CancelError = True
    CDBD.DialogTitle = "::: Control de Acceso CEC1 ::: Selecciona la Base de Datos Indicada"
    CDBD.InitDir = App.Path
    CDBD.Flags = &H4
    CDBD.DefaultExt = "mdb"
    CDBD.Filter = "Microsoft Office Access (*.mdb)|*.mdb"
    CDBD.ShowOpen
    InitDireccion = (CDBD.FileName)
    C_BD = TxtClave
    'If C_BD = "" Then Primera = True: Exit Sub
    If ValidarBD(InitDireccion) = True Then
        Primera = True
        TxtClave.Locked = False
        TxtClave.Text = ""
        txtbd.Text = ""
    Else
        Primera = False
        txtbd.Text = InitDireccion
        TxtClave.Locked = True
        Exit Sub
    End If

Problema:
    If Err.Number = 0 Then
        If GetSetting(App.Title, "DireccionBDCEC1", "Ruta", InitDireccion) = "" Then
            If ValidarBD(InitDireccion) = True Then
                CDBD.FileTitle = ""
                InitDireccion = ""
                Primera = True
            Else
                Primera = False
            End If
        Else
            If ValidarBD(InitDireccion) = False Then
                CDBD.FileTitle = ""
                InitDireccion = ""
                Primera = True
            Else
                Primera = False
            End If
        End If
        
    ElseIf Err.Number = 32755 Then
        If GetSetting(App.Title, "DireccionBDCEC1", "Ruta", InitDireccion) = "" Then Call TerminarProg
        If ValidarBD(InitDireccion) = False Then Primera = False: Exit Sub
        Call TerminarProg
        Exit Sub
    ElseIf Err.Number <> 0 Then
            If ValidarBD(InitDireccion) = False Then Primera = False: Exit Sub
            Mensaje = "Base de Datos Incorrecta!!!" & Chr(10) & Chr(10) & "Si:Volver a Intentar | No: Salir del Programa"
            respuesta = MsgBox(Mensaje, 4 + 32 + 0, "Atención")
            If respuesta = vbYes Then
                Primera = True
            Else
                If GetSetting(App.Title, "DireccionBDCEC1", "Ruta", InitDireccion) = "" Then
                    Call TerminarProg
                Else
                    If ValidarBD(GetSetting(App.Title, "DireccionBDCEC1", "Ruta", InitDireccion)) = False Then
                        C_BD = ""
                        InitDireccion = ""
                        Primera = False
                        Exit Sub
                    End If
                        If ValidarBD(GetSetting(App.Title, "DireccionBDCEC1", "Ruta", InitDireccion)) = True Then
                            If GetSetting(App.Title, "DireccionBDCEC1", "Ruta", InitDireccion) = "" Then
                                Call TerminarProg
                            Else
                                C_BD = ""
                                InitDireccion = ""
                                Primera = False
                                Exit Sub
                            End If
                        End If
                End If
            End If
    End If
End Sub

Private Sub Command1_Click()

    SaveSetting App.Title, "DireccionBDCEC1", "Ruta", InitDireccion 'Guardamos la direccion en el registro de la maquina
    SaveSetting App.Title, "DireccionBDCEC1", "C_BDs", C_BD
    Direccion = GetSetting(App.Title, "DireccionBDCEC1", "Ruta", InitDireccion) 'asignamos la
    C_BD1 = GetSetting(App.Title, "DireccionBDCEC1", "C_BDs", C_BD)
    
    MDIPrincipal.CmdMenu.Enabled = True
    MDIPrincipal.MenuDespleglable 0
    FrmPrincipal.Top = 0
    FrmPrincipal.Left = 0
    FrmPrincipal.Show
End Sub

Private Sub Command2_Click()
    Unload Me
    End
End Sub

Private Sub Command3_Click()
    Primera = True
    
    Do While Primera = True
    
        If TxtClave.Text = "" Then
            MsgBox "Escribe la contraseña de la base de datos", , "Atencón!!!"
            TxtClave.SetFocus
            Exit Sub
        End If

        Call AbrirBD 'Abrimos la base de datos
        'If Primera = True Then MsgBox "Debes especificar la base de datos!!!"
    Loop
    
End Sub

Private Sub Form_Load()
    txtbd.Text = ""
    TxtClave.Text = ""
    GuardarDir
    PosicionInicial Me
End Sub
