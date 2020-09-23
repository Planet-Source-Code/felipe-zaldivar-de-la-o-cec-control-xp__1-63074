VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.MDIForm MDIPrincipal 
   AutoShowChildren=   0   'False
   BackColor       =   &H00FAE5D3&
   Caption         =   "MDIPrincipal"
   ClientHeight    =   10635
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8520
   Icon            =   "MDIPrincipal.frx":0000
   LinkTopic       =   "MDIForm1"
   WindowState     =   2  'Maximized
   Begin VB.Data DataHisorialServer 
      Align           =   2  'Align Bottom
      Caption         =   "DataHisorialServer"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9135
      Visible         =   0   'False
      Width           =   8520
   End
   Begin MSComctlLib.ImageList IL1 
      Left            =   5040
      Top             =   2400
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin VB.PictureBox PicVentanas 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E89C78&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1155
      Left            =   0
      ScaleHeight     =   1155
      ScaleWidth      =   8520
      TabIndex        =   26
      Top             =   9480
      Width           =   8520
      Begin MSComctlLib.ListView LV1 
         Height          =   1065
         Left            =   0
         TabIndex        =   19
         Top             =   45
         Width           =   9735
         _ExtentX        =   17171
         _ExtentY        =   1879
         Arrange         =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         OLEDragMode     =   1
         _Version        =   393217
         Icons           =   "imlIcons"
         ForeColor       =   16777215
         BackColor       =   13655080
         Appearance      =   0
         OLEDragMode     =   1
         NumItems        =   0
      End
   End
   Begin VB.PictureBox PicContenedor 
      Align           =   3  'Align Left
      AutoRedraw      =   -1  'True
      BackColor       =   &H00D05C28&
      BorderStyle     =   0  'None
      Height          =   9135
      Left            =   0
      ScaleHeight     =   9135
      ScaleWidth      =   4005
      TabIndex        =   21
      Top             =   0
      Width           =   4005
      Begin VB.VScrollBar SBMenu 
         CausesValidation=   0   'False
         Height          =   3640
         LargeChange     =   1800
         Left            =   3720
         SmallChange     =   300
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   480
         Value           =   15
         Width           =   255
      End
      Begin VB.PictureBox PicMenu 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00E89C78&
         BorderStyle     =   0  'None
         Height          =   9855
         Left            =   0
         ScaleHeight     =   9855
         ScaleWidth      =   3735
         TabIndex        =   22
         Top             =   0
         Width           =   3735
         Begin VB.PictureBox f1 
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            Height          =   1575
            Index           =   3
            Left            =   240
            ScaleHeight     =   1575
            ScaleWidth      =   2535
            TabIndex        =   38
            Top             =   5280
            Width           =   2535
            Begin VB.CommandButton CmdBloquear 
               BackColor       =   &H80000005&
               Caption         =   "Bloquear Servidor"
               Height          =   255
               Left            =   240
               TabIndex        =   12
               Top             =   120
               Width           =   2055
            End
            Begin VB.CommandButton Command4 
               BackColor       =   &H80000005&
               Caption         =   "Acerca de ..."
               Height          =   255
               Left            =   240
               TabIndex        =   15
               Top             =   1200
               Width           =   2055
            End
            Begin VB.CommandButton CmbCU 
               BackColor       =   &H80000005&
               Caption         =   "Cambiar de Usuario"
               Height          =   255
               Left            =   240
               TabIndex        =   13
               Top             =   480
               Width           =   2055
            End
            Begin VB.CommandButton CmbApagar 
               BackColor       =   &H80000005&
               Caption         =   "Apagar Servidor"
               Height          =   255
               Left            =   240
               TabIndex        =   14
               Top             =   840
               Width           =   2055
            End
         End
         Begin VB.PictureBox p1 
            AutoRedraw      =   -1  'True
            BackColor       =   &H00D05C28&
            BorderStyle     =   0  'None
            Height          =   495
            Index           =   3
            Left            =   240
            ScaleHeight     =   495
            ScaleWidth      =   2535
            TabIndex        =   36
            Top             =   4800
            Width           =   2535
            Begin VB.Image ImgMain 
               Height          =   240
               Index           =   3
               Left            =   2160
               Picture         =   "MDIPrincipal.frx":08CA
               Top             =   120
               Width           =   240
            End
            Begin VB.Line lnBack 
               BorderColor     =   &H00FFFFFF&
               Index           =   4
               Visible         =   0   'False
               X1              =   0
               X2              =   6465
               Y1              =   480
               Y2              =   480
            End
            Begin VB.Label lblCInfo 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "CEC - CONTROL"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   195
               Index           =   4
               Left            =   600
               TabIndex        =   37
               Top             =   150
               Width           =   1440
            End
            Begin VB.Image imgPicTitle 
               Height          =   440
               Index           =   4
               Left            =   40
               Picture         =   "MDIPrincipal.frx":0C54
               Stretch         =   -1  'True
               Top             =   30
               Width           =   440
            End
         End
         Begin VB.PictureBox p1 
            AutoRedraw      =   -1  'True
            BackColor       =   &H00D05C28&
            BorderStyle     =   0  'None
            Height          =   495
            Index           =   4
            Left            =   240
            ScaleHeight     =   495
            ScaleWidth      =   2535
            TabIndex        =   34
            Top             =   6840
            Width           =   2535
            Begin VB.Image ImgMain 
               Height          =   240
               Index           =   4
               Left            =   2160
               Picture         =   "MDIPrincipal.frx":151E
               Top             =   120
               Width           =   240
            End
            Begin VB.Image imgPicTitle 
               Height          =   440
               Index           =   3
               Left            =   40
               Picture         =   "MDIPrincipal.frx":18A8
               Stretch         =   -1  'True
               Top             =   30
               Width           =   440
            End
            Begin VB.Label lblCInfo 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Ventanas"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   195
               Index           =   3
               Left            =   600
               TabIndex        =   35
               Top             =   150
               Width           =   810
            End
            Begin VB.Line lnBack 
               BorderColor     =   &H00FFFFFF&
               Index           =   3
               Visible         =   0   'False
               X1              =   0
               X2              =   6465
               Y1              =   480
               Y2              =   480
            End
         End
         Begin VB.PictureBox f1 
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            Height          =   1815
            Index           =   4
            Left            =   240
            ScaleHeight     =   1815
            ScaleWidth      =   2535
            TabIndex        =   33
            Top             =   7320
            Width           =   2535
            Begin VB.CommandButton CmdCerrar 
               BackColor       =   &H80000005&
               Caption         =   "Cerrar Ventanas"
               Height          =   255
               Left            =   240
               TabIndex        =   40
               Top             =   1440
               Width           =   2055
            End
            Begin VB.CommandButton CmdMin 
               BackColor       =   &H80000005&
               Caption         =   "Minimizar Ventanas"
               Height          =   255
               Left            =   240
               TabIndex        =   39
               Top             =   1080
               Width           =   2055
            End
            Begin VB.OptionButton OptC 
               BackColor       =   &H80000005&
               Caption         =   "C&ascada"
               Height          =   255
               Left            =   240
               TabIndex        =   16
               Top             =   120
               Width           =   1695
            End
            Begin VB.OptionButton OptMV 
               BackColor       =   &H80000005&
               Caption         =   "Mosaico ver&tical"
               Height          =   255
               Left            =   240
               TabIndex        =   18
               Top             =   720
               Width           =   1695
            End
            Begin VB.OptionButton OptMH 
               BackColor       =   &H80000005&
               Caption         =   "Mo&saico horizontal"
               Height          =   255
               Left            =   240
               TabIndex        =   17
               Top             =   420
               Width           =   1815
            End
         End
         Begin VB.PictureBox f1 
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            Height          =   1095
            Index           =   2
            Left            =   240
            ScaleHeight     =   1095
            ScaleWidth      =   2535
            TabIndex        =   32
            Top             =   3720
            Width           =   2535
            Begin VB.CommandButton Command2 
               BackColor       =   &H80000005&
               Caption         =   "&Institución"
               Height          =   255
               Left            =   240
               TabIndex        =   11
               Top             =   720
               Width           =   2055
            End
            Begin VB.CommandButton CmdImpresora 
               BackColor       =   &H80000005&
               Caption         =   "&Control de Impresoras"
               Height          =   495
               Left            =   1320
               TabIndex        =   10
               Top             =   720
               Visible         =   0   'False
               Width           =   975
            End
            Begin VB.CommandButton CmdVisor 
               BackColor       =   &H80000005&
               Caption         =   "&Visor de imágenes"
               Height          =   495
               Left            =   1320
               TabIndex        =   9
               Top             =   120
               Width           =   975
            End
            Begin VB.CommandButton CmdClientes 
               BackColor       =   &H80000005&
               Caption         =   "&Buscar Maquinas"
               Height          =   495
               Left            =   240
               TabIndex        =   8
               Top             =   120
               Width           =   975
            End
         End
         Begin VB.PictureBox p1 
            AutoRedraw      =   -1  'True
            BackColor       =   &H00D05C28&
            BorderStyle     =   0  'None
            Height          =   495
            Index           =   2
            Left            =   240
            ScaleHeight     =   495
            ScaleWidth      =   2535
            TabIndex        =   30
            Top             =   3240
            Width           =   2535
            Begin VB.Image ImgMain 
               Height          =   240
               Index           =   2
               Left            =   2160
               Picture         =   "MDIPrincipal.frx":2572
               Top             =   120
               Width           =   240
            End
            Begin VB.Line lnBack 
               BorderColor     =   &H00FFFFFF&
               Index           =   2
               Visible         =   0   'False
               X1              =   0
               X2              =   6465
               Y1              =   480
               Y2              =   480
            End
            Begin VB.Label lblCInfo 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Utilerías"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   195
               Index           =   2
               Left            =   600
               TabIndex        =   31
               Top             =   150
               Width           =   735
            End
            Begin VB.Image imgPicTitle 
               Height          =   440
               Index           =   2
               Left            =   40
               Picture         =   "MDIPrincipal.frx":28FC
               Stretch         =   -1  'True
               Top             =   30
               Width           =   440
            End
         End
         Begin VB.PictureBox f1 
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            Height          =   1335
            Index           =   1
            Left            =   240
            ScaleHeight     =   1335
            ScaleWidth      =   2535
            TabIndex        =   29
            Top             =   1920
            Width           =   2535
            Begin VB.CommandButton CmdMenuProc 
               BackColor       =   &H80000005&
               Caption         =   "Pr&ocesos"
               Height          =   495
               Left            =   1320
               TabIndex        =   7
               Top             =   720
               Width           =   975
            End
            Begin VB.CommandButton CmdMenuHist 
               BackColor       =   &H80000005&
               Caption         =   "&Historial"
               Height          =   495
               Left            =   240
               TabIndex        =   4
               Top             =   120
               Width           =   975
            End
            Begin VB.CommandButton CmdOPR 
               BackColor       =   &H80000005&
               Caption         =   "&Procesos restringidos"
               Height          =   495
               Left            =   240
               TabIndex        =   6
               Top             =   720
               Width           =   975
            End
            Begin VB.CommandButton CmdLog 
               BackColor       =   &H80000005&
               Caption         =   "&Log"
               Height          =   495
               Left            =   1320
               TabIndex        =   5
               Top             =   120
               Width           =   975
            End
         End
         Begin VB.PictureBox p1 
            AutoRedraw      =   -1  'True
            BackColor       =   &H00D05C28&
            BorderStyle     =   0  'None
            Height          =   495
            Index           =   1
            Left            =   240
            ScaleHeight     =   495
            ScaleWidth      =   2535
            TabIndex        =   27
            Top             =   1440
            Width           =   2535
            Begin VB.Image ImgMain 
               Height          =   240
               Index           =   1
               Left            =   2160
               Picture         =   "MDIPrincipal.frx":31C6
               Top             =   120
               Width           =   240
            End
            Begin VB.Line lnBack 
               BorderColor     =   &H00FFFFFF&
               Index           =   1
               Visible         =   0   'False
               X1              =   0
               X2              =   6465
               Y1              =   480
               Y2              =   480
            End
            Begin VB.Label lblCInfo 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Opciones"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   195
               Index           =   1
               Left            =   600
               TabIndex        =   28
               Top             =   150
               Width           =   810
            End
            Begin VB.Image imgPicTitle 
               Height          =   440
               Index           =   1
               Left            =   40
               Picture         =   "MDIPrincipal.frx":3550
               Stretch         =   -1  'True
               Top             =   30
               Width           =   440
            End
         End
         Begin VB.PictureBox p1 
            AutoRedraw      =   -1  'True
            BackColor       =   &H00D05C28&
            BorderStyle     =   0  'None
            Height          =   495
            Index           =   0
            Left            =   240
            ScaleHeight     =   495
            ScaleWidth      =   2535
            TabIndex        =   24
            Top             =   120
            Width           =   2535
            Begin VB.Image ImgMain 
               Height          =   240
               Index           =   0
               Left            =   2160
               Picture         =   "MDIPrincipal.frx":3E1A
               Top             =   120
               Width           =   240
            End
            Begin VB.Image imgPicTitle 
               Height          =   440
               Index           =   0
               Left            =   40
               Picture         =   "MDIPrincipal.frx":41A4
               Stretch         =   -1  'True
               Top             =   30
               Width           =   440
            End
            Begin VB.Label lblCInfo 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Principal"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   195
               Index           =   0
               Left            =   600
               TabIndex        =   25
               Top             =   150
               Width           =   750
            End
            Begin VB.Line lnBack 
               BorderColor     =   &H00FFFFFF&
               Index           =   0
               Visible         =   0   'False
               X1              =   0
               X2              =   6465
               Y1              =   480
               Y2              =   480
            End
         End
         Begin VB.PictureBox f1 
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            Height          =   855
            Index           =   0
            Left            =   240
            ScaleHeight     =   855
            ScaleWidth      =   2535
            TabIndex        =   23
            Top             =   600
            Width           =   2535
            Begin VB.CommandButton CmdOUsr 
               BackColor       =   &H80000005&
               Caption         =   "&Usuarios"
               Height          =   255
               Left            =   1320
               TabIndex        =   3
               Top             =   480
               Width           =   975
            End
            Begin VB.CommandButton CmdOGrupos 
               BackColor       =   &H80000005&
               Caption         =   "&Grupos"
               Height          =   255
               Left            =   240
               TabIndex        =   0
               Top             =   120
               Width           =   975
            End
            Begin VB.CommandButton CmdMenuRep 
               BackColor       =   &H80000005&
               Caption         =   "&Maquina"
               Height          =   255
               Left            =   1320
               TabIndex        =   1
               Top             =   120
               Width           =   975
            End
            Begin VB.CommandButton CmdMenuUsr 
               BackColor       =   &H80000005&
               Caption         =   "&Reportes"
               Height          =   255
               Left            =   240
               TabIndex        =   2
               Top             =   480
               Width           =   975
            End
         End
         Begin VB.Image ImgDer 
            Height          =   240
            Left            =   3240
            Picture         =   "MDIPrincipal.frx":4A6E
            Top             =   1440
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.Image ImgDer2 
            Height          =   240
            Left            =   3240
            Picture         =   "MDIPrincipal.frx":4DF8
            Top             =   1680
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.Image ImgIzq2 
            Height          =   240
            Left            =   3240
            Picture         =   "MDIPrincipal.frx":5382
            Top             =   1200
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.Image ImgIzq 
            Height          =   240
            Left            =   3240
            Picture         =   "MDIPrincipal.frx":590C
            Top             =   960
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.Image ImgAbajo2 
            Height          =   240
            Left            =   2880
            Picture         =   "MDIPrincipal.frx":5C96
            Top             =   1680
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.Image ImgAbajo 
            Height          =   240
            Left            =   2880
            Picture         =   "MDIPrincipal.frx":6220
            Top             =   1440
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.Image ImgArriba2 
            Height          =   240
            Left            =   2880
            Picture         =   "MDIPrincipal.frx":65AA
            Top             =   1200
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.Image ImgArriba 
            Height          =   240
            Left            =   2880
            Picture         =   "MDIPrincipal.frx":6B34
            Top             =   960
            Visible         =   0   'False
            Width           =   240
         End
      End
      Begin VB.Image ImgRoll 
         Height          =   240
         Left            =   3720
         Picture         =   "MDIPrincipal.frx":6EBE
         Top             =   120
         Width           =   240
      End
   End
End
Attribute VB_Name = "MDIPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long

Dim Tope As Long
Dim hSysMenu As Long ' hwnd para remover el boton (X)
Dim MouseOpt  As Boolean


Private Sub CmbApagar_Click()
    ActivarForm FrmCUsr
    FrmCUsr.OpcionAcceso = 2
    FrmCUsr.Caption = "::: Agagar Servidor :::"
    FrmCUsr.PicMain.Picture = FrmCUsr.PicApagar.Picture
End Sub

Private Sub CmbCU_Click()
    ActivarForm FrmCUsr
    FrmCUsr.OpcionAcceso = 1
    FrmCUsr.Caption = "::: Cambio de usuario :::"
    FrmCUsr.PicMain.Picture = FrmCUsr.PicCU.Picture
End Sub

Private Sub CmdBloquear_Click()
    ActivarForm FrmCUsr
    FrmCUsr.OpcionAcceso = 3
    FrmCUsr.Caption = "::: Bloquear servidor :::"
    FrmCUsr.PicMain.Picture = FrmCUsr.PicBloq.Picture
End Sub

Private Sub CmdCerrar_Click()
    On Error Resume Next
    Dim frm As Form
    MDIPrincipal.Enabled = False
    For Each frm In Forms
        If frm.Name <> "MDIPrincipal" And frm.Name <> "FrmPrincipal" Then
            Unload frm
        End If
        DoEvents
    Next
    MDIPrincipal.Enabled = True
End Sub

Private Sub CmdClientes_Click()
    ActivarForm FrmNetwork
End Sub

Private Sub CmdLog_Click()
    ActivarForm FrmLog
End Sub

Public Sub MenuDespleglable(OpcionMenuDesplegable As Integer)
    Dim TTemporal As Long
    Dim i As Integer
    If OpcionMenuDesplegable = 1 Then
        ImgRoll.Picture = ImgDer.Picture
        TTemporal = SBMenu.Left + SBMenu.Width - 415
        TTemporal = (TTemporal / 20)
        For i = 1 To 20
            PicContenedor.Width = PicContenedor.Width - TTemporal
            Call PicContenedor.Refresh
        Next
        PicContenedor.Width = 415
        PicMenu.Visible = False
        ImgRoll.Left = (ImgRoll.Height / 4)
    Else
        ImgRoll.Picture = ImgIzq.Picture
        TTemporal = SBMenu.Left + SBMenu.Width
        TTemporal = (TTemporal / 20)
        PicMenu.Visible = True
        ImgRoll.Left = PicMenu.Width
        For i = 1 To 20
            PicContenedor.Width = PicContenedor.Width + TTemporal
            Call PicContenedor.Refresh
        Next
        PicContenedor.Width = SBMenu.Left + SBMenu.Width
    End If
End Sub

Private Sub CmdMenuHist_Click()
    ActivarForm FrmHistorial
End Sub

Private Sub CmdMin_Click()
    On Error Resume Next
    Dim frm As Form
    MDIPrincipal.Enabled = False
    For Each frm In Forms
        If frm.Name <> "MDIPrincipal" And frm.Name <> "FrmPrincipal" Then
            frm.WindowState = vbMinimized
            frm.Visible = False
        End If
        DoEvents
    Next
    MDIPrincipal.Enabled = True
End Sub

Private Sub CmdOGrupos_Click()
    ActivarForm FrmGrupos
End Sub

Private Sub CmdOPR_Click()
    ActivarForm FrmAltaProcesos
End Sub

Private Sub CmdMenuProc_Click()
    ActivarForm FrmConsProcesos
End Sub

Private Sub CmdMenuRep_Click()
    ActivarForm FrmConsMaquina
End Sub

Private Sub CmdMenuUsr_Click()
    ActivarForm FrmConsReportes
End Sub

Private Sub CmdOUsr_Click()
    ActivarForm FrmAltasAlumnos
End Sub

Private Sub CmdVisor_Click()
    ActivarForm FrmVisor
End Sub

Private Sub Command2_Click()
    ActivarForm FrmEscuela
End Sub

Private Sub RevisarBD()
    Dim FechaE As Date
    Dim Dias As Integer
    ConexionPrincipal
    Sql = "select * from Tbl_Config where Clave='Config'"
    Rs.Open Sql, Conecta, adOpenStatic, adLockPessimistic
    If Rs.EOF Then
        Rs.AddNew
        Rs!Clave = "Config"
        Rs.Update
    End If
    Dias = Rs!Dias_Eliminar
    If Dias <= 0 Then Exit Sub
    FechaE = DateAdd("d", -Dias, Format(Date, "mm/dd/yyyy"))
    Call eliminar("Tbl_Reportes", "Fecha_Reporte", FechaE)
    Call eliminar("Tbl_Procesos_Reg", "Fecha", FechaE)
    Call eliminar("Tbl_Historial", "Fecha_Entrada", FechaE)
End Sub

Private Sub eliminar(Tabla As String, Campo As String, FechaE As Date)
    ConexionHistorial
    SqlCH = "Delete * From " & Tabla & " where " & Campo & " <=#" & FechaE & "#"
    RsCH.Open SqlCH, Conecta, adOpenDynamic, adLockBatchOptimistic
    DoEvents
End Sub

Private Sub ImgMain_Click(Index As Integer)
    Dim Icontrol As Integer
    If f1(Index).Visible = True Then
        f1(Index).Visible = False
        ImgMain(Index).Picture = ImgAbajo.Picture
        For Icontrol = Index + 1 To P1.Count - 1
            If f1(Icontrol - 1).Visible = True Then
                P1(Icontrol).Top = f1(Icontrol - 1).Top + f1(Icontrol - 1).Height + Tope
            Else
                P1(Icontrol).Top = P1(Icontrol - 1).Top + P1(Icontrol - 1).Height + Tope
            End If
            f1(Icontrol).Top = P1(Icontrol).Top + P1(Icontrol).Height
        Next
    Else
        f1(Index).Visible = True
        ImgMain(Index).Picture = ImgArriba.Picture
        For Icontrol = Index + 1 To P1.Count - 1
            If f1(Icontrol - 1).Visible = True Then
                P1(Icontrol).Top = f1(Icontrol - 1).Top + f1(Icontrol - 1).Height + Tope
            Else
                P1(Icontrol).Top = P1(Icontrol - 1).Top + P1(Icontrol - 1).Height + Tope
            End If
            f1(Icontrol).Top = P1(Icontrol).Top + P1(Icontrol).Height
        Next
    End If
    Call MDIForm_Resize
End Sub

Private Sub ImgMain_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If f1(Index).Visible = True Then
        ImgMain(Index).Picture = ImgArriba2.Picture
    Else
        ImgMain(Index).Picture = ImgAbajo2.Picture
    End If
End Sub

Private Sub ImgRoll_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If PicMenu.Visible = True Then
        ImgRoll.Picture = ImgIzq2.Picture
        MenuDespleglable 1
    Else
        ImgRoll.Picture = ImgDer2.Picture
        MenuDespleglable 0
    End If
End Sub

Private Sub MDIForm_Initialize()
    If App.PrevInstance Then
        ActivatePrevInstance
    End If
    Call ModeStartUP
    Call ChkManifest
    Call ChkBasedeDatos
    Call CargarBD
    Call RevisarBD
    Call ChequeoGral(1)
    Call AgregarServidor
    InitCommonControls
    DoEvents
    FrmLogin.Show
End Sub

Private Sub AgregarServidor()
    ConexionMaquinas
    SqlCM = "Select * from Tbl_Maquina Where C_Maq='SERVIDOR'"
    RsCM.Open SqlCM, Conecta, adOpenStatic, adLockPessimistic
    If RsCM.EOF Then
        RsCM.AddNew
        RsCM!C_Maq = "SERVIDOR"
        RsCM.Update
    End If
    RsCM.Close
End Sub

Public Sub ChequeoGral(Opt As Integer)
    '////iniciamos el historial general
    '''Checamos que las maquinas que hayan quedado ocupadas
    '''las resetemos
    'desocupamos la maquina que tenia el usuario
    With DataHisorialServer
        .DatabaseName = Direccion
        .RecordSource = "SELECT * FROM [Tbl_Maquina] where Maq_Ocupada=true or Maq_Inicio=true or C_Acceso<>''"
        .Connect = ";Pwd=" & C_BD
        Call .Refresh
        If Not .Recordset.EOF Then .Recordset.MoveFirst
        Do While Not .Recordset.EOF
            .Recordset.Edit
            .Recordset("Maq_Ocupada") = False
            .Recordset("Maq_Inicio") = False
            .Recordset("Maq_Fin") = False
            .Recordset("C_Acceso") = ""
            .Recordset.Update
            If Not .Recordset.EOF Then .Recordset.MoveNext
        Loop
    
    '''borramos las maquinas que quedaron ocupadas por cualquier motivo
    '''jejeje
    
        .DatabaseName = Direccion
        .RecordSource = "SELECT * FROM [Tbl_Historial] where  Hora_Salida=null"
        .Connect = ";Pwd=" & C_BD
        Call .Refresh
        If Not .Recordset.EOF Then .Recordset.MoveFirst
        Do While Not .Recordset.EOF
            If IsNull(.Recordset("Hora_Salida")) Then
                .Recordset.Edit
                .Recordset("Hora_Salida") = Time
                .Recordset.Update
            End If
            If Not .Recordset.EOF Then .Recordset.MoveNext
        Loop
    End With
End Sub

Private Sub EstadoMenus(Edo As Boolean)
    Dim ICtrl As Integer
    For ICtrl = 0 To P1.Count - 1
        Call ImgMain_Click(ICtrl)
        f1(ICtrl).Visible = False
        ImgMain(ICtrl).Picture = ImgAbajo.Picture
        DoEvents
    Next
End Sub

Private Sub MDIForm_Load()

    hSysMenu = GetSystemMenu(Me.hwnd, False)
    RemoveMenu hSysMenu, SC_CLOSE, MF_BYCOMMAND

    Dim Icontrol As Integer
    Tope = P1(0).Top
    f1(0).Top = Tope + P1(0).Height
    For Icontrol = 0 To P1.Count - 1
        SetWindowRgn P1(Icontrol).hwnd, CreateRoundRectRgn(0, 0, P1(Icontrol).Width / 15, P1(Icontrol).Height / 5, 6, 6), True
        f1(Icontrol).Width = P1(Icontrol).Width - 11
        P1(Icontrol).Left = Tope
        f1(Icontrol).Left = Tope
    Next
    PicContenedor.BackColor = PicMenu.BackColor
    PicMenu.Width = Tope * 2 + P1(0).Width
    SBMenu.Left = PicMenu.Width
    ImgRoll.Left = PicMenu.Width
    PicContenedor.Width = SBMenu.Left + SBMenu.Width
    LV1.Width = Me.Width - 210
    ImgRoll.Enabled = False
    MenuDespleglable 1
    Call EstadoMenus(False)
    DoEvents
End Sub

Private Sub MDIForm_Resize()
    Dim MaxValue As Long
    Dim Icontrol As Integer
    Dim TamañoPV As Long
    Dim TamañoPM As Long
    Dim TamañoLV As Long
    
    MaxValue = 0
    For Icontrol = 0 To P1.Count - 1
        If f1(Icontrol).Visible = True Then
            MaxValue = MaxValue + (f1(Icontrol).Height + P1(Icontrol).Height + Tope)
        Else
            MaxValue = MaxValue + (P1(Icontrol).Height + Tope)
        End If
    Next
    MaxValue = MaxValue + f1(0).Top + P1(0).Top - Tope
    
    TamañoPV = (Me.Height - PicVentanas.Height) - ((ImgRoll.Top * 2) + (ImgRoll.Height * 2) + (Tope * 2)) - 50
    TamañoPM = Me.Height - PicVentanas.Height - 510
    TamañoLV = Me.Width - 210
    
    If TamañoLV > 0 Then LV1.Width = TamañoLV
    If TamañoPV > 0 Then SBMenu.Height = TamañoPV
    If TamañoPM > 0 Then PicMenu.Height = TamañoPM
    
    If MaxValue > (Me.Height - PicVentanas.Height) Then
        SBMenu.Visible = True
        SBMenu.Max = MaxValue - Me.Height + PicVentanas.Height
        SBMenu.Value = 0
    Else
        SBMenu.Visible = False
        PicMenu.Top = 0
    End If
    arrange2
End Sub

Private Sub OptC_Click()
    MDIPrincipal.Arrange 0
End Sub

Private Sub OptMH_Click()
    MDIPrincipal.Arrange 1
End Sub

Private Sub OptMV_Click()
    MDIPrincipal.Arrange 2
End Sub

Private Sub SBMenu_Change()
    PicMenu.Top = -SBMenu.Value
End Sub

Private Sub SBMenu_Scroll()
    SBMenu_Change
End Sub

Private Sub arrange2()
    Dim Icontrol As Integer
    For Icontrol = 0 To P1.Count - 1
        PicMenu.Height = PicMenu.Height + (f1(Icontrol).Height + P1(Icontrol).Height + Tope)
    Next
    PicMenu.Height = PicMenu.Height - PicVentanas.Height
End Sub

'*************************************************************
'Código PicVentanas

Public Sub AgregarVentana(frm As Form, Caption As String, Optional ToolTip As String)
    Dim LVItem As Integer
    Dim ILItem As Integer
    Dim ILIndex As Integer
    LVItem = BMVentana(Caption)
    ILItem = BMIVentana(Caption)
    
    If LVItem = 0 Then
        If ILItem = 0 Then
            IL1.ListImages.Add IL1.ListImages.Count + 1, Caption, frm.Icon
            ILIndex = IL1.ListImages.Count
        Else
            ILIndex = ILItem
        End If
        Set LV1.Icons = IL1
        LV1.ListItems.Add , Caption, Caption, ILIndex
        'LV1.ListItems(LV1.ListItems.Count).ToolTipText = ToolTip
        LV1.ListItems(LV1.ListItems.Count).Tag = frm.hwnd
    End If
    
End Sub

Public Sub RemoverVentana(Frm2 As Form, Caption As String)
    Dim LItem As Integer
    LItem = BMVentana(Caption)
    If LItem > 0 Then
        LV1.ListItems.Remove (LItem)
        Call LV1.Refresh
    End If
End Sub

Private Function BMIVentana(Caption As String) As Integer
    Dim ILItem As Integer
    BMIVentana = 0

    For ILItem = 1 To IL1.ListImages.Count
        If IL1.ListImages(ILItem).Key = Caption Then
            BMIVentana = ILItem
            Exit Function
        End If
    Next
End Function

Private Function BMVentana(Caption As String) As Integer
    Dim LVItem As Integer
    BMVentana = 0
    For LVItem = 1 To LV1.ListItems.Count
        If LV1.ListItems(LVItem).Key = Caption Then
            BMVentana = LVItem
            Exit Function
        End If
    Next
End Function

Private Sub LV1_Click()
    If LV1.SelectedItem Is Nothing Then Exit Sub
    Dim Frm3 As Form
    Dim hwnd As Long
    hwnd = LV1.SelectedItem.Tag
    SendMessage (hwnd), WM_CHILDACTIVATE, 0, 0
    For Each Frm3 In Forms
        If Frm3.hwnd = hwnd Then
            Set Frm3 = Frm3
            Exit For
        End If
    Next
    If Not Frm3 Is Nothing Then
        Frm3.WindowState = 0
        Frm3.Show
    End If
End Sub

Private Sub CargarBD()
    C_BD = "control"
    Direccion = App.Path & "\BD\Bdcec1.mdb"
End Sub

'***********************************************************************
