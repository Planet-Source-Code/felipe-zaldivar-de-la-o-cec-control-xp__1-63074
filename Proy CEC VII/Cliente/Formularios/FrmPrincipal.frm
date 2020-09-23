VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form FrmPrincipal 
   BackColor       =   &H00FFCEBB&
   BorderStyle     =   0  'None
   ClientHeight    =   11520
   ClientLeft      =   0
   ClientTop       =   -45
   ClientWidth     =   15360
   Icon            =   "FrmPrincipal.frx":0000
   MaxButton       =   0   'False
   ScaleHeight     =   11520
   ScaleWidth      =   15360
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.PictureBox PicContenedor 
      BackColor       =   &H00E89C78&
      BorderStyle     =   0  'None
      Height          =   6735
      Left            =   0
      ScaleHeight     =   6735
      ScaleWidth      =   8055
      TabIndex        =   44
      Top             =   0
      Visible         =   0   'False
      Width           =   8055
      Begin MSComctlLib.ListView LV1 
         Height          =   855
         Left            =   0
         TabIndex        =   26
         Top             =   170
         Width           =   8055
         _ExtentX        =   14208
         _ExtentY        =   1508
         Arrange         =   2
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         OLEDragMode     =   1
         HotTracking     =   -1  'True
         _Version        =   393217
         Icons           =   "ILMenu"
         ColHdrIcons     =   "IL_Usr"
         ForeColor       =   16777215
         BackColor       =   15244408
         Appearance      =   0
         OLEDragMode     =   1
         NumItems        =   0
      End
      Begin TabDlg.SSTab TabOpciones 
         Height          =   5415
         Left            =   180
         TabIndex        =   27
         Top             =   1140
         Width           =   7695
         _ExtentX        =   13573
         _ExtentY        =   9551
         _Version        =   393216
         Tabs            =   6
         TabsPerRow      =   6
         TabHeight       =   520
         BackColor       =   15244408
         TabCaption(0)   =   "Acceso"
         TabPicture(0)   =   "FrmPrincipal.frx":08CA
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Picture1"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Usuarios"
         TabPicture(1)   =   "FrmPrincipal.frx":08E6
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Picture3"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "Procesos"
         TabPicture(2)   =   "FrmPrincipal.frx":0902
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Picture4"
         Tab(2).Control(0).Enabled=   0   'False
         Tab(2).ControlCount=   1
         TabCaption(3)   =   "Restricciones"
         TabPicture(3)   =   "FrmPrincipal.frx":091E
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "Picture9"
         Tab(3).Control(0).Enabled=   0   'False
         Tab(3).ControlCount=   1
         TabCaption(4)   =   "Configuración"
         TabPicture(4)   =   "FrmPrincipal.frx":093A
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "Picture5"
         Tab(4).Control(0).Enabled=   0   'False
         Tab(4).ControlCount=   1
         TabCaption(5)   =   "Acerca de ..."
         TabPicture(5)   =   "FrmPrincipal.frx":0956
         Tab(5).ControlEnabled=   0   'False
         Tab(5).Control(0)=   "SFAbout"
         Tab(5).Control(0).Enabled=   0   'False
         Tab(5).ControlCount=   1
         Begin VB.PictureBox Picture1 
            BackColor       =   &H00E89C78&
            BorderStyle     =   0  'None
            Height          =   4695
            Left            =   240
            ScaleHeight     =   4695
            ScaleWidth      =   7215
            TabIndex        =   65
            Top             =   480
            Width           =   7215
            Begin VB.PictureBox Picture13 
               BackColor       =   &H00D05C28&
               BorderStyle     =   0  'None
               Height          =   1455
               Left            =   240
               ScaleHeight     =   1455
               ScaleWidth      =   6735
               TabIndex        =   71
               Top             =   3120
               Width           =   6735
               Begin VB.TextBox TxtNP 
                  BorderStyle     =   0  'None
                  Height          =   285
                  IMEMode         =   3  'DISABLE
                  Left            =   1560
                  PasswordChar    =   "*"
                  TabIndex        =   8
                  Top             =   1080
                  Width           =   2655
               End
               Begin VB.TextBox TxtNNC 
                  BorderStyle     =   0  'None
                  Height          =   285
                  Left            =   1560
                  TabIndex        =   7
                  Top             =   555
                  Width           =   2655
               End
               Begin VB.CommandButton CmdNU 
                  BackColor       =   &H00D05C28&
                  Caption         =   "&Nuevo Usuario"
                  Height          =   495
                  Left            =   4920
                  TabIndex        =   9
                  Top             =   720
                  Width           =   1215
               End
               Begin VB.TextBox TxtNV 
                  BorderStyle     =   0  'None
                  Height          =   285
                  Left            =   1560
                  TabIndex        =   6
                  Top             =   75
                  Width           =   5055
               End
               Begin VB.Label Label15 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Password:"
                  ForeColor       =   &H00FFFFFF&
                  Height          =   195
                  Left            =   120
                  TabIndex        =   74
                  Top             =   1125
                  Width           =   735
               End
               Begin VB.Label Label14 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Número de cuenta:"
                  ForeColor       =   &H00FFFFFF&
                  Height          =   195
                  Left            =   120
                  TabIndex        =   73
                  Top             =   600
                  Width           =   1365
               End
               Begin VB.Label Label13 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Nombre:"
                  ForeColor       =   &H00FFFFFF&
                  Height          =   195
                  Left            =   120
                  TabIndex        =   72
                  Top             =   120
                  Width           =   600
               End
            End
            Begin VB.PictureBox Picture2 
               BackColor       =   &H00D05C28&
               BorderStyle     =   0  'None
               Height          =   2895
               Left            =   240
               ScaleHeight     =   2895
               ScaleWidth      =   6735
               TabIndex        =   66
               Top             =   120
               Width           =   6735
               Begin VB.CommandButton CmdCerrarSesion 
                  BackColor       =   &H00D05C28&
                  Caption         =   "&Cerrar Sesión"
                  Height          =   375
                  Left            =   4200
                  TabIndex        =   5
                  Top             =   2280
                  Width           =   1335
               End
               Begin VB.TextBox TxtIP 
                  BorderStyle     =   0  'None
                  Height          =   405
                  Left            =   2760
                  TabIndex        =   3
                  Text            =   "TxtIP"
                  Top             =   1680
                  Width           =   2775
               End
               Begin VB.CommandButton CmdConectar 
                  BackColor       =   &H00D05C28&
                  Caption         =   "&Accesar"
                  Height          =   375
                  Left            =   2760
                  TabIndex        =   4
                  Top             =   2280
                  Width           =   1335
               End
               Begin VB.TextBox TxtUsuario 
                  BorderStyle     =   0  'None
                  Height          =   375
                  Left            =   2760
                  TabIndex        =   0
                  Text            =   "TxtUsuario"
                  Top             =   240
                  Width           =   2775
               End
               Begin VB.TextBox TxtPassword 
                  BorderStyle     =   0  'None
                  Height          =   375
                  IMEMode         =   3  'DISABLE
                  Left            =   2760
                  PasswordChar    =   "*"
                  TabIndex        =   1
                  Text            =   "TxtPassword"
                  Top             =   720
                  Width           =   2775
               End
               Begin VB.TextBox TxtMaquina 
                  BorderStyle     =   0  'None
                  Height          =   375
                  Left            =   2760
                  TabIndex        =   2
                  Text            =   "TxtMaquina"
                  Top             =   1200
                  Width           =   2775
               End
               Begin VB.Label Label4 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Servidor:"
                  ForeColor       =   &H00FFFFFF&
                  Height          =   195
                  Left            =   1080
                  TabIndex        =   70
                  Top             =   1800
                  Width           =   630
               End
               Begin VB.Label Label3 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Máquina:"
                  ForeColor       =   &H00FFFFFF&
                  Height          =   195
                  Left            =   1080
                  TabIndex        =   69
                  Top             =   1320
                  Width           =   660
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Número de Cuenta:"
                  ForeColor       =   &H00FFFFFF&
                  Height          =   195
                  Left            =   1080
                  TabIndex        =   68
                  Top             =   360
                  Width           =   1380
               End
               Begin VB.Label Label2 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Password:"
                  ForeColor       =   &H00FFFFFF&
                  Height          =   195
                  Left            =   1080
                  TabIndex        =   67
                  Top             =   840
                  Width           =   735
               End
            End
         End
         Begin VB.PictureBox Picture3 
            BackColor       =   &H00E89C78&
            BorderStyle     =   0  'None
            Height          =   4695
            Left            =   -74760
            ScaleHeight     =   4695
            ScaleWidth      =   7215
            TabIndex        =   64
            Top             =   480
            Width           =   7215
            Begin MSComctlLib.TreeView TVUsuarios 
               Height          =   4335
               Left            =   165
               TabIndex        =   10
               Top             =   165
               Width           =   6855
               _ExtentX        =   12091
               _ExtentY        =   7646
               _Version        =   393217
               Indentation     =   0
               LabelEdit       =   1
               LineStyle       =   1
               Style           =   7
               FullRowSelect   =   -1  'True
               SingleSel       =   -1  'True
               ImageList       =   "IL_Usr"
               Appearance      =   0
            End
         End
         Begin VB.PictureBox Picture4 
            BackColor       =   &H00E89C78&
            BorderStyle     =   0  'None
            Height          =   4695
            Left            =   -74760
            ScaleHeight     =   4695
            ScaleWidth      =   7215
            TabIndex        =   63
            Top             =   480
            Width           =   7215
            Begin VB.ListBox LstProcesos 
               Appearance      =   0  'Flat
               Height          =   4320
               Left            =   165
               TabIndex        =   11
               Top             =   165
               Width           =   6855
            End
         End
         Begin VB.PictureBox Picture5 
            BackColor       =   &H00E89C78&
            BorderStyle     =   0  'None
            Height          =   4695
            Left            =   -74760
            ScaleHeight     =   4695
            ScaleWidth      =   7215
            TabIndex        =   51
            Top             =   480
            Width           =   7215
            Begin VB.PictureBox Picture6 
               BackColor       =   &H00D05C28&
               BorderStyle     =   0  'None
               Height          =   1815
               Left            =   120
               ScaleHeight     =   1815
               ScaleWidth      =   6975
               TabIndex        =   58
               Top             =   120
               Width           =   6975
               Begin VB.TextBox TxtTSesion 
                  BorderStyle     =   0  'None
                  Height          =   285
                  Left            =   3720
                  Locked          =   -1  'True
                  TabIndex        =   17
                  Text            =   "------------------------------------------------------"
                  Top             =   1440
                  Width           =   2535
               End
               Begin VB.TextBox TxtAAC 
                  BorderStyle     =   0  'None
                  Height          =   285
                  Left            =   3720
                  Locked          =   -1  'True
                  TabIndex        =   16
                  Text            =   "0"
                  Top             =   997
                  Width           =   615
               End
               Begin VB.TextBox TxtMinutos 
                  BorderStyle     =   0  'None
                  Height          =   285
                  Left            =   3720
                  Locked          =   -1  'True
                  TabIndex        =   15
                  Text            =   "0"
                  Top             =   555
                  Width           =   615
               End
               Begin VB.Label Label10 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Tipo de sesión:"
                  ForeColor       =   &H00FFFFFF&
                  Height          =   195
                  Left            =   600
                  TabIndex        =   62
                  Top             =   1485
                  Width           =   1080
               End
               Begin VB.Label Label8 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Amonestaciones Acumuladas:"
                  ForeColor       =   &H00FFFFFF&
                  Height          =   195
                  Left            =   600
                  TabIndex        =   61
                  Top             =   1035
                  Width           =   2130
               End
               Begin VB.Label Label12 
                  Alignment       =   2  'Center
                  BackStyle       =   0  'Transparent
                  Caption         =   "Mis datos"
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
                  Height          =   255
                  Left            =   0
                  TabIndex        =   60
                  Top             =   120
                  Width           =   6735
               End
               Begin VB.Label Label7 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Tiempo  restante para reportar anomalias:"
                  ForeColor       =   &H00FFFFFF&
                  Height          =   195
                  Left            =   600
                  TabIndex        =   59
                  Top             =   600
                  Width           =   2925
               End
            End
            Begin VB.PictureBox Picture7 
               BackColor       =   &H00D05C28&
               BorderStyle     =   0  'None
               Height          =   2535
               Left            =   120
               ScaleHeight     =   2535
               ScaleWidth      =   3495
               TabIndex        =   54
               Top             =   2040
               Width           =   3495
               Begin VB.CheckBox ChkPR 
                  BackColor       =   &H00D05C28&
                  Caption         =   "Bloquear procesos restringidos"
                  ForeColor       =   &H00FFFFFF&
                  Height          =   255
                  Left            =   120
                  TabIndex        =   21
                  Top             =   2160
                  Width           =   2535
               End
               Begin VB.CheckBox ChkAL 
                  BackColor       =   &H00D05C28&
                  Caption         =   "Ayuda en linea"
                  ForeColor       =   &H00FFFFFF&
                  Height          =   255
                  Left            =   120
                  TabIndex        =   20
                  Top             =   1680
                  Width           =   1695
               End
               Begin VB.TextBox TxtReportes 
                  BorderStyle     =   0  'None
                  Height          =   285
                  Left            =   2640
                  Locked          =   -1  'True
                  TabIndex        =   18
                  Text            =   "0"
                  Top             =   720
                  Width           =   615
               End
               Begin VB.TextBox TxtAmonestaciones 
                  BorderStyle     =   0  'None
                  Height          =   285
                  Left            =   2640
                  Locked          =   -1  'True
                  TabIndex        =   19
                  Text            =   "0"
                  Top             =   1200
                  Width           =   615
               End
               Begin VB.Label Label6 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Tiempo   para reportar anomalias:"
                  ForeColor       =   &H00FFFFFF&
                  Height          =   195
                  Left            =   120
                  TabIndex        =   57
                  Top             =   720
                  Width           =   2355
               End
               Begin VB.Label Label5 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Maximo de amonestaciones:"
                  ForeColor       =   &H00FFFFFF&
                  Height          =   195
                  Left            =   120
                  TabIndex        =   56
                  Top             =   1200
                  Width           =   2010
               End
               Begin VB.Label Label11 
                  Alignment       =   2  'Center
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
                  Height          =   255
                  Left            =   120
                  TabIndex        =   55
                  Top             =   240
                  Width           =   3255
               End
            End
            Begin VB.PictureBox Picture8 
               BackColor       =   &H00D05C28&
               BorderStyle     =   0  'None
               Height          =   2535
               Left            =   3720
               ScaleHeight     =   2535
               ScaleWidth      =   3375
               TabIndex        =   52
               Top             =   2040
               Width           =   3375
               Begin VB.CheckBox ChkPP 
                  BackColor       =   &H00D05C28&
                  Caption         =   "(Profesor - Administrador) -  (Profesor - Administrador)"
                  ForeColor       =   &H80000005&
                  Height          =   615
                  Left            =   240
                  MaskColor       =   &H00FFFFFF&
                  TabIndex        =   24
                  Top             =   1680
                  Width           =   2895
               End
               Begin VB.CheckBox ChkAA 
                  BackColor       =   &H00D05C28&
                  Caption         =   "Alumno - Alumno"
                  ForeColor       =   &H80000005&
                  Height          =   255
                  Left            =   240
                  MaskColor       =   &H00FFFFFF&
                  TabIndex        =   22
                  Top             =   720
                  Width           =   1575
               End
               Begin VB.CheckBox ChkAP 
                  BackColor       =   &H00D05C28&
                  Caption         =   "Alumno -  (Profesor - Administrador)"
                  ForeColor       =   &H80000005&
                  Height          =   255
                  Left            =   240
                  MaskColor       =   &H00FFFFFF&
                  TabIndex        =   23
                  Top             =   1260
                  Width           =   2895
               End
               Begin VB.Label Label9 
                  Alignment       =   2  'Center
                  BackStyle       =   0  'Transparent
                  Caption         =   "Conversaciones Permitidas"
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
                  Height          =   255
                  Left            =   240
                  TabIndex        =   53
                  Top             =   240
                  Width           =   2895
               End
            End
         End
         Begin VB.PictureBox Picture9 
            BackColor       =   &H00E89C78&
            BorderStyle     =   0  'None
            Height          =   4695
            Left            =   -74760
            ScaleHeight     =   4695
            ScaleWidth      =   7215
            TabIndex        =   45
            Top             =   480
            Width           =   7215
            Begin VB.PictureBox Picture10 
               BackColor       =   &H00D05C28&
               BorderStyle     =   0  'None
               Height          =   3615
               Left            =   120
               ScaleHeight     =   3615
               ScaleWidth      =   3015
               TabIndex        =   49
               Top             =   120
               Width           =   3015
               Begin VB.ListBox LstRestringidos 
                  Appearance      =   0  'Flat
                  Height          =   2955
                  ItemData        =   "FrmPrincipal.frx":0972
                  Left            =   120
                  List            =   "FrmPrincipal.frx":0979
                  Sorted          =   -1  'True
                  TabIndex        =   12
                  Top             =   120
                  Width           =   2775
               End
               Begin VB.Label LblNPR 
                  Alignment       =   2  'Center
                  BackStyle       =   0  'Transparent
                  Caption         =   "0"
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
                  Height          =   255
                  Left            =   120
                  TabIndex        =   50
                  Top             =   3240
                  Width           =   2775
               End
            End
            Begin VB.PictureBox Picture11 
               BackColor       =   &H00D05C28&
               BorderStyle     =   0  'None
               Height          =   3615
               Left            =   3240
               ScaleHeight     =   3615
               ScaleWidth      =   3855
               TabIndex        =   47
               Top             =   120
               Width           =   3855
               Begin VB.ListBox LstPBloqueados 
                  Appearance      =   0  'Flat
                  Height          =   2955
                  ItemData        =   "FrmPrincipal.frx":098E
                  Left            =   120
                  List            =   "FrmPrincipal.frx":0995
                  Sorted          =   -1  'True
                  TabIndex        =   13
                  Top             =   120
                  Width           =   3615
               End
               Begin VB.Label LblNPB 
                  Alignment       =   2  'Center
                  BackStyle       =   0  'Transparent
                  Caption         =   "0"
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
                  Height          =   255
                  Left            =   120
                  TabIndex        =   48
                  Top             =   3240
                  Width           =   3615
               End
            End
            Begin VB.PictureBox Picture12 
               BackColor       =   &H00D05C28&
               BorderStyle     =   0  'None
               Height          =   735
               Left            =   120
               ScaleHeight     =   735
               ScaleWidth      =   6975
               TabIndex        =   46
               Top             =   3840
               Width           =   6975
               Begin VB.TextBox TxtProcSel 
                  Appearance      =   0  'Flat
                  BorderStyle     =   0  'None
                  Height          =   525
                  Left            =   120
                  Locked          =   -1  'True
                  MultiLine       =   -1  'True
                  TabIndex        =   14
                  Text            =   "FrmPrincipal.frx":09A9
                  Top             =   100
                  Width           =   6735
               End
            End
         End
         Begin ShockwaveFlashObjectsCtl.ShockwaveFlash SFAbout 
            Height          =   4695
            Left            =   -74820
            TabIndex        =   25
            Top             =   495
            Width           =   7335
            _cx             =   12938
            _cy             =   8281
            FlashVars       =   ""
            Movie           =   ""
            Src             =   ""
            WMode           =   "Window"
            Play            =   0   'False
            Loop            =   0   'False
            Quality         =   "High"
            SAlign          =   ""
            Menu            =   0   'False
            Base            =   ""
            AllowScriptAccess=   "always"
            Scale           =   "ExactFit"
            DeviceFont      =   0   'False
            EmbedMovie      =   -1  'True
            BGColor         =   ""
            SWRemote        =   ""
            MovieData       =   ""
            SeamlessTabbing =   -1  'True
            Profile         =   0   'False
            ProfileAddress  =   ""
            ProfilePort     =   0
         End
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "FIFO's"
      Height          =   2895
      Left            =   8040
      TabIndex        =   37
      Top             =   2880
      Visible         =   0   'False
      Width           =   1335
      Begin VB.ListBox Lstconfig 
         Height          =   255
         ItemData        =   "FrmPrincipal.frx":09B6
         Left            =   120
         List            =   "FrmPrincipal.frx":09B8
         TabIndex        =   75
         Top             =   1440
         Width           =   1095
      End
      Begin VB.ListBox LstMensajes 
         Height          =   255
         ItemData        =   "FrmPrincipal.frx":09BA
         Left            =   120
         List            =   "FrmPrincipal.frx":09BC
         TabIndex        =   43
         Top             =   1800
         Width           =   1095
      End
      Begin VB.ListBox LstMsjEspera 
         Height          =   255
         Left            =   120
         TabIndex        =   42
         Top             =   720
         Width           =   1095
      End
      Begin VB.ListBox LstVO 
         Height          =   255
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   41
         Top             =   2520
         Width           =   1095
      End
      Begin VB.ListBox LstVL 
         Height          =   255
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   40
         Top             =   2160
         Width           =   1095
      End
      Begin VB.TextBox TxtUrlLog 
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   39
         Top             =   1080
         Width           =   1095
      End
      Begin VB.ListBox Lstrecibidos 
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Etiquetas"
      Height          =   2895
      Left            =   8040
      TabIndex        =   29
      Top             =   0
      Visible         =   0   'False
      Width           =   1575
      Begin VB.Label LblAdmin 
         AutoSize        =   -1  'True
         Caption         =   "LblAdmin"
         Height          =   195
         Left            =   120
         TabIndex        =   36
         Top             =   360
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.Label LblNivel 
         AutoSize        =   -1  'True
         Caption         =   "LblNivel"
         Height          =   195
         Left            =   120
         TabIndex        =   35
         Top             =   720
         Visible         =   0   'False
         Width           =   570
      End
      Begin VB.Label LblNivelS 
         AutoSize        =   -1  'True
         Caption         =   "LblNivelS"
         Height          =   195
         Left            =   120
         TabIndex        =   34
         Top             =   1080
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.Label LblmiPuerto 
         AutoSize        =   -1  'True
         Caption         =   "lblMipuerto"
         Height          =   195
         Left            =   120
         TabIndex        =   33
         Top             =   1440
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.Label LblTReportes 
         AutoSize        =   -1  'True
         Caption         =   "LblTReportes"
         Height          =   195
         Left            =   120
         TabIndex        =   32
         Top             =   1800
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.Label LblTTrascurrido 
         AutoSize        =   -1  'True
         Caption         =   "LblTTrascurrido"
         Height          =   195
         Left            =   120
         TabIndex        =   31
         Top             =   2160
         Visible         =   0   'False
         Width           =   1110
      End
      Begin VB.Label LblTiempoGastado 
         AutoSize        =   -1  'True
         Caption         =   "LblTiempoGastado"
         Height          =   195
         Left            =   120
         TabIndex        =   30
         Top             =   2520
         Visible         =   0   'False
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Diversos Controles "
      Height          =   1455
      Left            =   0
      TabIndex        =   28
      Top             =   6720
      Visible         =   0   'False
      Width           =   2295
      Begin VB.Timer TmrOnTop 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   1680
         Top             =   960
      End
      Begin VB.Timer TmrUPD 
         Interval        =   1000
         Left            =   510
         Top             =   960
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   1290
         Top             =   960
      End
      Begin VB.Timer TmrMonitor 
         Interval        =   1
         Left            =   900
         Top             =   960
      End
      Begin VB.Timer Tmr_Hora 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   120
         Top             =   960
      End
      Begin MSComctlLib.ImageList IL_Usr 
         Left            =   730
         Top             =   240
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   3
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmPrincipal.frx":09BE
               Key             =   "Profr"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmPrincipal.frx":1698
               Key             =   "Admin"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmPrincipal.frx":2372
               Key             =   "Alumno"
            EndProperty
         EndProperty
      End
      Begin MSWinsockLib.Winsock W_Yo 
         Left            =   1340
         Top             =   315
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSWinsockLib.Winsock WUPD 
         Left            =   1800
         Top             =   315
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSComctlLib.ImageList ILMenu 
         Left            =   120
         Top             =   240
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   8
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmPrincipal.frx":304C
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmPrincipal.frx":3926
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmPrincipal.frx":4600
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmPrincipal.frx":4EDA
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmPrincipal.frx":57B4
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmPrincipal.frx":648E
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmPrincipal.frx":6D68
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmPrincipal.frx":7642
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "FrmPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private hWndOldActive As Long '///que no desaparesca la ventana
Public Usr_Accesado As Boolean '' si el usuario a accesado al sistema
Public CerrarSistema As Boolean
Public ocupado As Boolean
Private Declare Function GetTickCount Lib "kernel32" () As Long 'numero de marcados
'Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long) ' call sleep(200) 'domir a la pc

Dim TempFile As String

Private Type da
    FileToSend As String
    FileName As String
    RemoteIP As String
    FileSize As Double
    SaveAs As String
    Pstatus As Double
    lastamount As Double
End Type

Private Info As da

''///contadores para enviar el snapshot
Dim e As Integer
Dim R As Double
''///

Dim i As Integer ''contador para descargar los usuarios cuando hay un error en el servidor
Dim I3 As Integer ''contador para abrir ventas de PM
Dim i5 As Integer '/// contador para quitar usuarios cuando salen jeje
Dim i4 As Integer '//// contador para mandar mensaje(escribiendo o no escribiendo) a ventanas

Dim usrkey As String '' guarda el key del usuario seleccionado en el arbol esto es para cuando queremos comunicarnos con alguien que esta en linea
Dim Datos As String ''guarda la cadena que se nos ha enviado
Dim Datos2 As String ''guarda lo que esta despues del comodin primario "+"
Dim Posicion1 As Integer ''guarda la posicion del comodin primario "+"
Dim Posicion2 As Integer ''guarda la posicion del comodin secundario":"
''' guardan datos del usuario seleccionado en el arbol  *****111******
Dim Usr_I1 As Integer ''guarda el puerto en el que esta ese usuario
Dim Usr_N1 As Integer ''guarda el nivel jerarquico de ese usuario
Dim Usr_M1 As String '''guarda la maquina del usuario
Dim Usr_C1 As String '''guarda la cuenta del usuario
''''termina ******111******
''' agrega un usario al arbo  *****222******
Dim Usr_I2 As Integer ''agrega el puerto
Dim Usr_N2 As Integer ''agrega el nivel jerarquico
Dim Usr_M2 As String '''agrega la maquina
Dim Usr_C2 As String '''agrega la cuenta
''''termina ******222******
''' guarda datos de cuando alguien quiere conversar con nosotros *****333******
Dim Usr_I3 As Integer ''agrega el puerto
Dim Usr_N3 As Integer ''agrega el nivel jerarquico
Dim Usr_M3 As String '''agrega la maquina
Dim Usr_C3 As String '''agrega la cuenta
''''****/////////////////*********
Dim Usr_I4 As Integer ''agrega el puerto
Dim Usr_N4 As Integer ''agrega el nivel jerarquico
Dim Usr_M4 As String '''agrega la maquina
Dim Usr_C4 As String '''agrega la cuenta
''''termina ******333******

'************ declaraciones para las ventanas
Dim VentanaIE As InternetExplorer
Dim VentanaIE_A As New ShellWindows

Dim CVIE As String
Dim LogArchivo As String
Dim VIE_Estado As String
Dim Titulo_URL As String
Dim Busqueda_URL As Integer
Dim ContenidoURL As String
'************/////////////////////////////////

Dim TiempoTrascurrido As Long '''guarda los minutos para ver si estan dentro del rango de los reportes

'///////////////opciones para el snapshot
Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Type PCURSORINFO
    cbSize As Long
    flags As Long
    hCursor As Long
    ptScreenPos As POINTAPI
End Type
'To grab cursor shape -require at least win98 as per Microsoft documentation...
Private Declare Function GetCursorInfo Lib "user32.dll" (ByRef pci As PCURSORINFO) As Long
'To get a Handle to the cursor
Private Declare Function GetCursor Lib "user32" () As Long
'To draw cursor shape on bitmap
Private Declare Function DrawIcon Lib "user32" (ByVal hDc As Long, ByVal x As Long, ByVal y As Long, ByVal hIcon As Long) As Long

'to get the cursor position
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
'to end a waiting loopp
Dim GotIt As Boolean
'To use the scrollbars

Dim ConectandoP As Boolean
Dim lngVer As Long
Dim lngHor As Long
Const iconSize As Integer = 9

Private Sub ChkAA_Click()
    If ChkAA.Tag = "" Then Exit Sub
    ChkAA.Value = ChkAA.Tag
End Sub

Private Sub ChkAL_Click()
    If ChkAL.Tag = "" Then Exit Sub
    ChkAL.Value = ChkAL.Tag
End Sub

Private Sub ChkAP_Click()
    If ChkAP.Tag = "" Then Exit Sub
    ChkAP.Value = ChkAP.Tag
End Sub

Private Sub ChkPP_Click()
    If ChkPP.Tag = "" Then Exit Sub
    ChkPP.Value = ChkPP.Tag
End Sub

Private Sub ChkPR_Click()
    If ChkPR.Tag = "" Then Exit Sub
    ChkPR.Value = ChkPR.Tag
End Sub

'////////////////fin de opciones para en snapshot
Private Sub CmdCerrarSesion_Click()
    If ConectandoP = True Then Exit Sub
    CerrarSesion 1
End Sub

Public Sub CmdMinimizar_Click()
    ControlBoxOPT Me.hWnd, 4
End Sub

Private Sub ControlBoxOPT(Ventana, Opcion As Integer)
    Select Case Opcion
        Case 0:
            Dim x%
            x% = SendMessage(Ventana, WM_CLOSE, 0, 0)
        Case 1:
            x = ShowWindow(Ventana, SW_SHOW)
        Case 2:
            x = ShowWindow(Ventana, SW_HIDE)
        Case 3:
            x = ShowWindow(Ventana, SW_MAXIMIZE)
        Case 4:
            x = ShowWindow(Ventana, SW_MINIMIZE)
    End Select
End Sub

Private Sub CmdNU_Click()
    Dim CantTonos, segundos As Long
    If TxtIP.Text = "" Then Exit Sub
    W_Yo.Close
    W_Yo.Connect TxtIP.Text, 1257
    CantTonos = GetTickCount
    Do
        If segundos > 5 Then
            MsgBox "No se ha podido realizar la conexión consulta al administrador"
            Exit Sub
        End If
        DoEvents
    Loop Until W_Yo.State = sckConnected
    If W_Yo.State = sckConnected Then
        Call Enviar("NUEVOUSUARIO¯[®©]¤¤¤" & TxtNV.Text & "¤¢©§¦[BOLA]¦§©¢¤" & TxtNNC.Text & "¤¢©§¦[BOLA]¦§©¢¤" & TxtNP.Text)
    End If
End Sub

Private Sub VaciarCampos()
    LstPBloqueados.Clear
    LstRestringidos.Clear
    TxtProcSel.Text = ""
End Sub
''' aqui van las funciones especificas de cada control o formulario

Private Sub Form_Load()
    If App.PrevInstance Then
        ActivatePrevInstance
    End If
    'Call ModeStartUP 'Si se quita el apostrofe se activara el programa para iniciar con el registro
    'para activar la proteccion completa activar el TmrOnTop
    Call ChkManifest
    Call CargarIJDLL
    Call Redondear(Me)
    Call IniciarAbout
    Call CargarMenu
    Call CargarArbol
    CerrarSesion 1
    TxtUsuario.Text = "bola"
    TxtPassword.Text = "bola"
    InitCommonControls
End Sub

Private Sub CargarIJDLL()
    Dim PathDll As String
    PathDll = "C:\WINDOWS\system32\IJL15.DLL"
    If FileExists(PathDll) = True Then Exit Sub
    Dim FileNum     As Integer
    Dim DataArray() As Byte
    DataArray = LoadResData("IJL15", "DLL")
    FileNum = FreeFile
    Open PathDll For Binary As #FileNum
    Put #FileNum, 1, DataArray()
    Close #FileNum
End Sub

Private Sub IniciarAbout()
    SFAbout.Movie = App.Path & "\About.swf"
End Sub

Private Sub ColocarControles()
On Error GoTo Error
    PicContenedor.Visible = False
    LV1.Left = (PicContenedor.Width / 2) - (LV1.Width / 2)
    PicContenedor.Left = ((Me.Width / 2) - (PicContenedor.Width / 2))
    PicContenedor.Top = (Me.Height / 2) - (PicContenedor.Height / 2)
    PicContenedor.Visible = True
    Exit Sub
Error:
    Resume Next
    PicContenedor.Visible = True
End Sub

Private Sub CerrarUDP(Opt As Boolean)
On Error GoTo SinHost

    If Opt = True Then
        WUPD.Close
    Else
        With WUPD
            .Close
            .Protocol = sckUDPProtocol
            .LocalPort = 7200
            .RemotePort = 7201
            .RemoteHost = "255.255.255.255"
            .SendData "bola"
        End With
    End If
    Exit Sub
    
SinHost:
    If Err.Number = 10065 Then
        Call HostOpcional
    End If
    Resume Next
End Sub

Private Sub HostOpcional()
On Error GoTo Error

    With WUPD
        .Close
        .Protocol = sckUDPProtocol
        .LocalPort = 7200
        .RemotePort = 7201
        .RemoteHost = .LocalHostName
        .SendData "bola"
    End With
    Exit Sub
    
Error:
    If Err.Number = 0 Then
        Exit Sub
    End If
    Resume Next
End Sub

Private Sub Form_Terminate()
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unhook Me.hWnd
    Cancel = 1
End Sub

Private Sub LV1_Click()
    If LV1.SelectedItem Is Nothing Then Exit Sub
    Dim OptSWin As New Class_ShutDown
    
    If LV1.SelectedItem.Key = "Ayuda" Then
    ElseIf LV1.SelectedItem.Key = "Ayuda Directa" Then
        If Conectado = True Then
            If ChkAL.Value = 1 Then
                If ExisteVentanaII(Admin) = False Then
                    Call NuevaVentana(0, 3, Admin, Mi_Puerto, Usr_Nivel, TxtUsuario.Text, "°°°Conversación aceptada°°°", Usr_NivelS, "Administrador")
                End If
            End If
        End If
    ElseIf LV1.SelectedItem.Key = "Reportes" Then
        If Conectado = True Then
            If R_Activo = True Then
                If Max_ReportesY > 5 Then
                    MsgBox "No Puedes enviar mas de 5 reportes!!!" & Chr(10) & "Consulta a tu Administrador!!!", vbDefaultButton1, "Atención!!!"
                Else
                    FrmReporte.Maquina_Reporte = TxtMaquina.Text
                    FrmReporte.Quien_Reporta = TxtUsuario.Text
                    FrmReporte.Show
                End If
            End If
        End If
    ElseIf LV1.SelectedItem.Key = "Cerrar Sesion" Then
        OptSWin.ExitWindows WE_LOGOFF
        Unhook Me.hWnd
        Unload Me
        Unhook Me.hWnd
        End
    ElseIf LV1.SelectedItem.Key = "Reiniciar" Then
        OptSWin.ExitWindows WE_REBOOT
        Unhook Me.hWnd
        Unload Me
        Unhook Me.hWnd
        End
    ElseIf LV1.SelectedItem.Key = "Apagar" Then
        OptSWin.ExitWindows WE_SHUTDOWN
        Unload Me
        Unhook Me.hWnd
        End
    ElseIf LV1.SelectedItem.Key = "Apagar ATX" Then
        OptSWin.ExitWindows WE_POWEROFF
        Unhook Me.hWnd
        Unload Me
        Unhook Me.hWnd
        End
    End If
End Sub

Private Sub TmrOnTop_Timer()
    MaxTop Me.hWnd, True
End Sub

Private Sub TmrUPD_Timer()
On Error GoTo SinHost
    WUPD.SendData "ID¯[®©]¤¤¤[SOLICITUD]"
    Exit Sub
SinHost:
    If Err.Number = 10065 Then
        Call HostOpcional
    End If
    Resume Next
End Sub

Private Sub CmdConectar_Click()
    On Error GoTo Error
    Dim CantTonos As Long 'guarda la cantidad de marcaciones
    Dim segundos As Long 'guarda los segundos fuera
    If Conectado = True Then
        W_Yo.Close
    End If
    LstProcesos.Clear
    If W_Yo.State = sckClosed Then
        W_Yo.Connect TxtIP.Text, 1257
        CantTonos = GetTickCount
        CerrarSesion 0
        ConectandoP = True
        Do
            segundos = Round((GetTickCount - CantTonos) / 1000)
            Me.Caption = "::: Control de Acceso al CEC1 ::: | Estado: Verificando Número de Cuenta y Password"
            If segundos > 5 Then
                MsgBox "No se ha podido realizar la conexión consulta al administrador"
                ConectandoP = False
                CerrarSesion 1
                TxtUsuario.SetFocus
                DoEvents
                Exit Sub
            End If
            DoEvents
        Loop Until W_Yo.State = sckConnected
        If W_Yo.State = sckConnected Then
            '////// Si accesamos a la terminal
            '////// Mandamos nuestra clave de Acceso
            Call Enviar("LOGIN¯[®©]¤¤¤" & TxtUsuario.Text & "¤¢©§¦[BOLA]¦§©¢¤" & TxtPassword.Text & "¤¢©§¦[BOLA]¦§©¢¤" & TxtMaquina.Text)
            ConectandoP = False
        End If
    End If
    Exit Sub

Error:
    If Err.Number > 0 Then
        MsgBox Err.Number & " - " & Err.Description, , "Error Inesperado !!!"
    End If
    Resume Next
End Sub

Private Sub LstPBloqueados_Click()
    TxtProcSel.Text = LstPBloqueados.List(LstPBloqueados.ListIndex)
End Sub

Private Sub LstRestringidos_Click()
    TxtProcSel.Text = LstRestringidos.List(LstRestringidos.ListIndex)
End Sub

Private Sub Timer1_Timer()
    Dim av As Integer
    av = av + 1
    LblTTrascurrido = av
    If av >= Tmr_Hora.Interval Then Timer1.Enabled = False
End Sub

Private Sub Tmr_Hora_Timer()
    Dim resto As Double
    Dim MR As String
    
    If CLng(LblTTrascurrido.Caption) > CLng(TiempoTrascurrido) Then
        TiempoTrascurrido = TiempoTrascurrido + 1
        LblTiempoGastado.Caption = TiempoTrascurrido
        resto = (Val(LblTTrascurrido.Caption) - Val(TiempoTrascurrido)) / 60
        If InStr(1, resto, ".") Then
            MR = Left(resto, InStr(1, resto, ".") - 1)
        Else
            MR = resto
        End If
        TxtMinutos.Text = MR
        If FrmReporte.Visible = True Then
            FrmReporte.Caption = MR & " Minuto(s) Restante(s) para enviar Reportes"
        End If
    Else
        TiempoTrascurrido = 0
        Tmr_Hora.Enabled = False
        R_Activo = False
        If FrmReporte.Visible = True Then Unload FrmReporte
        If Conectado = False Then Exit Sub
        MsgBox "Ha trascurrido el tiempo para reportar anomalias en la maquina!!!" _
        & Chr(10) & "Tiempo: " & LblTReportes.Caption & " Minuto(s).", vbDefaultButton1, "Atención!!!"
    End If
End Sub

Private Sub TmrMonitor_Timer()
    Dim MsjIndex As Long
    MsjIndex = 0

    Do While ocupado = False And MsjIndex < LstMsjEspera.ListCount
        If Conectado = False Then Exit Sub
        W_Yo.SendData (LstMsjEspera.List(MsjIndex) & ("¤¦[MYF]¦¤" & vbCrLf & "¤¦[MYF]¦¤"))
        LstMsjEspera.RemoveItem (MsjIndex)
        ocupado = True
        MsjIndex = MsjIndex + 1
        DoEvents
    Loop
End Sub


'LogArchivo = App.Path & "\" & Replace(Date & "_" & "UrlLog", "/", "_") & ".txt"
'Open LogArchivo For Append As #1
'Print #1, Date & " | " & Time; " | " & Titulo_URL
'Close #1
'*******//
'Called by the module anytime another window gains focus.  Since we only
'  know the hwnd of the new window, we need to keep track of the last
'  window to keep focus was (and don't assume that '** ' was put in the
'  caption by us (a window can have '** ' in it's caption))
Public Sub WindowActivated(hWnd As Long)
On Error Resume Next
    Dim i As Integer
    'First off, go through and find the old active window, and remove
    '  the '** ' from the front of the title
    For i = 0 To LstProcesos.ListCount - 1
        If LstProcesos.ItemData(i) = hWndOldActive Then
            If Mid(LstProcesos.List(i), 1, 3) = "** " Then
                LstProcesos.List(i) = Mid(LstProcesos.List(i), 4)
                End If
            Exit For
        End If
    Next
    'Then find the window that was activated, and put '** ' in front of the
    '  caption
    For i = 0 To LstProcesos.ListCount - 1
        If LstProcesos.ItemData(i) = hWnd Then
            LstProcesos.List(i) = "** " & LstProcesos.List(i)
            Exit For
        End If
    Next
    'Finally, set our variable of the active hwnd
    hWndOldActive = hWnd
End Sub

'Called by the module whenever a window caption is changed (or atleast,
'  believed to be changed)
Public Sub WindowRedraw(hWnd As Long)
On Error Resume Next
    Dim strCaption As String
    Dim i As Integer
    strCaption = String(255, " ")
  
    GetWindowText hWnd, strCaption, 254  'Grab the new caption, find the
                                       '  spot in the listbox, and put
                                        '  it in.
    For i = 0 To LstProcesos.ListCount - 1
        If LstProcesos.ItemData(i) = hWnd Then
            LstProcesos.List(i) = strCaption
            LstProcesos.ListIndex = i
            Exit For
        End If
    Next
End Sub

Private Function AnalizarProceso(hWnd As Long, Caption As String) As String
    Dim IPro As Integer
    Dim TPro As String
    Dim Pos As Integer
    AnalizarProceso = ""
    For IPro = 0 To LstRestringidos.ListCount - 1
        TPro = LCase(LstRestringidos.List(IPro))
        Pos = InStr(1, LCase(Caption), LCase(TPro))
        If Pos > 0 Then
            LstPBloqueados.AddItem Caption
            LblNPB.Caption = LstPBloqueados.ListCount
            TerminarProceso , hWnd
            AnalizarProceso = "[BLOQUEADO] - "
            Call Enviar("AMONESTACIONUSR¯[®©]¤¤¤" & TxtUsuario.Text)
            MsgBox "El siguiente proceso ha sido bloqueado: " & Caption, , "Atención!!!": DoEvents
            Exit For
        End If
    Next
End Function

'Called by the module whenever a window is created
Public Sub WindowCreated(hWnd As Long)
    Dim i As Integer
    Dim lExStyle    As Long
    Dim bNoOwner    As Boolean
    Dim lreturn     As Long
    Dim sWindowText As String
    Dim Bloq As String
    Bloq = ""
    If Not hWnd = Me.hWnd Then
        If IsWindowVisible(hWnd) Then
            If GetParent(hWnd) = 0 Then
                bNoOwner = (GetWindow(hWnd, GW_OWNER) = 0)
                lExStyle = GetWindowLong(hWnd, GWL_EXSTYLE)
                If (((lExStyle And WS_EX_TOOLWINDOW) = 0) And bNoOwner) Or _
                    ((lExStyle And WS_EX_APPWINDOW) And Not bNoOwner) Then
                    sWindowText = Space$(1024)
                    lreturn = GetWindowText(hWnd, sWindowText, Len(sWindowText))
                    If lreturn Then
                        sWindowText = Left$(sWindowText, lreturn)
                        LstProcesos.AddItem sWindowText
                        LstProcesos.ItemData(LstProcesos.NewIndex) = hWnd
                        Bloq = AnalizarProceso(hWnd, sWindowText)
                        Call Enviar("PROCESORA¯[®©]¤¤¤" & TxtUsuario.Text & "¤¢©§¦[BOLA]¦§©¢¤" & TxtMaquina.Text & "¤¢©§¦[BOLA]¦§©¢¤" & Bloq & sWindowText)
                    End If
                End If
            End If
        End If
    End If
  
End Sub

'Called by the module whenever a window is destroyed
Public Sub WindowDestroyed(hWnd As Long)
On Error Resume Next
    Dim i As Integer
    For i = 0 To LstProcesos.ListCount - 1 'Loop around looking for the hwnd and
                                   '  remove it from the list
        If LstProcesos.ItemData(i) = hWnd Then
            LstProcesos.RemoveItem i
            Exit For
        End If
    Next
    Exit Sub
End Sub
'*******//

Private Sub W_Yo_DataArrival(ByVal bytesTotal As Long)
    Dim TxtDigerido As String
    Dim Dat As String
    W_Yo.GetData Datos, vbString
    If InStr(1, Datos, ("¤¦[MYF]¦¤" & vbCrLf & "¤¦[MYF]¦¤")) = 0 Then
        Dat = Datos
        If LCase(Mid(Dat, 1, 2)) = "ok" Then
            Dim temparray2() As String
            Dim fname2 As String
            Dim fsize2 As Double
            temparray2 = Split(Dat, "|")
            fname2 = temparray2(1)
            fsize2 = temparray2(2)
            If fname2 <> getfilename(Info.FileToSend) Or fsize2 <> Info.FileSize Then Exit Sub
            Close #1
            Open Info.FileToSend For Binary Access Read As #1
                If LOF(1) = 0 Then Exit Sub
                Dim SendBuffer As String
                SendBuffer = Space$(LOF(1))
                Get #1, , SendBuffer
            Close #1
            W_Yo.SendData SendBuffer & "******FINAL******"
            e = 2
            R = Timer
            Do Until Timer > R + 2
            DoEvents
            Loop
            Exit Sub
        End If
        Put #2, , Dat
    Else
        Do While InStr(1, Datos, ("¤¦[MYF]¦¤" & vbCrLf & "¤¦[MYF]¦¤"))
            Lstrecibidos.AddItem Mid(Datos, 1, InStr(1, Datos, ("¤¦[MYF]¦¤" & vbCrLf & "¤¦[MYF]¦¤")) - 1)
            Datos = Mid(Datos, InStr(1, Datos, ("¤¦[MYF]¦¤" & vbCrLf & "¤¦[MYF]¦¤")) + Len("¤¦[MYF]¦¤" & vbCrLf & "¤¦[MYF]¦¤"), Len(Datos))
        Loop
    
        If Len(Lstrecibidos.List(0)) = 0 Then
            Lstrecibidos.RemoveItem (0)
        End If
        
        If W_Yo.Tag = "USR_ACCESADO" Then
            Usr_Accesado = True
        Else
            Usr_Accesado = False
        End If
        
        Do While Lstrecibidos.ListCount > 0
            AnalizarDatos Usr_Accesado
        Loop
    End If
End Sub

Private Sub W_Yo_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
''''''aqui va el codigo para apagar la pc
    ServicioCerrado
End Sub

Private Sub W_Yo_Close()
    ServicioCerrado
End Sub

Private Sub ServicioCerrado()
    Dim ContSalida1 As Integer
    Unhook Me.hWnd
    W_Yo.Tag = ""
    W_Yo.Close
    If Admin <> "" Then
        Call InicioSesion(CStr(Admin & Chr(10) & "Finalizo Sesión"))
    End If
    For ContSalida1 = 0 To LstVO.ListCount - 1
        If PMensajeC(LstVO.List(ContSalida1)).Tag = Admin Then
            Unload PMensajeC(LstVO.List(ContSalida1))
            Exit For
        End If
    Next
    For i = 4 To TVUsuarios.Nodes.Count '////removemos todos los usuarios
        If i > TVUsuarios.Nodes.Count Then Exit Sub
        TVUsuarios.Nodes.Remove (i)
        i = i - 1
    Next i
    Tmr_Hora.Enabled = False
    CerrarSesion 1
End Sub

Private Sub W_Yo_ConnectionRequest(ByVal requestID As Long)
    W_Yo.Accept requestID
End Sub

Private Sub TB_Usr_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case Is = "Ayuda"
            MsgBox Button.Key
        Case Is = "AyudaR"
        
        Case Is = "Reportes"
        
        Case Is = "Cerrar"
            MsgBox Button.Key
        Case Is = "Reiniciar"
            MsgBox Button.Key
        Case Is = "Apagar"
            Unload Me
            Unhook Me.hWnd
            End
    End Select
End Sub

Private Sub TVUsuarios_DblClick()
    usrkey = TVUsuarios.SelectedItem.Key
    If TVUsuarios.SelectedItem.Text = TxtUsuario Then Exit Sub
    If usrkey = "Admin" Or usrkey = "Profr" Or usrkey = "Alumno" Then Exit Sub
    Posicion2 = InStr(1, usrkey, ":")
    Usr_M1 = Mid(usrkey, 1, Posicion2 - 1)
    usrkey = Mid(usrkey, Posicion2 + 1)
    Posicion2 = InStr(1, usrkey, ":")
    Usr_I1 = Mid(usrkey, 1, Posicion2 - 1)
    Usr_N1 = Mid(usrkey, Posicion2 + 1, 1)
    usrkey = Mid(usrkey, Posicion2 + 1)
    Posicion2 = InStr(1, usrkey, ":")
    Usr_C1 = Mid(usrkey, Posicion2 + 1)

    If ValidarConversacion(Usr_N1, Usr_Nivel) = True Then
        If ExisteVentanaII(Usr_C1) = False Then
            Call NuevaVentana(Usr_I1, Usr_N1, Usr_C1, Mi_Puerto, Usr_Nivel, TxtUsuario.Text, "°°°Conversación aceptada°°°", Usr_NivelS, ParseLevel(Usr_N1))
        End If
    End If
End Sub

Private Function ValidarConversacion(USolicitado As Integer, USolicitante As Integer) As Boolean
ValidarConversacion = True
    If USolicitante = 1 Then
        If USolicitado = 1 Then
            If ChkAA.Value = 1 Then
                ValidarConversacion = True
                Exit Function
            Else
                MsgBox "La conversación Alumno - Alumno no esta permitida!!!", , "Consulta a tu administrador!!!"
                ValidarConversacion = False
                Exit Function
            End If
        ElseIf USolicitado = 2 Then
            If ChkAP.Value = 1 Then
                ValidarConversacion = True
                Exit Function
            Else
                MsgBox "La conversación Alumno - Profesor no esta permitida!!!", , "Consulta a tu administrador!!!"
                ValidarConversacion = False
                Exit Function
            End If
        ElseIf USolicitado = 3 Then
            If ChkAP.Value = 1 Then
                ValidarConversacion = True
                Exit Function
            Else
                MsgBox "La conversación Alumno - Administrador no esta permitida!!!", , "Consulta a tu administrador!!!"
                ValidarConversacion = False
                Exit Function
            End If
        End If
    ElseIf USolicitante = 2 Then
        If USolicitado = 1 Then
            If ChkAP.Value = 1 Then
                ValidarConversacion = True
                Exit Function
            Else
                MsgBox "La conversación Profesor - Alumno no esta permitida!!!", , "Consulta a tu administrador!!!"
                ValidarConversacion = False
                Exit Function
            End If
        ElseIf USolicitado = 2 Then
            If ChkPP.Value = 1 Then
                ValidarConversacion = True
                Exit Function
            Else
                MsgBox "La conversación Profesor - Profesor no esta permitida!!!", , "Consulta a tu administrador!!!"
                ValidarConversacion = False
                Exit Function
            End If
        ElseIf USolicitado = 3 Then
            If ChkPP.Value = 1 Then
                ValidarConversacion = True
                Exit Function
            Else
                MsgBox "La conversación Profesor - Administrador no esta permitida!!!", , "Consulta a tu administrador!!!"
                ValidarConversacion = False
                Exit Function
            End If
        End If
    ElseIf USolicitante = 3 Then
        If USolicitado = 1 Then
            If ChkAP.Value = 1 Then
                ValidarConversacion = True
                Exit Function
            Else
                MsgBox "La conversación Administrador - Alumno no esta permitida!!!", , "Consulta a tu administrador!!!"
                ValidarConversacion = False
                Exit Function
            End If
        ElseIf USolicitado = 2 Then
            If ChkPP.Value = 1 Then
                ValidarConversacion = True
                Exit Function
            Else
                MsgBox "La conversación Administrador - Profesor no esta permitida!!!", , "Consulta a tu administrador!!!"
                ValidarConversacion = False
                Exit Function
            End If
        ElseIf USolicitado = 3 Then
            If ChkAP.Value = 1 Then
                ValidarConversacion = True
                Exit Function
            Else
                MsgBox "La conversación Administrador - Administrador no esta permitida!!!", , "Consulta a tu administrador!!!"
                ValidarConversacion = False
                Exit Function
            End If
        End If
    End If
End Function
'//////aqui van todas la funciones inpropias

Private Sub CargarArbol()
    ''''cargamos los nodos raiz de nuestro arbol
    TVUsuarios.Nodes.Add , , "Admin", "Administradores", "Admin"
    TVUsuarios.Nodes.Add , , "Profr", "Profesores", "Profr"
    TVUsuarios.Nodes.Add , , "Alumno", "Alumnos", "Alumno"
End Sub

Public Sub Enviar(Texto As String)
    '///con esto enviamos datos as servidor
    'MsgBox Texto
    If Texto = "" Then Exit Sub
    LstMsjEspera.AddItem Texto
End Sub

Public Function Conectado() As Boolean
    '///aqui checamos is estamos conectados
    If W_Yo.State = sckConnected Then
        Conectado = True
    Else
        Conectado = False
    End If
End Function

Private Sub PermisosUsr(NivelUsr As Integer)
''////////Establecemos los permisos para crear un usuario
If NivelUsr = 1 Then
    Usr_NivelS = "Alumno"
ElseIf NivelUsr = 2 Then
    Usr_NivelS = "Profesor"
ElseIf NivelUsr = 3 Then
    Usr_NivelS = "Administrador"
End If
    TxtTSesion.Text = Usr_NivelS
    LblNivelS = Usr_NivelS
    LstMsjEspera.Clear
End Sub

Private Sub AnalizarDatos(AccesadoC As Boolean)
Dim Datos3 As String ''//guarda datos secundarios de la cadena datos2
Dim Titulo As String ''titulo del msgbox
Dim I2 As Integer ''contador del bucle para agregar usuarios al arbol exepto a nosotros
Dim MBMensaje As String ''guarda mensajes que nos envian y que se van a mostrar en msgbox
Dim TxtDigerido As String
Dim OpcionC As String
Dim DatosC As String
Dim Posicion1 As Integer

MBMensaje = ""
Titulo = ""
Datos3 = ""
I2 = 0

        TxtDigerido = Lstrecibidos.List(0)
        Lstrecibidos.RemoveItem (0)
        Posicion1 = InStr(1, TxtDigerido, "¯[®©]¤¤¤")
        OpcionC = Mid$(TxtDigerido, 1, Posicion1 - 1)
        DatosC = Mid$(TxtDigerido, Posicion1 + 8)
    
        If OpcionC = "NOACCESO" Then
            MBMensaje = Split(DatosC, "¤¢©§¦[BOLA]¦§©¢¤")(0)
            Titulo = Split(DatosC, "¤¢©§¦[BOLA]¦§©¢¤")(1)
            MsgBox MBMensaje, vbDefaultButton1, Titulo
            CmdConectar.Enabled = True
            CerrarSesion 1
            Exit Sub
        End If

        If OpcionC = "ACCESO1" Then  ''' si el usuario pasa todas las condiciones
            If AccesadoC = False Then
                T_Reportes = 0
                Posicion2 = InStr(1, DatosC, "¤¢©§¦[BOLA]¦§©¢¤")
                MBMensaje = Mid$(DatosC, 1, Posicion2 - 1)
                Datos3 = Mid(DatosC, Posicion2 + 16)
                Posicion2 = InStr(1, Datos3, "¤¢©§¦[BOLA]¦§©¢¤")
                Titulo = Mid(Datos3, 1, Posicion2 - 1)
                Datos3 = Mid(Datos3, Posicion2 + 16)
                Posicion2 = InStr(1, Datos3, "¤¢©§¦[BOLA]¦§©¢¤")
                Admin = Mid(Datos3, 1, Posicion2 - 1)
                Datos3 = Mid(Datos3, Posicion2 + 16)
                Posicion2 = InStr(1, Datos3, "¤¢©§¦[BOLA]¦§©¢¤")
                Usr_Nivel = Mid(Datos3, 1, Posicion2 - 1)
                Datos3 = Mid(Datos3, Posicion2 + 16)
                Posicion2 = InStr(1, Datos3, "¤¢©§¦[BOLA]¦§©¢¤")
                Mi_Puerto = Mid(Datos3, 1, Posicion2 - 1)
                Datos3 = Mid(Datos3, Posicion2 + 16)
                Posicion2 = InStr(1, Datos3, "¤¢©§¦[BOLA]¦§©¢¤")
                T_Reportes = Mid(Datos3, 1, Posicion2 - 1)
                Hora_Yo = Mid(Datos3, Posicion2 + 16)
                W_Yo.Tag = "USR_LOGEADO"
                Call PermisosUsr(Usr_Nivel)
                LblTReportes.Caption = T_Reportes
                LblAdmin.Caption = Admin
                LblNivel.Caption = Usr_Nivel
                LblmiPuerto.Caption = Mi_Puerto
                LblTTrascurrido.Caption = T_Reportes * 60
                R_Activo = True
                Tmr_Hora.Interval = 1000
                Tmr_Hora.Enabled = True
                Posicion2 = 0
                Datos3 = ""
                MBMensaje = ""
                Titulo = ""
                CerrarSesion 0
                Me.Caption = "::: Control de Acceso al CEC1 ::: | Estado: Conectado al sistema"
                Call InicioSesion(CStr(TxtUsuario.Text & Chr(10) & "Bienvenido al Sistema !!!"))
                Exit Sub
            Else
                Exit Sub
            End If
            Exit Sub
        End If
        
        If OpcionC = "NUEVOUSUARIOCREADO" Then
            If DatosC = "SI" Then
                TxtUsuario.Text = TxtNNC.Text
                TxtPassword.Text = TxtNP.Text
                MsgBox "Tu cuenta ha sido actualizada satisfactoriamente!!!", , "Atención!!!"
                CmdConectar.SetFocus
            ElseIf DatosC = "CLAVEREPETIDA" Then
                MsgBox "Esta cuenta de usuario ya existe. Inténtalo de nuevo con otra!!!", , "Atención!!!"
            ElseIf DatosC = "YAREGISTRADO" Then
                MsgBox "Tu cuenta ya ha sido registrada, consulta al administrador", , "Atención!!!"
            ElseIf DatosC = "NOENCONTRADO" Then
                MsgBox "Tus datos no estan dados de alta, consulta al administrador", , "Atención!!!"
            End If
            Exit Sub
        End If
        
        If OpcionC = "PROCESOSR" Then 'RECEPCION DE LOS PROCESOS RESTRINGIDOS
            Dim Pos As Integer
            Dim PCad As String
            LstRestringidos.Clear
            Do While InStr(1, DatosC, ("¤¢©§¦[BOLA]¦§©¢¤"))
                Pos = InStr(1, DatosC, ("¤¢©§¦[BOLA]¦§©¢¤"))
                PCad = Mid(DatosC, 1, Pos - 1)
                DatosC = Mid(DatosC, Pos + 16)
                LstRestringidos.AddItem PCad
                LblNPR.Caption = LstRestringidos.ListCount
            Loop
            StartHook Me.hWnd
            Exit Sub
        End If
        
        If OpcionC = "CONFIGURACION" Then 'RECEPCION DE LOS PROCESOS RESTRINGIDOS
            Dim Pos2 As Integer
            Dim PCad2 As String
            Lstconfig.Clear
            Do While InStr(1, DatosC, ("¤¢©§¦[BOLA]¦§©¢¤"))
                Pos2 = InStr(1, DatosC, ("¤¢©§¦[BOLA]¦§©¢¤"))
                PCad2 = Mid(DatosC, 1, Pos2 - 1)
                DatosC = Mid(DatosC, Pos2 + 16)
                If PCad2 = "Verdadero" Or PCad2 = "True" Then
                    PCad2 = "1"
                ElseIf PCad2 = "Falso" Or PCad2 = "False" Then
                    PCad2 = "0"
                End If
                Lstconfig.AddItem PCad2
                DoEvents
            Loop
            LLenarConfiguracion
            Exit Sub
        End If
        
        If OpcionC = "BLOQUEARUSUARIO" Then
            MsgBox "Tu cuenta ha sido bloqueada, la maquina se apagará en 2 minutos, guarda todos tus archivos...", , "Atención!!!": DoEvents
            Exit Sub
        End If
        
        If OpcionC = "AGREGAR" Then ''''////cuando entra un usuario o si ya hay usuarioss
            ''en linea los agregamos a nuestro arbol
            Posicion2 = InStr(1, DatosC, "¤¢©§¦[BOLA]¦§©¢¤")
            Usr_M2 = Mid(DatosC, 1, Posicion2 - 1)
            Datos3 = Mid(DatosC, Posicion2 + 16)
            Posicion2 = InStr(1, Datos3, "¤¢©§¦[BOLA]¦§©¢¤")
            Usr_I2 = Mid(Datos3, 1, Posicion2 - 1)
            Usr_N2 = Mid(Datos3, Posicion2 + 16, 1)
            Datos3 = Mid(Datos3, Posicion2 + 16)
            Posicion2 = InStr(1, Datos3, "¤¢©§¦[BOLA]¦§©¢¤")
            Usr_C2 = Mid(Datos3, Posicion2 + 16)
        
            For I2 = 1 To TVUsuarios.Nodes.Count
                If TVUsuarios.Nodes(I2).Text = Usr_C2 Then
                    Exit Sub
                End If
            Next I2
            
            If Usr_N2 = 1 Then
                TVUsuarios.Nodes.Add "Alumno", tvwChild, Usr_M2 & "¤¢©§¦[BOLA]¦§©¢¤" & Usr_I2 & "¤¢©§¦[BOLA]¦§©¢¤" & Usr_N2 & "¤¢©§¦[BOLA]¦§©¢¤" & Usr_C2, Usr_C2, "Alumno"
            ElseIf Usr_N2 = 2 Then
                TVUsuarios.Nodes.Add "Profr", tvwChild, Usr_M2 & "¤¢©§¦[BOLA]¦§©¢¤" & Usr_I2 & "¤¢©§¦[BOLA]¦§©¢¤" & Usr_N2 & "¤¢©§¦[BOLA]¦§©¢¤" & Usr_C2, Usr_C2, "Profr"
            ElseIf Usr_N2 = 3 Then
                TVUsuarios.Nodes.Add "Admin", tvwChild, Usr_M2 & "¤¢©§¦[BOLA]¦§©¢¤" & Usr_I2 & "¤¢©§¦[BOLA]¦§©¢¤" & Usr_N2 & "¤¢©§¦[BOLA]¦§©¢¤" & Usr_C2, Usr_C2, "Admin"
            End If
            'On Error Resume Next
            Call InicioSesion(CStr(Usr_C2 & Chr(10) & "acaba de Iniciar Sesión"))
            Exit Sub
        End If
        
        If OpcionC = "FECHAR" Then
            If Not IsDate(DatosC) Then Exit Sub
            FechaEntrada = CDate(DatosC)
            Exit Sub
        End If
        
        If OpcionC = "AYUDAR" Then
            LstMensajes.AddItem DatosC
            Call MandarMensaje
            Exit Sub
        End If
        
        If OpcionC = "ESC" Then
            
            Dim ESCCuenta As String
            Dim ESCOpcion As Integer
            Dim ESCContador As Integer
            Dim ESCPuerto As Integer
            
            If LstVO.ListCount = 0 Then Exit Sub
            
            Posicion2 = InStr(1, DatosC, "¯-_[††]_-¯")
            ESCCuenta = Mid(DatosC, 1, Posicion2 - 1)
            ESCPuerto = Mid(DatosC, Posicion2 + 10)
            ESCOpcion = Right(ESCCuenta, 1)
            ESCCuenta = Mid(ESCCuenta, 1, Len(ESCCuenta) - 1)
            
            'MsgBox ESCCuenta & Chr(10) & ESCPuerto & Chr(10) & ESCOpcion
            
            For ESCContador = 0 To LstVO.ListCount - 1
                If PMensajeC(LstVO.List(ESCContador)).Tag = ESCCuenta Then
                    If ESCOpcion = 1 Then
                        PMensajeC(LstVO.List(ESCContador)).SBEM.Panels(1).Text = ESCCuenta & " esta escribiendo un mensaje!!!"
                        Exit Sub
                    ElseIf ESCOpcion = 0 Then
                        PMensajeC(LstVO.List(ESCContador)).SBEM.Panels(1).Text = ""
                        Exit Sub
                    End If
                End If
                DoEvents
            Next
            Exit Sub
        End If
        
        If OpcionC = "PPROHIBIDA" Then
            Dim Me_Bloqueo As Integer
            MBMensaje = Split(DatosC, "¤¢©§¦[BOLA]¦§©¢¤")(0)
            Titulo = Split(DatosC, "¤¢©§¦[BOLA]¦§©¢¤")(1)
            Me_Bloqueo = Right(Titulo, 1)
            Titulo = Mid(Titulo, 1, Len(Titulo) - 1)
            MsgBox MBMensaje, vbDefaultButton1, Titulo
            If Me_Bloqueo = 1 Then
                FrmBloqueo.Show
                '////si el usuario alcanza el maximo de amonestaciones la teminal se apaga
            End If
            Exit Sub
        End If
        
        If OpcionC = "REMOTEPHOTO" Then ''Tomar el snapshot
            Dim TempFoto As String
            TempFoto = CStr("Desktop" & TxtUsuario.Text)
            If ProcesarFoto(TempFoto) = True Then
                EnviarFoto (TempFoto)
            End If
            Exit Sub
        End If
        
        If OpcionC = "BLOQUEAR" Then
            Call bloquearPC
            Exit Sub
        End If
        
        If OpcionC = "DESBLOQUEAR" Then
            Call DesbloquearPC
            Exit Sub
        End If
        
        If OpcionC = "NOTIFICACION" Then
            FrmMensaje.TextMensaje.Text = DatosC
            FrmMensaje.Show
            Exit Sub
        End If
        
        If OpcionC = "COMANDOAPI" Then
            Dim OptSWin As New Class_ShutDown
            If DatosC = "BLOQUEAR" Then
                Call bloquearPC
            ElseIf DatosC = "DESBLOQUEAR" Then
                Call DesbloquearPC
            ElseIf DatosC = "CERRARSESION" Then
                OptSWin.ExitWindows WE_LOGOFF
            ElseIf DatosC = "REINICIAR" Then
                OptSWin.ExitWindows WE_REBOOT
            ElseIf DatosC = "APAGAR" Then
                OptSWin.ExitWindows WE_SHUTDOWN
            ElseIf DatosC = "APAGARATX" Then
                OptSWin.ExitWindows WE_POWEROFF
            ElseIf DatosC = "TERMINARPROG" Then
                Unhook Me.hWnd
                Unload Me
                End
            ElseIf DatosC = "TERMINARREG" Then
                Call UnStartUP
            End If
                DoEvents
            Exit Sub
        End If
    
        If OpcionC = "NOMBREBOLA" Then
            Me.Caption = LblNivelS.Caption & ": " & DatosC & " en linea"
            Exit Sub
        End If
        
        If OpcionC = "SALIDA" Then
            Dim TempSalida As String
            Dim ContSalida As Integer
            Dim TempRV2 As Integer
            For i5 = 4 To TVUsuarios.Nodes.Count
                TempSalida = TVUsuarios.Nodes(i5).Text
                If TempSalida = DatosC Then
                    TVUsuarios.Nodes.Remove (i5)
                    If Not Len(TempSalida) = 0 Then
                        Call InicioSesion(CStr(TempSalida & Chr(10) & "Finalizo Sesión"))
                    End If
                    For ContSalida = 0 To LstVO.ListCount - 1
                        If PMensajeC(LstVO.List(ContSalida)).Tag = TempSalida Then
                            Unload PMensajeC(LstVO.List(ContSalida))
                            Exit Sub
                        End If
                    Next
                    Exit Sub
                End If
                DoEvents
            Next i5
            Exit Sub
        End If
End Sub

Private Sub bloquearPC()
    Dim TempFotoII As String
    TempFotoII = CStr("Bloqueo-" & TxtUsuario.Text)
    ProcesarFoto (TempFotoII)
    FrmBloqueo.FotoBloqueo = TempFotoII
    Call DesHabilitarRegistro
    FrmBloqueo.Show
End Sub
           
Private Sub DesbloquearPC()
    Unload FrmBloqueo
    Call HabilitarRegistro
    Exit Sub
End Sub

'**********************procesar,tomar, convertir y enviar el screenshot

Private Function ProcesarFoto(NombreFoto As String) As Boolean
    ProcesarFoto = False
    GetScreenShot FrmSnapShot.Picture1, NombreFoto
    ProcesarFoto = True
End Function

Private Sub EnviarFoto(NombrePhoto As String)
On Error GoTo Error
    Dim TempFile As String
    TempFile = (PicFolder & "\" & NombrePhoto & ".jpg")
    Open TempFile For Append As #1
    If LOF(1) = 0 Then
        MsgBox "Archivo vacio!!!", , "Atención!!!"
        Close #1
        Exit Sub
    End If
    Info.FileToSend = TempFile
    Info.FileSize = LOF(1)
    Close #1
    W_Yo.SendData "PETICION|" & getfilename(Info.FileToSend) & "|" & Info.FileSize & "|"
    DoEvents
    DoEvents
    DoEvents
    Exit Sub

Error:
    If Err.Number > 0 Then
        MsgBox Err.Number & " - " & Err.Description, , "Error Inesperado !!!"
    End If
    Resume Next
End Sub

Function getfilename(ByVal filepath As String)
    Dim ta() As String
    ta = Split(filepath, "\")
    getfilename = ta(UBound(ta))
End Function

'**********************fin procesar,tomar, convertir y enviar el screenshot

'''terminan funciones inpropias

''''**para notificar que ya se ha enviado el paquete
Private Sub W_Yo_SendComplete()
    ocupado = False
End Sub
''''**final (para notificar que ya se ha enviado el paquete)
Private Sub LLenarConfiguracion()
    ChkAL.Tag = Val(Lstconfig.List(0))
    ChkPR.Tag = Val(Lstconfig.List(1))
    ChkAA.Tag = Val(Lstconfig.List(2))
    ChkAP.Tag = Val(Lstconfig.List(3))
    ChkPP.Tag = Val(Lstconfig.List(4))
    ChkAL.Value = Val(Lstconfig.List(0))
    ChkPR.Value = Val(Lstconfig.List(1))
    ChkAA.Value = Val(Lstconfig.List(2))
    ChkAP.Value = Val(Lstconfig.List(3))
    ChkPP.Value = Val(Lstconfig.List(4))
    TxtAmonestaciones.Text = Lstconfig.List(5)
    TxtReportes.Text = Lstconfig.List(6)
    TxtAAC.Text = Lstconfig.List(7)
End Sub

Private Sub WUPD_DataArrival(ByVal bytesTotal As Long)
    On Error Resume Next

    Dim IdHost As String
    Dim IdHostFinal As String
    WUPD.GetData IdHost
    
    If InStr(1, IdHost, "ID¯[®©]¤¤¤[") Then
        IdHostFinal = Mid(IdHost, InStr(1, IdHost, "ID¯[®©]¤¤¤") + 4)
        IdHostFinal = Mid(IdHostFinal, 1, InStr(1, IdHostFinal, "]") - 1)
    End If

    If IdHostFinal <> "" And IdHostFinal <> TxtIP.Text Then
        TxtIP.Text = IdHostFinal
        TxtMaquina.Text = UCase(W_Yo.LocalHostName)
        Usr_Maquina = TxtMaquina.Text
    End If
End Sub

Private Sub CubrirPantalla()
    Me.Left = 0
    Me.Top = 0
    Me.Height = Screen.Height
    Me.Width = Screen.Width
End Sub

Private Sub MandarMensaje()
    Dim VMensajes() As String
    Dim UsrE_C1 As String
    Dim UsrE_C2 As String
    Dim UsrE_F1 As String
    Dim UsrE_N1 As Integer
    Dim UsrE_N2 As Integer
    Dim UsrE_I1 As Integer
    Dim UsrE_I2 As Integer
    
    Do While LstMensajes.ListCount > 0
        VMensajes = Split(LstMensajes.List(0), "¤¢©§¦[BOLA]¦§©¢¤")
        UsrE_I1 = VMensajes(0) 'puerto usuario que envia
        UsrE_N1 = VMensajes(1) 'nivel usuario que envia
        UsrE_C1 = VMensajes(2) 'cuenta usuario que envia
        UsrE_I2 = VMensajes(3) 'puerto usuario a quien se envia
        UsrE_N2 = VMensajes(4) 'nivel usuario a quien se envia
        UsrE_C2 = VMensajes(5) 'cuenta usuario a quien se envia
        UsrE_F1 = VMensajes(6) 'Mensaje
        If UsrE_I2 > 0 Then
            If ExisteVentana(UsrE_C2, UsrE_F1) = False Then
                Call NuevaVentana(UsrE_I1, UsrE_N1, UsrE_C1, Mi_Puerto, Usr_Nivel, TxtUsuario.Text, UsrE_F1, Usr_NivelS, ParseLevel(UsrE_N2))
            End If
        Else
            If ExisteVentana(UsrE_C2, UsrE_F1) = False Then
                Call NuevaVentana(UsrE_I2, UsrE_N2, UsrE_C2, Mi_Puerto, Usr_Nivel, TxtUsuario.Text, UsrE_F1, Usr_NivelS, ParseLevel(UsrE_N1))
            End If
        End If
        If LstMensajes.ListCount > 0 Then LstMensajes.RemoveItem (0)
        Erase VMensajes
    Loop
End Sub

Private Function ParseLevel(UNV As Integer) As String
    If UNV = 1 Then
         ParseLevel = "Alumno"
         Exit Function
    ElseIf UNV = 2 Then
         ParseLevel = "Profesor"
         Exit Function
    ElseIf UNV = 3 Then
         ParseLevel = "Administrador"
         Exit Function
    End If
End Function

Private Sub CargarLstVL()
    Dim IVl As Integer
    LstVL.Clear
    For IVl = 1 To 100
        LstVL.AddItem IVl
    Next
End Sub

Public Sub RemoverVentana(VRI As Integer)
    If LstVO.ListCount = 0 Then Exit Sub
    Dim IRV As Integer
    For IRV = 0 To LstVO.ListCount - 1
        If LstVO.List(IRV) = VRI Then LstVL.AddItem VRI: LstVO.RemoveItem (IRV): Exit Sub
    Next
End Sub

Private Sub CerrarVentana(CVUsuario As String)
    Dim ICV As Integer
    For ICV = 0 To LstVO.ListCount - 1
        If PMensajeC(LstVO.List(ICV)).Tag = CVUsuario Then Unload PMensajeC(LstVO.List(ICV)): Exit Sub
    Next
End Sub

Private Function ExisteVentana(CMUsuario As String, FraseUSRMsj As String) As Boolean
    If LstVO.ListCount = 0 Then ExisteVentana = False
    Dim IEV As Integer
    For IEV = 0 To LstVO.ListCount - 1
        If PMensajeC(LstVO.List(IEV)).Tag = CMUsuario Then PMensajeC(LstVO.List(IEV)).Txt_Respuesta.Text = PMensajeC(LstVO.List(IEV)).Txt_Respuesta.Text _
        & CMUsuario & ":" & FraseUSRMsj & Chr(10): ExisteVentana = True: Exit Function
    Next
    ExisteVentana = False
End Function

Public Sub NuevaVentana(PuertoA As Integer, NivelA As Integer, CuentaA As String, PuertoB As Integer, NivelB As Integer, CuentaB As String, FrasePM As String, Caption1 As String, Caption2 As String)
    If LstVL.ListCount = 0 Then Exit Sub
    Dim VX As Integer
    VX = LstVL.List(0)
    ReDim Preserve PMensajeC(VX)
    PMensajeC(VX).Tag = CuentaA
    PMensajeC(VX).Puerto1 = PuertoA
    PMensajeC(VX).Puerto2 = PuertoB
    PMensajeC(VX).Nivel1 = NivelA
    PMensajeC(VX).Nivel2 = NivelB
    PMensajeC(VX).Cuenta1 = CuentaA
    PMensajeC(VX).Cuenta2 = CuentaB
    PMensajeC(VX).NVIndex = VX
    PMensajeC(VX).Caption = "::: Ayuda Directa ::: " & Caption1 & ": " & CuentaB & " | " & Caption2 & ": " & CuentaA & " :::"
    PMensajeC(VX).Txt_Respuesta.Text = CuentaA & ":" & FrasePM & Chr(10)
    PMensajeC(VX).Visible = True
    LstVO.AddItem (VX)
    LstVL.RemoveItem (0)
End Sub

Private Function ExisteVentanaII(CMUsuario As String) As Boolean
    If LstVO.ListCount = 0 Then ExisteVentanaII = False
    Dim IEVII As Integer
    For IEVII = 0 To LstVO.ListCount - 1
        If PMensajeC(LstVO.List(IEVII)).Tag = CMUsuario Then ExisteVentanaII = True: Exit Function
    Next
    ExisteVentanaII = False
End Function

Private Sub CerrarSesion(Opt As Integer)
    If Opt = 1 Then
        If ConectandoP = True Then Exit Sub
        Unhook Me.hWnd
        Call DesHabilitarRegistro
        Tmr_Hora.Enabled = False
        'TmrOnTop.Enabled = True ' precaucion se activara la proteccion completa
        
        W_Yo.Tag = ""
        W_Yo.Close
        Me.BorderStyle = 0
        
        TxtUsuario.Text = ""
        TxtPassword.Text = ""
        TxtMaquina.Text = ""
        TxtUrlLog.Text = ""
        TxtIP.Text = ""
        Admin = ""
        
        LstProcesos.Clear
        LstRestringidos.Clear
        LstPBloqueados.Clear
        LstMsjEspera.Clear
        Lstrecibidos.Clear
        
        TxtProcSel.Text = ""
        LblNPB.Caption = "0"
        LblNPR.Caption = "0"
        TxtMinutos.Text = "0"
        TxtAAC.Text = "0"
        TxtReportes.Text = "0"
        TxtAmonestaciones.Text = "0"
        
        TxtTSesion.Text = "------------------------------------------------------"
        
        ChkAL.Value = 0
        ChkPR.Value = 0
        ChkAA.Value = 0
        ChkAA.Value = 0
        ChkAP.Value = 0
        ChkPP.Value = 0
        
        Lstconfig.Clear
        LstMensajes.Clear
        LstVL.Clear
        LstVO.Clear
        
        CargarLstVL
        
        CmdConectar.Enabled = True
        CmdCerrarSesion.Enabled = False
        CmdNU.Enabled = True
        
        TxtMaquina.Locked = False
        TxtUsuario.Locked = False
        TxtPassword.Locked = False
        TxtIP.Locked = False
        Call CerrarUDP(False)
        
        TxtMaquina.Text = UCase(W_Yo.LocalHostName)
        
        Me.Caption = ""
        Me.WindowState = 0
        Me.Left = 0
        Me.Top = 0
        Me.Width = Screen.Width
        Me.Height = Screen.Height
        
        Call ColocarControles
    Else
        Me.Visible = False
        Call HabilitarRegistro
        CmdCerrarSesion.Enabled = True
        
        TmrOnTop.Enabled = False
        MaxTop Me.hWnd, False
        
        TxtMaquina.Locked = True
        TxtUsuario.Locked = True
        TxtPassword.Locked = True
        TxtIP.Locked = True
        CmdConectar.Enabled = False
        CmdNU.Enabled = False
        
        If W_Yo.Tag <> "" Then
            ConectandoP = False
            Me.WindowState = 0
            Me.BorderStyle = 1
            Me.Width = PicContenedor.Width + 1000
            Me.Height = PicContenedor.Height + 1000
            Me.Left = (Screen.Width / 2) - Me.Width / 2
            Me.Top = (Screen.Height / 2) - Me.Height / 2
            Call ColocarControles
        End If
        Me.Visible = True
    End If
    TxtNNC.Text = ""
    TxtNV.Text = ""
    TxtNP.Text = ""
    On Error Resume Next: TxtUsuario.SetFocus
End Sub

Private Sub CargarMenu()
    Dim TitulosB() As Variant
    Dim i As Integer
    TitulosB = Array("Ayuda", "Ayuda Directa", "Reportes", "Cerrar Sesion", "Reiniciar", "Apagar", "Apagar ATX")
    LV1.ListItems.Clear
    For i = 0 To 6
        LV1.ListItems.Add , TitulosB(i), TitulosB(i), i + 1
        DoEvents
    Next
    Erase TitulosB
End Sub
