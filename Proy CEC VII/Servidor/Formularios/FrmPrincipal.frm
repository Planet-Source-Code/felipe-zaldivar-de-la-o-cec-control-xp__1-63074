VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form FrmPrincipal 
   BackColor       =   &H00FFCEBB&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "::: Control de Acceso a Centros de Computo :::"
   ClientHeight    =   7860
   ClientLeft      =   555
   ClientTop       =   1950
   ClientWidth     =   11115
   ClipControls    =   0   'False
   Icon            =   "FrmPrincipal.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7860
   ScaleWidth      =   11115
   Begin MSComctlLib.StatusBar SB 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   48
      Top             =   7485
      Width           =   11115
      _ExtentX        =   19606
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   14420
            Text            =   "Programado por Felipe Zaldivar de la O  © 2005"
            TextSave        =   "Programado por Felipe Zaldivar de la O  © 2005"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            TextSave        =   "22/10/2005"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            AutoSize        =   2
            TextSave        =   "11:39 p.m."
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab TabMonitor 
      Height          =   7215
      Left            =   3360
      TabIndex        =   39
      Top             =   120
      Width           =   7605
      _ExtentX        =   13414
      _ExtentY        =   12726
      _Version        =   393216
      Style           =   1
      Tabs            =   6
      Tab             =   4
      TabsPerRow      =   6
      TabHeight       =   520
      BackColor       =   16764603
      TabCaption(0)   =   "&Usuarios"
      TabPicture(0)   =   "FrmPrincipal.frx":08CA
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Picture9"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "&Procesos"
      TabPicture(1)   =   "FrmPrincipal.frx":08E6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Picture8"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "&Reportes"
      TabPicture(2)   =   "FrmPrincipal.frx":0902
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Picture7"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "&Maquinas"
      TabPicture(3)   =   "FrmPrincipal.frx":091E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Picture1"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "M&ensajes"
      TabPicture(4)   =   "FrmPrincipal.frx":093A
      Tab(4).ControlEnabled=   -1  'True
      Tab(4).Control(0)=   "Picture10"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "&Configuración"
      TabPicture(5)   =   "FrmPrincipal.frx":0956
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Picture5"
      Tab(5).ControlCount=   1
      Begin VB.PictureBox Picture5 
         BackColor       =   &H00E89C78&
         BorderStyle     =   0  'None
         Height          =   6615
         Left            =   -74880
         ScaleHeight     =   6615
         ScaleWidth      =   7335
         TabIndex        =   77
         Top             =   480
         Width           =   7335
         Begin VB.PictureBox Picture4 
            BackColor       =   &H00D05C28&
            BorderStyle     =   0  'None
            Height          =   3855
            Left            =   120
            ScaleHeight     =   3855
            ScaleWidth      =   7095
            TabIndex        =   86
            Top             =   2040
            Width           =   7095
            Begin VB.CheckBox ChkAyuda 
               BackColor       =   &H00D05C28&
               Caption         =   "No"
               ForeColor       =   &H80000005&
               Height          =   255
               Left            =   120
               TabIndex        =   34
               Top             =   720
               UseMaskColor    =   -1  'True
               Width           =   615
            End
            Begin VB.CheckBox ChkContesta 
               BackColor       =   &H00D05C28&
               Caption         =   "No"
               ForeColor       =   &H80000005&
               Height          =   375
               Left            =   120
               TabIndex        =   35
               Top             =   1560
               UseMaskColor    =   -1  'True
               Width           =   735
            End
            Begin VB.TextBox TxtMensaje 
               Height          =   615
               Left            =   120
               MaxLength       =   100
               MultiLine       =   -1  'True
               TabIndex        =   36
               Top             =   2400
               Width           =   6855
            End
            Begin VB.CheckBox ChkBP 
               BackColor       =   &H00D05C28&
               Caption         =   "No"
               ForeColor       =   &H80000005&
               Height          =   255
               Left            =   120
               TabIndex        =   37
               Top             =   3480
               UseMaskColor    =   -1  'True
               Width           =   615
            End
            Begin VB.Label Label24 
               AutoSize        =   -1  'True
               BackColor       =   &H00D05C28&
               Caption         =   "Bloquear procesos restringidos:"
               ForeColor       =   &H80000005&
               Height          =   195
               Left            =   120
               TabIndex        =   91
               Top             =   3120
               Width           =   2205
            End
            Begin VB.Label Label23 
               AutoSize        =   -1  'True
               BackColor       =   &H00D05C28&
               Caption         =   "Mensaje:"
               ForeColor       =   &H80000005&
               Height          =   195
               Left            =   120
               TabIndex        =   90
               Top             =   2040
               Width           =   645
            End
            Begin VB.Label Label22 
               AutoSize        =   -1  'True
               BackColor       =   &H00D05C28&
               Caption         =   "Contestadora automática:"
               ForeColor       =   &H80000005&
               Height          =   195
               Left            =   120
               TabIndex        =   89
               Top             =   1200
               Width           =   1815
            End
            Begin VB.Label Label21 
               AutoSize        =   -1  'True
               BackColor       =   &H00D05C28&
               Caption         =   "Ayuda en linea:"
               ForeColor       =   &H80000005&
               Height          =   195
               Left            =   120
               TabIndex        =   88
               Top             =   360
               Width           =   1095
            End
            Begin VB.Label Label10 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackColor       =   &H00D05C28&
               Caption         =   "Varios"
               ForeColor       =   &H80000005&
               Height          =   195
               Left            =   3315
               TabIndex        =   87
               Top             =   120
               Width           =   465
            End
         End
         Begin VB.PictureBox Picture2 
            BackColor       =   &H00D05C28&
            BorderStyle     =   0  'None
            ForeColor       =   &H00FFFFFF&
            Height          =   1815
            Left            =   120
            ScaleHeight     =   1815
            ScaleWidth      =   3510
            TabIndex        =   81
            Top             =   120
            Width           =   3510
            Begin VB.ComboBox CboReportes 
               Height          =   315
               Left            =   2400
               Style           =   2  'Dropdown List
               TabIndex        =   28
               Top             =   420
               Width           =   855
            End
            Begin VB.ComboBox CboAmonestaciones 
               Height          =   315
               Left            =   2400
               Style           =   2  'Dropdown List
               TabIndex        =   30
               Top             =   1380
               Width           =   855
            End
            Begin VB.ComboBox CboHistorial 
               Height          =   315
               Left            =   2400
               Style           =   2  'Dropdown List
               TabIndex        =   29
               Top             =   900
               Width           =   855
            End
            Begin VB.Label Label19 
               AutoSize        =   -1  'True
               BackColor       =   &H00D05C28&
               Caption         =   "Días de Historial:"
               ForeColor       =   &H80000005&
               Height          =   195
               Left            =   240
               TabIndex        =   85
               Top             =   960
               Width           =   1215
            End
            Begin VB.Label Label17 
               AutoSize        =   -1  'True
               BackColor       =   &H00D05C28&
               Caption         =   "Tiempo Reportes:"
               ForeColor       =   &H80000005&
               Height          =   195
               Left            =   240
               TabIndex        =   84
               Top             =   480
               Width           =   1260
            End
            Begin VB.Label Label20 
               AutoSize        =   -1  'True
               BackColor       =   &H00D05C28&
               Caption         =   "Maximo de Amonestaciones:"
               ForeColor       =   &H80000005&
               Height          =   195
               Left            =   240
               TabIndex        =   83
               Top             =   1440
               Width           =   2025
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackColor       =   &H00D05C28&
               Caption         =   "Conversaciones Permitidas:"
               ForeColor       =   &H80000005&
               Height          =   195
               Left            =   840
               TabIndex        =   82
               Top             =   120
               Width           =   1950
            End
         End
         Begin VB.PictureBox Picture3 
            BackColor       =   &H00D05C28&
            BorderStyle     =   0  'None
            Height          =   1815
            Left            =   3720
            ScaleHeight     =   1815
            ScaleWidth      =   3510
            TabIndex        =   79
            Top             =   120
            Width           =   3510
            Begin VB.CheckBox ChkPP 
               BackColor       =   &H00D05C28&
               Caption         =   "Profesor -  Profesor"
               ForeColor       =   &H80000005&
               Height          =   375
               Left            =   120
               TabIndex        =   33
               Top             =   1200
               UseMaskColor    =   -1  'True
               Width           =   1815
            End
            Begin VB.CheckBox ChkAA 
               BackColor       =   &H00D05C28&
               Caption         =   "Alumno - Alumno"
               ForeColor       =   &H80000005&
               Height          =   255
               Left            =   120
               TabIndex        =   31
               Top             =   480
               UseMaskColor    =   -1  'True
               Width           =   1695
            End
            Begin VB.CheckBox ChkAP 
               BackColor       =   &H00D05C28&
               Caption         =   "Alumno -  Profesor"
               ForeColor       =   &H80000005&
               Height          =   255
               Left            =   120
               TabIndex        =   32
               Top             =   840
               UseMaskColor    =   -1  'True
               Width           =   1815
            End
            Begin VB.Label Label18 
               AutoSize        =   -1  'True
               BackColor       =   &H00D05C28&
               Caption         =   "Conversaciones Permitidas:"
               ForeColor       =   &H80000005&
               Height          =   195
               Left            =   120
               TabIndex        =   80
               Top             =   120
               Width           =   1950
            End
         End
         Begin VB.PictureBox Picture6 
            BackColor       =   &H00D05C28&
            BorderStyle     =   0  'None
            Height          =   495
            Left            =   120
            ScaleHeight     =   495
            ScaleWidth      =   7095
            TabIndex        =   78
            Top             =   6000
            Width           =   7095
            Begin VB.CommandButton CmdAceptar 
               BackColor       =   &H00D05C28&
               Caption         =   "&Aceptar"
               Height          =   255
               Left            =   3000
               TabIndex        =   38
               Top             =   120
               Width           =   1215
            End
         End
      End
      Begin VB.PictureBox Picture10 
         BackColor       =   &H00E89C78&
         BorderStyle     =   0  'None
         Height          =   6615
         Left            =   120
         ScaleHeight     =   6615
         ScaleWidth      =   7335
         TabIndex        =   71
         Top             =   480
         Width           =   7335
         Begin VB.PictureBox Picture14 
            BackColor       =   &H00D05C28&
            BorderStyle     =   0  'None
            Height          =   3135
            Left            =   120
            ScaleHeight     =   3135
            ScaleWidth      =   7095
            TabIndex        =   75
            Top             =   120
            Width           =   7095
            Begin VB.TextBox TxtNotificacion 
               Height          =   1695
               Left            =   240
               MaxLength       =   250
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   17
               Top             =   600
               Width           =   6615
            End
            Begin VB.CommandButton CmdMsjMasivo 
               BackColor       =   &H00D05C28&
               Caption         =   "E&nviar"
               Height          =   375
               Left            =   5400
               TabIndex        =   18
               Top             =   2520
               Width           =   1455
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Mensaje a todos los usuarios:"
               ForeColor       =   &H80000005&
               Height          =   195
               Left            =   240
               TabIndex        =   76
               Top             =   240
               Width           =   2085
            End
         End
         Begin VB.PictureBox Picture11 
            BackColor       =   &H00D05C28&
            BorderStyle     =   0  'None
            Height          =   3135
            Left            =   120
            ScaleHeight     =   3135
            ScaleWidth      =   7095
            TabIndex        =   72
            Top             =   3360
            Width           =   7095
            Begin VB.OptionButton Optbloq 
               BackColor       =   &H00D05C28&
               Caption         =   "Bloquear Máquinas"
               ForeColor       =   &H80000005&
               Height          =   315
               Left            =   240
               MaskColor       =   &H80000005&
               TabIndex        =   19
               Top             =   360
               Width           =   1935
            End
            Begin VB.OptionButton OptDsbloq 
               BackColor       =   &H00D05C28&
               Caption         =   "Desbloquear Máquinas"
               ForeColor       =   &H80000005&
               Height          =   255
               Left            =   2520
               MaskColor       =   &H80000005&
               TabIndex        =   20
               Top             =   360
               Width           =   2055
            End
            Begin VB.OptionButton OptApagar 
               BackColor       =   &H00D05C28&
               Caption         =   "Apagar Maquinas"
               ForeColor       =   &H80000005&
               Height          =   255
               Left            =   2520
               MaskColor       =   &H80000005&
               TabIndex        =   23
               Top             =   960
               Width           =   2055
            End
            Begin VB.OptionButton OptReiniciar 
               BackColor       =   &H00D05C28&
               Caption         =   "Reiniciar Máquinas"
               ForeColor       =   &H80000005&
               Height          =   255
               Left            =   240
               MaskColor       =   &H80000005&
               TabIndex        =   22
               Top             =   960
               Width           =   1935
            End
            Begin VB.OptionButton OptCSesion 
               BackColor       =   &H00D05C28&
               Caption         =   "Cerrar Sesión Máquinas"
               ForeColor       =   &H80000005&
               Height          =   255
               Left            =   4920
               MaskColor       =   &H80000005&
               TabIndex        =   21
               Top             =   360
               Width           =   2055
            End
            Begin VB.CommandButton CmdMasivo 
               BackColor       =   &H00D05C28&
               Caption         =   "Acep&tar"
               Height          =   375
               Left            =   2640
               TabIndex        =   27
               Top             =   2280
               Width           =   2055
            End
            Begin VB.OptionButton OptAAtx 
               BackColor       =   &H00D05C28&
               Caption         =   "Apagar Maquinas ATX"
               ForeColor       =   &H80000005&
               Height          =   255
               Left            =   4920
               MaskColor       =   &H80000005&
               TabIndex        =   24
               Top             =   960
               Width           =   2055
            End
            Begin VB.OptionButton OptTerminar 
               BackColor       =   &H00D05C28&
               Caption         =   "Terminar ""Cliente"""
               ForeColor       =   &H80000005&
               Height          =   255
               Left            =   240
               MaskColor       =   &H80000005&
               TabIndex        =   25
               Top             =   1560
               Width           =   1935
            End
            Begin VB.OptionButton OptReg 
               BackColor       =   &H00D05C28&
               Caption         =   "Remover del Registro"
               ForeColor       =   &H80000005&
               Height          =   255
               Left            =   2520
               MaskColor       =   &H80000005&
               TabIndex        =   26
               Top             =   1560
               Width           =   2055
            End
            Begin VB.Label LblcmdOpcion 
               Caption         =   "BLOQUEAR"
               Height          =   375
               Left            =   120
               TabIndex        =   74
               Top             =   2160
               Visible         =   0   'False
               Width           =   2055
            End
            Begin VB.Label LblOptCaption 
               Caption         =   "Bloquear Máquinas"
               Height          =   375
               Left            =   120
               TabIndex        =   73
               Top             =   2640
               Visible         =   0   'False
               Width           =   2055
            End
         End
      End
      Begin VB.PictureBox Picture9 
         BackColor       =   &H00E89C78&
         BorderStyle     =   0  'None
         Height          =   6615
         Left            =   -74880
         ScaleHeight     =   6615
         ScaleWidth      =   7335
         TabIndex        =   70
         Top             =   480
         Width           =   7335
         Begin MSComctlLib.ListView LstVwUsrIn 
            Height          =   6375
            Left            =   120
            TabIndex        =   2
            Top             =   120
            Width           =   7065
            _ExtentX        =   12462
            _ExtentY        =   11245
            SortKey         =   4
            View            =   3
            LabelEdit       =   1
            Sorted          =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            AllowReorder    =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            Appearance      =   1
            NumItems        =   5
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Clave"
               Object.Width           =   2030
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Maquina"
               Object.Width           =   2558
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Puerto"
               Object.Width           =   2558
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "IP"
               Object.Width           =   2558
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Key             =   "LV_Inicio"
               Text            =   "Inicio"
               Object.Width           =   2558
            EndProperty
         End
      End
      Begin VB.PictureBox Picture8 
         BackColor       =   &H00E89C78&
         BorderStyle     =   0  'None
         Height          =   6615
         Left            =   -74880
         ScaleHeight     =   6615
         ScaleWidth      =   7335
         TabIndex        =   64
         Top             =   480
         Width           =   7335
         Begin MSFlexGridLib.MSFlexGrid FGP 
            Height          =   4815
            Left            =   120
            TabIndex        =   8
            Tag             =   "1"
            Top             =   120
            Width           =   7100
            _ExtentX        =   12515
            _ExtentY        =   8493
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
            BackColorBkg    =   -2147483643
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
         Begin VB.TextBox TxtUH 
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   5430
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   6
            Top             =   5400
            Width           =   1790
         End
         Begin VB.TextBox TxtUF 
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   5430
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   4
            Top             =   5040
            Width           =   1790
         End
         Begin VB.TextBox TxtUP 
            BorderStyle     =   0  'None
            Height          =   645
            Left            =   870
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   7
            Top             =   5820
            Width           =   6350
         End
         Begin VB.TextBox TxtUM 
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   870
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   5
            Top             =   5400
            Width           =   3735
         End
         Begin VB.TextBox TxtUC 
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   870
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   3
            Top             =   5040
            Width           =   3735
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Proceso:"
            Height          =   195
            Left            =   75
            TabIndex        =   69
            Top             =   5820
            Width           =   630
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Usuario:"
            Height          =   195
            Left            =   75
            TabIndex        =   68
            Top             =   5085
            Width           =   585
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha:"
            Height          =   195
            Left            =   4830
            TabIndex        =   67
            Top             =   5085
            Width           =   495
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Hora:"
            Height          =   195
            Left            =   4830
            TabIndex        =   66
            Top             =   5445
            Width           =   390
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Maquina:"
            Height          =   195
            Left            =   75
            TabIndex        =   65
            Top             =   5445
            Width           =   660
         End
      End
      Begin VB.PictureBox Picture7 
         BackColor       =   &H00E89C78&
         BorderStyle     =   0  'None
         Height          =   6615
         Left            =   -74880
         ScaleHeight     =   6615
         ScaleWidth      =   7335
         TabIndex        =   57
         Top             =   480
         Width           =   7335
         Begin MSFlexGridLib.MSFlexGrid FGR 
            Height          =   4755
            Left            =   120
            TabIndex        =   15
            Tag             =   "1"
            Top             =   120
            Width           =   7100
            _ExtentX        =   12515
            _ExtentY        =   8387
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
            BackColorBkg    =   -2147483643
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
         Begin VB.TextBox TxtRR 
            BorderStyle     =   0  'None
            Height          =   645
            Left            =   885
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   14
            Top             =   5805
            Width           =   6310
         End
         Begin VB.TextBox TxtRH 
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   5745
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   13
            Top             =   5415
            Width           =   1460
         End
         Begin VB.TextBox TxtRF 
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   5745
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   11
            Top             =   5040
            Width           =   1460
         End
         Begin VB.TextBox TxtRM 
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   2820
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   10
            Top             =   5040
            Width           =   2175
         End
         Begin VB.TextBox TxtRU 
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   885
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   9
            Top             =   5040
            Width           =   975
         End
         Begin VB.TextBox TxtRT 
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   885
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   12
            Top             =   5415
            Width           =   4110
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Título:"
            Height          =   195
            Left            =   120
            TabIndex        =   63
            Top             =   5460
            Width           =   465
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Reporte:"
            Height          =   195
            Left            =   120
            TabIndex        =   62
            Top             =   5850
            Width           =   615
         End
         Begin VB.Label Label27 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Usuario:"
            Height          =   195
            Left            =   120
            TabIndex        =   61
            Top             =   5085
            Width           =   585
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha:"
            Height          =   195
            Left            =   5100
            TabIndex        =   60
            Top             =   5085
            Width           =   495
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Hora:"
            Height          =   195
            Left            =   5100
            TabIndex        =   59
            Top             =   5460
            Width           =   390
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Maquina:"
            Height          =   195
            Left            =   1980
            TabIndex        =   58
            Top             =   5085
            Width           =   660
         End
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00E89C78&
         BorderStyle     =   0  'None
         Height          =   6495
         Left            =   -74880
         ScaleHeight     =   6495
         ScaleWidth      =   7335
         TabIndex        =   56
         Top             =   520
         Width           =   7335
         Begin MSComctlLib.ListView LVM 
            Height          =   6195
            Left            =   150
            TabIndex        =   16
            Top             =   150
            Width           =   7005
            _ExtentX        =   12356
            _ExtentY        =   10927
            Arrange         =   2
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            OLEDragMode     =   1
            FlatScrollBar   =   -1  'True
            _Version        =   393217
            Icons           =   "ILM"
            ForeColor       =   16777215
            BackColor       =   13655080
            Appearance      =   0
            OLEDragMode     =   1
            NumItems        =   0
         End
         Begin MSComctlLib.ImageList ILM 
            Left            =   0
            Top             =   0
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   64
            ImageHeight     =   64
            MaskColor       =   12632256
            _Version        =   393216
            BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
               NumListImages   =   3
               BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FrmPrincipal.frx":0972
                  Key             =   ""
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FrmPrincipal.frx":3BCC
                  Key             =   ""
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FrmPrincipal.frx":6E26
                  Key             =   ""
               EndProperty
            EndProperty
         End
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFCEBB&
      Caption         =   "Servidor  |  Usuarios en Linea"
      ForeColor       =   &H80000007&
      Height          =   7215
      Left            =   120
      TabIndex        =   45
      Top             =   120
      Width           =   3135
      Begin MSComctlLib.TreeView TVUsuarios 
         Height          =   6375
         Left            =   120
         TabIndex        =   1
         Top             =   720
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   11245
         _Version        =   393217
         Indentation     =   0
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         FullRowSelect   =   -1  'True
         ImageList       =   "ImageList1"
         Appearance      =   0
      End
      Begin MSComDlg.CommonDialog Common 
         Left            =   1560
         Top             =   5520
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton Cmd_Cerrar_T 
         Caption         =   "Cerrar Terminales"
         Height          =   375
         Left            =   6240
         TabIndex        =   46
         Top             =   5880
         Width           =   1455
      End
      Begin VB.TextBox TxtUsuario 
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   0
         Text            =   "TxtUsuario"
         Top             =   337
         Width           =   1815
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Administrador:"
         Height          =   195
         Left            =   120
         TabIndex        =   47
         Top             =   360
         Width           =   990
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "Fecha y Hora"
      Height          =   1575
      Left            =   3600
      TabIndex        =   42
      Top             =   360
      Visible         =   0   'False
      Width           =   3615
      Begin VB.ListBox LstVO 
         Height          =   255
         Left            =   2400
         Sorted          =   -1  'True
         TabIndex        =   55
         Top             =   720
         Width           =   1095
      End
      Begin VB.ListBox LstVL 
         Height          =   255
         Left            =   2400
         Sorted          =   -1  'True
         TabIndex        =   54
         Top             =   1080
         Width           =   1095
      End
      Begin VB.ListBox LstMensajes 
         Height          =   255
         ItemData        =   "FrmPrincipal.frx":A080
         Left            =   2400
         List            =   "FrmPrincipal.frx":A082
         TabIndex        =   53
         Top             =   360
         Width           =   1095
      End
      Begin VB.ListBox Lstrecibidos 
         Height          =   255
         Left            =   120
         TabIndex        =   52
         Top             =   1080
         Width           =   1095
      End
      Begin VB.ListBox LstBloqueo 
         Height          =   255
         Left            =   1320
         TabIndex        =   51
         Top             =   1080
         Width           =   975
      End
      Begin VB.ListBox LstPuerto 
         Height          =   255
         ItemData        =   "FrmPrincipal.frx":A084
         Left            =   1320
         List            =   "FrmPrincipal.frx":A08B
         TabIndex        =   50
         Top             =   720
         Width           =   975
      End
      Begin VB.ListBox LstMsjEspera 
         Height          =   255
         ItemData        =   "FrmPrincipal.frx":A09A
         Left            =   120
         List            =   "FrmPrincipal.frx":A0A1
         TabIndex        =   49
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox TxtFecha 
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   44
         Text            =   "TxtFecha"
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox TxtHora 
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   43
         Text            =   "TxtHora"
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Diversos Controles "
      Height          =   1455
      Left            =   3600
      TabIndex        =   41
      Top             =   1920
      Visible         =   0   'False
      Width           =   3615
      Begin VB.Timer TmrRecibidos 
         Left            =   2400
         Top             =   315
      End
      Begin VB.Timer TmrMonitor 
         Interval        =   1
         Left            =   1800
         Top             =   315
      End
      Begin VB.Timer Tmr_Hora 
         Interval        =   1
         Left            =   1320
         Top             =   315
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   120
         Top             =   240
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   12632256
         UseMaskColor    =   0   'False
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   3
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmPrincipal.frx":A0B3
               Key             =   "Admin"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmPrincipal.frx":AD8D
               Key             =   "Profr"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmPrincipal.frx":BA67
               Key             =   "Alumno"
               Object.Tag             =   "sdf"
            EndProperty
         EndProperty
      End
      Begin MSWinsockLib.Winsock W_Usr 
         Index           =   0
         Left            =   120
         Top             =   840
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSComDlg.CommonDialog CD 
         Left            =   720
         Top             =   285
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSWinsockLib.Winsock WTCP 
         Left            =   600
         Top             =   840
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSWinsockLib.Winsock WUPD 
         Left            =   1080
         Top             =   840
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000002&
      Caption         =   "Controles Data"
      Height          =   3975
      Left            =   7200
      TabIndex        =   40
      Top             =   360
      Visible         =   0   'False
      Width           =   4455
      Begin VB.Data DataChequeo 
         Caption         =   "DataChequeo///checa las maquinas ocupadas"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   120
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   3480
         Width           =   4335
      End
      Begin VB.Data DataCNU 
         Caption         =   "DataCNU ////data consultas usr"
         Connect         =   "Access 2000;"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   120
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   2760
         Visible         =   0   'False
         Width           =   4335
      End
      Begin VB.Data DataANU 
         Caption         =   "DataANU ///altas usuario"
         Connect         =   "Access 2000;"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   420
         Left            =   120
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   3120
         Visible         =   0   'False
         Width           =   4335
      End
      Begin VB.Data DataRep 
         Caption         =   "DataRep ///agregar reportes"
         Connect         =   "Access 2000;"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   120
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   2400
         Width           =   4335
      End
      Begin VB.Data DataModifAm 
         Caption         =   "DataModifAm ///checar frase"
         Connect         =   "Access 2000;"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   120
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   2040
         Width           =   4335
      End
      Begin VB.Data DataReservadas 
         Caption         =   "DataReservadas ///checa palabras reservadas"
         Connect         =   "Access 2000;"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   120
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   1680
         Width           =   4335
      End
      Begin VB.Data DataModif 
         Caption         =   "DataModif ///Modificaciones a la BD"
         Connect         =   "Access 2000;"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   120
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   1320
         Visible         =   0   'False
         Width           =   4335
      End
      Begin VB.Data DataMaq 
         Caption         =   "DataMaq  //Todo respecto a la Maquina"
         Connect         =   "Access 2000;"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   120
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   960
         Visible         =   0   'False
         Width           =   4335
      End
      Begin VB.Data DataChec 
         Caption         =   "DataChec //checa si es la BD deseada"
         Connect         =   "Access 2000;"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   120
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   600
         Visible         =   0   'False
         Width           =   4335
      End
      Begin VB.Data DataLogin 
         Caption         =   "DataLogin //Acceso al sistema"
         Connect         =   "Access 2000;"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   120
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   240
         Visible         =   0   'False
         Width           =   4335
      End
   End
   Begin VB.Menu Mnu_OpUsr 
      Caption         =   "Opciones"
      Visible         =   0   'False
      Begin VB.Menu Cmd_IG 
         Caption         =   "Información General"
      End
      Begin VB.Menu Cmd_IP 
         Caption         =   "Información Personal"
      End
      Begin VB.Menu Cmd_Conv 
         Caption         =   "Conversación"
      End
      Begin VB.Menu Cmd_TF 
         Caption         =   "Tomar Foto"
      End
      Begin VB.Menu Cmd_Block 
         Caption         =   "Bloquear"
      End
   End
End
Attribute VB_Name = "FrmPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Dim Cuantos As Long

Private Type da
    FileToSend As String
    FileName As String
    RemoteIP As String
    FileSize As Double
    SaveAs As String
    Pstatus As Double
    LastAmount As Double
End Type

Private Info As da

Private Type InfoUsuario
    Cuenta_IU As String
    Nivel_IU As Integer
    Index_IU As Integer
    NivelS_IU As String
    Maquina_IU As String
End Type

Private InfoUsuario_T As InfoUsuario
Public respuesta As String '' guarda la respuesta de un msgbox determinado por el usuario
'Public Usuario As Integer '' guarda el puerto de un usuario determinado
Public Usr_Accesado As Boolean '' si el usuario a accesado al sistema
Public Usr_Cuenta As String ''guarda temporalente la cuenta de quien va a accesar
Public Usr_Password As String ''guarda temporalmente el password de quien va a accesar
Public Usr_Maquina As String ''guarda temporalente el nombre de la maquina de quien va a accesar
Public Ocupado As Boolean
Public Usuario As Integer ''' numero de conexiones en el sistema

Dim contador As Integer 'contador de veces que intenta maximo 3
Dim Opcion As String ''guarda la opcion que se nos ha enviado o enviamos
'Dim Datos As String ''guarda la cadena que se nos ha enviado
Dim Datos2 As String ''guarda lo que esta despues del comodin primario "+"
Dim Datos3 As String ''datos opcionales en el acceso
Dim Posicion1 As Integer ''guarda la posicion del comodin primario "+"
Dim Posicion2 As Integer ''guarda la posicion del comodin secundario":"
''' guardan datos del usuario seleccionado en el arbol  *****111******
Dim Usr_I1 As Integer ''guarda el puerto en el que esta ese usuario
Dim Usr_N1 As Integer ''guarda el nivel jerarquico de ese usuario
Dim Usr_M1 As String '''guarda la maquina del usuario
Dim Usr_C1 As String '''guarda la cuenta del usuario
''''termina ******111******
''' guardan datos para cuando alguien quiere establecer conversacioncon nosotros *****111******
Dim Usr_I2 As Integer ''guarda el puerto en el que esta ese usuario
Dim Usr_N2 As Integer ''guarda el nivel jerarquico de ese usuario
Dim Usr_M2 As String '''guarda la maquina del usuario
Dim Usr_C2 As String '''guarda la cuenta del usuario
'''******/////////////*********
Dim Usr_I3 As Integer ''guarda el puerto en el que esta ese usuario
Dim Usr_N3 As Integer ''guarda el nivel jerarquico de ese usuario
Dim Usr_M3 As String '''guarda la maquina del usuario
Dim Usr_C3 As String '''guarda la cuenta del usuario
''''termina ******111******
Dim Fecha_A As Date 'guarda la fecha anterior del registro en el historial general
Dim I3 As Integer '''contador para abrir ventas PM
Dim I4 As Integer ''' contador para desocupar la maquina ocupada
Dim i5 As Integer ''' contador para desocupar la maquina ocupada
Dim i6 As Integer
Dim i7 As Integer
Dim i8 As Integer
Dim I10 As Integer ''' contador para enviar datos a todos los usuarios exepto al owner
Dim hSysMenu As Long ' hwnd para remover el boton (X)
Dim Pa As String 'path de la imagen
Dim fPath As String 'Paht y nombre de la imagen

Private Sub ChkAyuda_Click()
    If ChkAyuda.Value = 1 Then
        ChkAyuda.Caption = "Si"
    Else
        ChkAyuda.Caption = "No"
    End If
End Sub

Private Sub ChkBP_Click()
    If ChkBP.Value = 1 Then
        ChkBP.Caption = "Si"
    Else
        ChkBP.Caption = "No"
    End If
End Sub

Private Sub ChkContesta_Click()
    If ChkContesta.Value = 1 Then
        ChkContesta.Caption = "Si"
    Else
        ChkContesta.Caption = "No"
    End If
End Sub

'//////aqui van las funciones de controles o formularios
Private Sub Cmd_Block_Click()
    ObtenerInformacionTvUsr
    'MsgBox Usr_I_Temp1 & Chr(10) & Usr_N_Temp1 & Chr(10) & Usr_C_Temp1 & Chr(10) & 0 & Chr(10) & 3 & Chr(10) & TxtUsuario.Text
    If InfoUsuario_T.Index_IU >= 0 Then
        If Conectado(InfoUsuario_T.Index_IU) Then
            Call Enviar("BLOQUEAR¯[®©]¤¤¤", InfoUsuario_T.Index_IU)
            DoEvents
        End If
    Else
        Call FrmLog.AgregarLog("Atención!!!no hay usuarios conectados!!!")
    End If
    Exit Sub
End Sub

Private Sub Cmd_Conv_Click()
'///cuando damos doble clic en el arbol '''sirve para obtener datos especificos de un usuario y poder converesar con el
    ObtenerInformacionTvUsr
    DoEvents
    'MsgBox InfoUsuario_T.Cuenta_IU & InfoUsuario_T.Maquina_IU & InfoUsuario_T.Index_IU
    If InfoUsuario_T.Index_IU > 0 Then
        If ExisteVentanaII(InfoUsuario_T.Cuenta_IU) = False Then
            Call NuevaVentana(InfoUsuario_T.Index_IU, InfoUsuario_T.Nivel_IU, InfoUsuario_T.Cuenta_IU, 0, 3, TxtUsuario.Text, "°°°Conversación aceptada°°°", "Administrador", InfoUsuario_T.NivelS_IU)
            DoEvents
        End If
    Else
        Call FrmLog.AgregarLog("Atención!!! no hay usuarios conectados")
    End If
    Exit Sub
End Sub

Private Sub Cmd_IG_Click()
Dim Temp1 As String
Temp1 = TVUsuarios.SelectedItem.Text
If Temp1 = "Administradores" Or Temp1 = "Profesores" Or Temp1 = "Alumnos" Then Exit Sub

Dim itmFound As ListItem

Set itmFound = LstVwUsrIn. _
   FindItem(Temp1, lvwText, , 1)

   If itmFound Is Nothing Then   ' Si no hay coincidencia, informa al
      Call FrmLog.AgregarLog("Atención!!! Usuario no registrado")
      Exit Sub
   Else
        itmFound.EnsureVisible    ' Desplaza ListView para mostrar el                                     ' ListItem hallado.
        itmFound.Selected = True   ' Selecciona el ListItem.
        LstVwUsrIn.SetFocus
   End If
End Sub

Private Sub Cmd_IP_Click()
'///cuando damos doble clic en el arbol '''sirve para obtener datos especificos de un usuario y poder converesar con el
    ObtenerInformacionTvUsr
    If InfoUsuario_T.Index_IU > 0 Then
        FrmInfP.CuentaInfP = InfoUsuario_T.Cuenta_IU
        FrmInfP.MaquinaInfP = InfoUsuario_T.Maquina_IU
        FrmInfP.Show
        DoEvents
    End If
End Sub

Private Sub Cmd_TF_Click()
    ObtenerInformacionTvUsr

    If InfoUsuario_T.Index_IU >= 0 Then
        If Conectado(InfoUsuario_T.Index_IU) Then
            Call Enviar("REMOTEPHOTO¯[®©]¤¤¤", InfoUsuario_T.Index_IU)
            DoEvents
        End If
    Else
        Call FrmLog.AgregarLog("Atención!!! no hay usuarios conectados")
    End If
End Sub

Private Sub Cmd_Cerrar_T_Click()
''''aqui botamos a todos
    Dim i5 As Integer
    For i5 = 1 To W_Usr.UBound
        If W_Usr(i5).State = sckConnected And Left(W_Usr(i5).Tag, 8) = "ACCESADO" Then
            W_Usr(i5).SendData ("AVISO¯[®©]¤¤¤" & "El Centro de Computo se cerrara en 2 minutos:Guarda tus Trabajos!!!")
            '''///aqui ponemos el tiempo de espera
            ''//// CUANDO PASE  APAGAMOS TODAS LAS TERMINALES
        End If
        DoEvents
    Next i5
End Sub

Private Sub CmdAceptar_Click()
    If Preguntar("Los datos son correctos?") = False Then Exit Sub
    ConexionConfiguracion
    SqlCCon = "select * from Tbl_Config where Clave='Config'"
    RsCCon.Open SqlCCon, Conecta, adOpenStatic, adLockOptimistic
    If Not RsCCon.EOF Then
        RsCCon!Dias_Eliminar = Val(CboHistorial.Text)
        RsCCon!Max_Amonestaciones = Val(CboAmonestaciones.Text)
        RsCCon!Max_T_Reporte = Val(CboReportes.Text)
        RsCCon!Ayuda_Linea = ChkAyuda.Value
        RsCCon!Contestadora = ChkContesta.Value
        RsCCon!Mensaje = TxtMensaje.Text
        RsCCon!Bloquear_P = ChkBP.Value
        RsCCon!AA = ChkAA.Value
        RsCCon!AP = ChkAP.Value
        RsCCon!PP = ChkPP.Value
        RsCCon.Update
        DoEvents
    End If
    If Preguntar("¿Deseas enviar la nueva configuración a los usuarios?") = False Then Exit Sub
    Dim IConfig As Integer
    For IConfig = 1 To W_Usr.UBound
        If W_Usr(IConfig).State = sckConnected And Left(W_Usr(IConfig).Tag, 8) = "ACCESADO" Then
            Call EnviarConfiguracionUsr(IConfig, Mid$(W_Usr(IConfig).Tag, 10))
            DoEvents
        End If
        DoEvents
    Next
    CboReportes.SetFocus
End Sub

Private Function ConvertirChk(Chk As CheckBox) As Integer
    If Chk.Value = 1 Then
        ConvertirChk = -1
    Else
        ConvertirChk = 0
    End If
End Function

Private Sub CmdMasivo_Click()
    If W_Usr.UBound > 0 Then
        Call Masivo("COMANDOAPI¯[®©]¤¤¤" & LblcmdOpcion.Caption, 0)
        Call FrmLog.AgregarLog("Opción : " & LblOptCaption.Caption & " enviada satisfactoriamente!!!")
    Else
        Call FrmLog.AgregarLog("Opción no enviada (No hay usuarios conectados) ..." & vbNewLine & TxtNotificacion.Text)
    End If
    DoEvents
End Sub

Private Sub CmdMsjMasivo_Click()
    If TxtNotificacion.Text = "" Then Exit Sub
    If W_Usr.UBound > 0 Then
        Call Masivo("NOTIFICACION¯[®©]¤¤¤" & TxtNotificacion.Text, 0)
        Call FrmLog.AgregarLog("Mensaje a todos enviado satisfactoriamente a todos los usuarios..." & vbNewLine & TxtNotificacion.Text)
    Else
        Call FrmLog.AgregarLog("Mensaje no enviado (No hay usuarios conectados) ..." & vbNewLine & TxtNotificacion.Text)
    End If
    DoEvents
End Sub

Private Sub FGP_Click()
    If FGP.MouseRow = 0 Then Exit Sub
    TxtUC.Text = FGP.TextMatrix(FGP.MouseRow, 0)
    TxtUM.Text = FGP.TextMatrix(FGP.MouseRow, 1)
    TxtUF.Text = FGP.TextMatrix(FGP.MouseRow, 2)
    TxtUH.Text = FGP.TextMatrix(FGP.MouseRow, 3)
    TxtUP.Text = FGP.TextMatrix(FGP.MouseRow, 4)
End Sub

Private Sub FGR_Click()
    If FGR.MouseRow = 0 Then Exit Sub
    TxtRU.Text = FGR.TextMatrix(FGR.MouseRow, 0)
    TxtRM.Text = FGR.TextMatrix(FGR.MouseRow, 1)
    TxtRF.Text = FGR.TextMatrix(FGR.MouseRow, 2)
    TxtRH.Text = FGR.TextMatrix(FGR.MouseRow, 3)
    TxtRT.Text = FGR.TextMatrix(FGR.MouseRow, 4)
    TxtRR.Text = FGR.TextMatrix(FGR.MouseRow, 5)
End Sub

Private Sub Form_Activate()
    Dim Ctrl As Control
    Me.ZOrder
    For Each Ctrl In Controls
        If Not TypeOf Ctrl Is CommonDialog _
            And Not TypeOf Ctrl Is ImageList _
            And Not TypeOf Ctrl Is Timer _
            And Not TypeOf Ctrl Is Data _
            And Not TypeOf Ctrl Is Winsock _
            And Not TypeOf Ctrl Is PictureBox _
            And Not TypeOf Ctrl Is Menu Then
            Ctrl.ZOrder
        End If
        DoEvents
    Next
End Sub

Private Sub Form_Load()
    '''guardamos el directorio de la base de datos
    '''cargamos el historial general y revisamos que no haya campos vacios
    
    hSysMenu = GetSystemMenu(Me.hwnd, False)
    RemoveMenu hSysMenu, SC_CLOSE, MF_BYCOMMAND
    MDIPrincipal.AgregarVentana Me, "Principal", "Ventana principal"
    
    Redondear Me
    
    TxtFecha.Text = Date
    TxtHora.Text = Time
    TxtUsuario.Text = "FELIPE"
      
    Call CargarArbol
    Call CargarConfiguracion
    Call CargarLstVL
    Call CargarReportes
    Call AgregarMaquinasLVM
    Call CargarEscuela
    
    LstMsjEspera.Clear
    LstPuerto.Clear
    
    FGP.TextMatrix(0, 0) = "Usuario"
    FGP.TextMatrix(0, 1) = "Maquina"
    FGP.TextMatrix(0, 2) = "Fecha"
    FGP.TextMatrix(0, 3) = "Hora"
    FGP.TextMatrix(0, 4) = "Proceso"
    FGP.ColWidth(0) = 800
    FGP.ColWidth(1) = 800
    FGP.ColWidth(2) = 800
    FGP.ColWidth(3) = 800
    FGP.ColWidth(4) = 4000 - 175
    'inicializamos el sender del hostname
    With WUPD
        On Error Resume Next
        .Protocol = sckUDPProtocol
        .LocalPort = 7201
        .RemotePort = 7200
        .RemoteHost = "255.255.255.255"
        .SendData ""
        DoEvents
    End With
    'inicializamos los winsocks para los clientes
    W_Usr(0).Close
    W_Usr(0).LocalPort = 1257
    W_Usr(0).Listen
    
    PosicionInicial Me
End Sub

Private Sub Form_Terminate()
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    CerrarServidor
    MDIPrincipal.RemoverVentana Me, "Principal"
End Sub

Private Sub CerrarServidor()
    ''''aqui botamos a todos
    Dim i5 As Long
    Dim v_o As Integer
    
    Call MDIPrincipal.ChequeoGral(0)
    
    For i5& = 1 To W_Usr.UBound
        W_Usr(i5).Close
        DoEvents
    Next i5&
    
    If v_o > 0 Then
        For v_o = 1 To N_Ventanas
            Unload PMensaje(v_o)
        Next v_o
    End If
    
    N_Ventanas = 0
    End
End Sub

Private Sub LVM_Click()
    'MsgBox LVM.SelectedItem.Key
    'ConexionConsultasM
    'SqlCCM = "select * from Tbl_Maquina"
End Sub

Private Sub Option1_Click()

End Sub

Private Sub OptAAtx_Click()
    LblcmdOpcion = "APAGARATX"
    LblOptCaption.Caption = OptAAtx.Caption
End Sub

Private Sub OptApagar_Click()
    LblcmdOpcion = "APAGAR"
    LblOptCaption.Caption = OptReiniciar.Caption
End Sub

Private Sub Optbloq_Click()
    LblcmdOpcion = "BLOQUEAR"
    LblOptCaption.Caption = Optbloq.Caption
End Sub

Private Sub OptCSesion_Click()
    LblcmdOpcion = "CERRARSESION"
    LblOptCaption.Caption = OptCSesion.Caption
End Sub

Private Sub OptDsbloq_Click()
    LblcmdOpcion = "DESBLOQUEAR"
    LblOptCaption.Caption = OptDsbloq.Caption
End Sub

Private Sub OptReg_Click()
    LblcmdOpcion = "TERMINARREG"
    LblOptCaption.Caption = OptReg.Caption
End Sub

Private Sub OptReiniciar_Click()
    LblcmdOpcion = "REINICIAR"
    LblOptCaption.Caption = OptReiniciar.Caption
End Sub

Private Sub OptTerminar_Click()
    LblcmdOpcion = "TERMINARPROG"
    LblOptCaption.Caption = OptTerminar.Caption
End Sub

Private Sub Tmr_Hora_Timer()
    'Actualiza la hora cada segundo
    TxtHora.Text = Time
End Sub

Private Sub TmrMonitor_Timer()
    Do While LstMsjEspera.ListCount > 0
        If Conectado(LstPuerto.List(0)) = False Then Exit Sub
        W_Usr(LstPuerto.List(0)).SendData (LstMsjEspera.List(0) & "¤¦[MYF]¦¤" & vbCrLf & "¤¦[MYF]¦¤")
        LstMsjEspera.RemoveItem (0):        LstPuerto.RemoveItem (0)
    Loop
End Sub

Private Sub TVUsuarios_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        If TVUsuarios.SelectedItem.Text = "Administradores" Or _
        TVUsuarios.SelectedItem.Text = "Profesores" Or _
        TVUsuarios.SelectedItem.Text = "Alumnos" Then
            Exit Sub
        Else
            PopupMenu Mnu_OpUsr
        End If
    End If
End Sub

Private Sub W_Usr_DataArrival(Index As Integer, ByVal bytesTotal As Long)
On Error GoTo Problema
Dim Datos As String
Dim Dat As String 'Cadena binaria (Imagen)

W_Usr(Index).GetData Datos, vbString

If InStr(1, Datos, ("¤¦[MYF]¦¤" & vbCrLf & "¤¦[MYF]¦¤")) <> 0 Then
    Do While InStr(1, Datos, ("¤¦[MYF]¦¤" & vbCrLf & "¤¦[MYF]¦¤"))
        Lstrecibidos.AddItem Mid(Datos, 1, InStr(1, Datos, ("¤¦[MYF]¦¤" & vbCrLf & "¤¦[MYF]¦¤")) - 1)
        Datos = Mid(Datos, InStr(1, Datos, ("¤¦[MYF]¦¤" & vbCrLf & "¤¦[MYF]¦¤")) + Len("¤¦[MYF]¦¤" & vbCrLf & "¤¦[MYF]¦¤"), Len(Datos))
    Loop
    
    If Len(Lstrecibidos.List(0)) = 0 Then
        Lstrecibidos.RemoveItem (0)
    End If
    
    Do While Lstrecibidos.ListCount > 0
        AnalizarDatosS Index
    Loop
    
Else
    Dat = Datos
    If (Mid(Dat, 1, 8)) = "PETICION" Then
    Pa = App.Path & "\Imagenes"
        If DirExist(Pa) = True Then
            Dim temparray() As String
            Dim fName As String
            Dim fsize As Double
            temparray = Split(Dat, "|")
            fName = temparray(1)
            fsize = temparray(2)
            If Len(Pa) = 3 Then Pa = Mid(Pa, 1, 2)
            Pa = Pa & "\"
            fPath = Pa & Format(Now, "dd-mm-yyyy hh_mm AMPM") & fName
            Close #2
            Open fPath For Binary Access Write As #2
            W_Usr(Index).SendData "ok|" & fName & "|" & fsize & "|"
            DoEvents
            DoEvents
        End If
        Exit Sub
    End If
    
    If Right(Dat, 17) = "******FINAL******" Then
        Dim FinaArchivo As String
        FinaArchivo = Mid(Dat, 1, Len(Dat) - 17)
        Put #2, , FinaArchivo
        Close #2
        ShellExecute MDIPrincipal.hwnd, "", fPath, "", Pa, 0
        Call FrmLog.AgregarLog("Imagen: " & fPath)
        fPath = ""
        Pa = ""
        DoEvents
        DoEvents
        Exit Sub
    End If
    
    Put #2, , Dat
    
End If
Problema:
    If Err.Number = 0 Then
        Exit Sub
    End If
End Sub

'****************************Si alguien sale del sistema
Private Sub W_Usr_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    CerrarConexion (Index)
End Sub

Private Sub W_Usr_Close(Index As Integer)
    CerrarConexion (Index)
End Sub

Private Sub CerrarConexion(Index As Integer)
On Error GoTo Problema
    '/// si algun usuario se desconecta del modo seguro
    Dim N_Cliente As Integer
    Dim N_Maquina As String
    Dim Temp1 As String
    Dim Temp2 As Integer
    Dim Temp3 As Integer
    Dim temp4 As String
    Dim Posicion1 As Integer
    Dim TempRV2 As Integer
    
    If Conectado(Index) = True Then
        W_Usr(Index).Close
    End If
        W_Usr(Index).Tag = ""
        
    Temp2 = TVUsuarios.Nodes.Count
    If Temp2 < 4 Then Exit Sub
    For I3 = 4 To Temp2
        Temp1 = TVUsuarios.Nodes(I3).Key
        If InStr(1, Temp1, ":") Then
            Posicion1 = InStr(1, Temp1, ":")
            N_Maquina = Mid$(Temp1, 1, Posicion1 - 1)
            temp4 = Mid$(Temp1, Posicion1 + 1, InStr(Posicion1 + 1, Temp1, ":"))
            Posicion1 = InStr(1, temp4, ":")
            N_Cliente = Mid$(temp4, 1, Posicion1 - 1)
            If N_Cliente = Index Then
                Llave = TVUsuarios.Nodes(I3).Text
                Call Usr_Salida(Llave, N_Maquina) 'registramos su salida
                Call Masivo(("SALIDA¯[®©]¤¤¤" & Llave), N_Cliente) 'les decimos quien salio
                Call TVUsuarios.Nodes.Remove(I3)
                Call UnlUsr(Llave)
                Call CerrarVentana(Llave)
                Call FrmLog.AgregarLog(Llave & " Finzalizo sesión!!!")
                Call InicioSesion(CStr(Llave & Chr(10) & "Finalizo Sesión" & Chr(10) & "Maquina: " & N_Maquina))
                Call AgregarMaquinasLVM
                DoEvents
                Exit Sub
            End If
        End If
        DoEvents
    Next I3
Problema:
End Sub

'****************************Fin Si alguien sale del sistema
'///////aqui va la prog. del winsock
Private Sub W_Usr_ConnectionRequest(Index As Integer, ByVal requestID As Long)
    Dim i As Integer
    If Index = 0 Then
    
        If Usuario > 0 Then
            For i = 1 To Usuario 'W_Usr.UBound
            'MsgBox W_Usr(i).State & Chr(10) & W_Usr(i).Tag & Chr(10) & Usuario
                If Left(W_Usr(i).Tag, 8) <> "ACCESADO" Then
                    W_Usr(i).Close
                    W_Usr(i).LocalPort = 1258
                    W_Usr(i).Accept requestID
                    Exit Sub
                End If
                DoEvents
            Next i
            
        End If
        Usuario = Usuario + 1
        Load W_Usr(Usuario)
        W_Usr(Usuario).LocalPort = 0
        W_Usr(Usuario).Accept requestID
        DoEvents
        
    End If
'MsgBox W_Usr.UBound
End Sub

'//////terminan funciones
'///aqui van las funciones independientes

Private Sub Usr_Salida(Usr_S As String, N_Maq As String)
''///cuando un usuario cierra sesion, reinicia o apaga la maquina de modo seguro
'desocupamos la maquina que tenia el usuario
    DataModif.DatabaseName = Direccion
    DataModif.RecordSource = "SELECT * FROM [Tbl_Maquina]"
    DataModif.Connect = ";Pwd=" & C_BD
    Call DataModif.Refresh
    DataModif.Recordset.MoveLast
    DataModif.Recordset.MoveFirst
    
    If DataModif.Recordset.RecordCount = 0 Then
        Call FrmLog.AgregarLog("Atención!!!" & vbNewLine & "Base de Datos Alterada, el registro de actividades del usuario: " & Usr_S & " en la Maquina: " & N_Maq & "se ha perdido ")
    Else
        DataModif.Recordset.FindFirst "C_Acceso='" & Usr_S & "' and C_Maq='" & N_Maq & "' and ((Maq_Inicio)=true) and ((Maq_Fin)=false) "
        DataModif.Recordset.Edit
        DataModif.Recordset("Maq_Ocupada") = False
        DataModif.Recordset("Maq_Inicio") = False
        DataModif.Recordset("Maq_Fin") = False
        DataModif.Recordset("C_Acceso") = ""
        DataModif.Recordset.Update
    End If
'terminamos el historial
    DataModif.DatabaseName = Direccion
    DataModif.RecordSource = "SELECT * FROM [Tbl_Historial]"
    DataModif.Connect = ";Pwd=" & C_BD
    Call DataModif.Refresh
    
    DataModif.Recordset.MoveLast
    DataModif.Recordset.MoveFirst
    
    If DataModif.Recordset.RecordCount = 0 Then
        Call FrmLog.AgregarLog("Atención!!!" & vbNewLine & "Base de Datos Alterada, el registro de actividades del usuario: " & Usr_S & " en la Maquina: " & N_Maq & " se toma!!!")
    Else
        DataModif.Recordset.FindFirst "C_Acceso='" & Usr_S & "' and C_Maq='" & N_Maq & "'  and ((Hora_Entrada)Is Not Null) and ((Hora_Salida)Is  Null) "
        DataModif.Recordset.Edit
        DataModif.Recordset("Hora_Salida") = TxtHora.Text
        DataModif.Recordset.Update
    End If
End Sub



Private Sub CargarArbol()
    '/// cargamos el arbol con los elementos raiz
    TVUsuarios.Nodes.Add , , "Admin", "Administradores", "Admin"
    TVUsuarios.Nodes.Add , , "Profr", "Profesores", "Profr"
    TVUsuarios.Nodes.Add , , "Alumno", "Alumnos", "Alumno"
End Sub

Public Sub Enviar(Texto As String, PuertoEnviar As Integer)
    '///aqui enviamos datos a determinado usuario
    If Texto = "" Or PuertoEnviar = 0 Then Exit Sub
    LstMsjEspera.AddItem Texto
    LstPuerto.AddItem PuertoEnviar
End Sub

Public Sub Masivo(Texto As String, puerto As Integer)
    ''aqui enviamos datos a todos los usuarios ecepto a nosotros
    'MsgBox Texto & Chr(10) & Puerto
    For I10 = 1 To W_Usr.UBound
        If W_Usr(I10).State = sckConnected And Left(W_Usr(I10).Tag, 8) = "ACCESADO" Then
            If I10 <> puerto Then
                Mensaje = (Texto)
                Call Enviar(Mensaje, I10)
                Mensaje = ""
            End If
        End If
        DoEvents
    Next I10
End Sub

Public Function Conectado(puerto As Integer) As Boolean
    '' aqui checamos si los usuarios estan conectados
    If W_Usr(puerto).State <> sckClosed Then
        Conectado = True
    Else
        Conectado = False
   End If
End Function

Private Function ExisteVentana(CMUsuario As String, FraseUSRMsj As String) As Boolean
    If CMUsuario = "" Or FraseUSRMsj = "" Then ExisteVentana = True: Exit Function
    If LstVO.ListCount = 0 Then ExisteVentana = False
    Dim IEV As Integer
    For IEV = 0 To LstVO.ListCount - 1
        If FraseUSRMsj <> "" Then
            If PMensaje(LstVO.List(IEV)).Tag = CMUsuario Then PMensaje(LstVO.List(IEV)).Txt_Respuesta.Text = PMensaje(LstVO.List(IEV)).Txt_Respuesta.Text _
            & CMUsuario & ":" & FraseUSRMsj & Chr(10): ExisteVentana = True: Exit Function
        End If
    Next
    ExisteVentana = False
End Function

Public Sub NuevaVentana(PuertoA As Integer, NivelA As Integer, CuentaA As String, PuertoB As Integer, NivelB As Integer, CuentaB As String, FrasePM As String, Caption1 As String, Caption2 As String)
    If LstVL.ListCount = 0 Then Exit Sub
    Dim VX As Integer
    VX = LstVL.List(0)
    ReDim Preserve PMensaje(VX)
    PMensaje(VX).Tag = CuentaA
    PMensaje(VX).Puerto1 = PuertoA
    PMensaje(VX).Puerto2 = PuertoB
    PMensaje(VX).Nivel1 = NivelA
    PMensaje(VX).Nivel2 = NivelB
    PMensaje(VX).Cuenta1 = CuentaA
    PMensaje(VX).Cuenta2 = CuentaB
    PMensaje(VX).NVIndex = VX
    PMensaje(VX).Caption = "::: Ayuda Directa ::: " & Caption1 & ": " & CuentaB & " | " & Caption2 & ": " & CuentaA & " :::"
    PMensaje(VX).Txt_Respuesta.Text = CuentaA & ":" & FrasePM & Chr(10)
    PMensaje(VX).Visible = True
    LstVO.AddItem (VX)
    LstVL.RemoveItem (0)
End Sub

Private Sub AnalizarDatosS(IndexC As Integer)
Dim i9 As Integer
Dim TxtDigerido As String
Dim Opcionc As String
Dim DatosC As String
Dim Posicion1 As Integer
Dim AccesadoC As Boolean
Dim RemotePort As Integer
Dim RemoteIP As String
Dim MUP As String
Dim TMUP As String
Dim Mbox As Boolean

    TxtDigerido = Lstrecibidos.List(0)
    Posicion1 = InStr(1, TxtDigerido, "¯[®©]¤¤¤")
    Opcionc = Mid$(TxtDigerido, 1, Posicion1 - 1)
    DatosC = Mid$(TxtDigerido, Posicion1 + 8)
    
    If Left(W_Usr(IndexC).Tag, 8) = "ACCESADO" Then
        AccesadoC = True
    Else
        AccesadoC = False
    End If
    
    If Opcionc = "NUEVOUSUARIO" Then
    
        Dim NuevoUsuario() As String
        Dim NombreDU As String
        Dim ClaveDU As String
        Dim PasswordDU As String
       
        NuevoUsuario = Split(DatosC, "¤¢©§¦[BOLA]¦§©¢¤", , vbTextCompare)
        NombreDU = NuevoUsuario(0)
        ClaveDU = NuevoUsuario(1)
        PasswordDU = NuevoUsuario(2)
         
        ConexionNuevoUsuario
        SqlCNU = "Select * from Tbl_Acceso where C_Acceso='" & ClaveDU & "'"
        RsCNU.Open SqlCNU, Conecta, adOpenDynamic, adLockBatchOptimistic
        
        If Not RsCNU.EOF Then
            Mensaje = "NUEVOUSUARIOCREADO¯[®©]¤¤¤CLAVEREPETIDA"
            Call Enviar(Mensaje, IndexC)
        Else
            ConexionNuevoUsuario
            SqlCNU = "Select * from Tbl_Acceso where Nombre='" & NombreDU & "'"
            RsCNU.Open SqlCNU, Conecta, adOpenStatic, adLockOptimistic

            If Not RsCNU.EOF Then
                If Left(RsCNU!C_Acceso, 2) = "T-" And RsCNU!Fecha_Reg = CDate("01/01/1900") Then
                    RsCNU!C_Acceso = ClaveDU
                    RsCNU!Password = PasswordDU
                    RsCNU.Update
                    Mensaje = "NUEVOUSUARIOCREADO¯[®©]¤¤¤SI"
                    Call Enviar(Mensaje, IndexC)
                    Call FrmLog.AgregarLog("Nuevo usuario: " & ClaveDU & vbNewLine & "Nombre: " & NombreDU)
                Else
                    Mensaje = "NUEVOUSUARIOCREADO¯[®©]¤¤¤YAREGISTRADO"
                    Call Enviar(Mensaje, IndexC)
                End If
            Else
                Mensaje = "NUEVOUSUARIOCREADO¯[®©]¤¤¤NOENCONTRADO"
                Call Enviar(Mensaje, IndexC)
            End If
        End If
        GoTo Salida
    End If
    
 'contador para el bucle del envio de la lista de conectados
    If Opcionc = "LOGIN" Then
        If AccesadoC = False Then
            Posicion2 = InStr(1, DatosC, "¤¢©§¦[BOLA]¦§©¢¤")
            Usr_Cuenta = Mid$(DatosC, 1, Posicion2 - 1)
            Datos3 = Mid$(DatosC, Posicion2 + 16)
            Posicion2 = InStr(1, Datos3, "¤¢©§¦[BOLA]¦§©¢¤")
            Usr_Password = Mid$(Datos3, 1, Posicion2 - 1)
            Usr_Maquina = Mid$(Datos3, Posicion2 + 16)
            'MsgBox Usr_Cuenta & Chr(10) & Usr_Password & Chr(10) & Usr_Maquina
            
            If ValidarAcceso(Usr_Cuenta, Usr_Password, Usr_Maquina, IndexC) = True Then
                If Left(W_Usr(IndexC).Tag, 8) = "ACCESADO" Then
                    Call WhoEnter(IndexC, Usr_Cuenta)
                    For i9 = 4 To TVUsuarios.Nodes.Count
                        If TVUsuarios.Nodes(i9).Text <> Usr_Cuenta Then
                            Mensaje = ("AGREGAR¯[®©]¤¤¤" & TVUsuarios.Nodes(i9).Key)
                            Call Enviar(Mensaje, IndexC)
                        End If
                        DoEvents
                    Next
                End If
                RemotePort = W_Usr(IndexC).RemotePort
                RemoteIP = W_Usr(IndexC).RemoteHostIP
                Call EnviarFecha(IndexC)
                Call EnviarProcesos(IndexC)
                Call AgregarLstVw(Usr_Cuenta, Usr_Maquina, RemotePort, RemoteIP, Time())
                Call InicioSesion(CStr(Usr_Cuenta & Chr(10) & "Acaba de Iniciar Sesión"))
                Call FrmLog.AgregarLog(Usr_Cuenta & " Inició sesión!!!")
                Call AgregarMaquinasLVM
                GoTo Salida
            Else
                W_Usr(IndexC).Tag = ""
                GoTo Salida
            End If
        End If
    End If
    
    If Opcionc = "AMONESTACIONUSR" Then
        Call AgregarAmonestacion(DatosC, IndexC)
        GoTo Salida
    End If
    
    If Opcionc = "PROCESORA" Then
        AgregarFilaRPA (DatosC)
        GoTo Salida
    End If
    
    If Opcionc = "AYUDAR" Then
        '///cuando un usuario solicita conversacion con otro
        LstMensajes.AddItem DatosC
        Call MandarMensaje
        GoTo Salida
    End If
   
    If Opcionc = "ESC" Then
        Dim EscCuenta As String
        Dim EscOpcion As Integer
        Dim EscPuerto As Integer
        Dim EscPos As Integer
        Dim EscI As Integer
             
        EscPos = InStr(1, DatosC, "¯-_[††]_-¯")
        EscCuenta = Mid(DatosC, 1, EscPos - 1)
        EscPuerto = Mid(DatosC, EscPos + 10)
        EscOpcion = Right(EscCuenta, 1)
        EscCuenta = Mid(EscCuenta, 1, Len(EscCuenta) - 1)
        
        If EscPuerto > 0 Then
            Call Enviar(TxtDigerido, EscPuerto)
            GoTo Salida
        End If
        
        If LstVO.ListCount = 0 Then GoTo Salida
        For EscI = 0 To LstVO.ListCount - 1
            If PMensaje(LstVO.List(EscI)).Tag = EscCuenta Then
                If EscOpcion = 1 Then
                    PMensaje(LstVO.List(EscI)).SBEM.Panels(1).Text = EscCuenta & " esta escribiendo un mensaje!!!"
                    GoTo Salida
                Else
                    PMensaje(LstVO.List(EscI)).SBEM.Panels(1).Text = ""
                    GoTo Salida
                End If
            End If
        Next
        GoTo Salida
    End If

    If Opcionc = "REPORTE" Then
        Dim NReporte() As String
        Dim NTR As String
        Dim NMR As String
        Dim NCR As String
        Dim NMsj As String
        NReporte = Split(DatosC, "¤¢©§¦[BOLA]¦§©¢¤", , vbTextCompare)
        NCR = NReporte(0)
        NMR = NReporte(1)
        NTR = NReporte(2)
        NMsj = NReporte(3)
        Call GuardarReportes(NCR, NMR, NTR, NMsj)
        Call AgregarFilaRR(NCR, NMR, NTR, NMsj)
        
        Call FrmLog.AgregarLog("Nuevo Reporte: " & NTR & vbNewLine & _
              "Maquina: " & NMR & vbNewLine & _
              "Usuario: " & NCR & vbNewLine & _
              "Descripción: " & NMsj)
        
        Erase NReporte
        GoTo Salida
    End If
    Exit Sub
    
Salida:
    If Lstrecibidos.ListCount > 0 Then Lstrecibidos.RemoveItem (0)
End Sub

Public Function ValidarAcceso(Cuenta As String, Clave As String, Maquina As String, Usr_Index As Integer) As Boolean
Dim Problema As Boolean 'si hay problema en el acceso
Dim Motivo As String 'motivo del problema
Dim Usr_Nivel As Integer 'nivel jerarquico del usuario
Dim I2 As Integer ''contador para el blucle que nos dice si esta cuenta ya esta siendo utilizada o no

Problema = False

With DataLogin
    .DatabaseName = Direccion
    .RecordSource = "Select * from Tbl_Acceso"
    .Connect = ";Pwd=" & C_BD
    Call .Refresh
    If .Recordset.RecordCount = 0 Then
        'si es el primer usuario en la base de datos
        Call FrmLog.AgregarLog("Atención!!!" & vbNewLine & "Un usuario esta intentando conectarse, pero la base de datos de los usuarios está vacia" & Chr(10) & "Se recomienda registrar a los usuarios!!!")
        Motivo = "NOACCESO¯[®©]¤¤¤Tu cuenta no esta registrada!!!" & Chr(10) & "Consulta al Administrador" & "¤¢©§¦[BOLA]¦§©¢¤Acceso Denegado!!!"
        Problema = True
        Call FrmLog.AgregarLog("La base de datos de los usuarios está vacía!!!")
        GoTo Usr_Problema
    Else
        .Recordset.FindFirst "C_Acceso='" & Cuenta & "'"
        If .Recordset("C_Acceso") <> Cuenta Then
            'si no existe el usuario!!
            contador = contador + 1
            
            If contador = 3 Then
                Motivo = "NOACCESO¯[®©]¤¤¤Tu cuenta no esta registrada o esta mal escrita!!!" & Chr(10) & "Consulta al Administrador" & "¤¢©§¦[BOLA]¦§©¢¤Acceso Denegado!!!"
                Problema = True
                Call FrmLog.AgregarLog("Fallo acceso: " & Cuenta)
                GoTo Usr_Problema
            End If
            
            Motivo = "NOACCESO¯[®©]¤¤¤Cuenta o Password de usuario incorrecto, intenta otra vez!!!¤¢©§¦[BOLA]¦§©¢¤Acceso Denegado!!!"
            Problema = True
            Call FrmLog.AgregarLog("Acceso denegado: " & Cuenta)
            GoTo Usr_Problema
        End If
        
        If .Recordset("Password") <> Clave Then
            contador = contador + 1
            If contador = 3 Then
                Motivo = "NOACCESO¯[®©]¤¤¤Tu cuenta no esta registrada o esta mal escrita!!!" & Chr(10) & "Consulta al Administrador" & "¤¢©§¦[BOLA]¦§©¢¤Acceso Denegado!!!"
                Problema = True
                Call FrmLog.AgregarLog("Fallo acceso: " & Cuenta)
                GoTo Usr_Problema
            End If
            
            Motivo = "NOACCESO¯[®©]¤¤¤Cuenta o Password de usuario incorrecto, intenta otra vez!!!¤¢©§¦[BOLA]¦§©¢¤Acceso Denegado!!!"
            Problema = True
            Call FrmLog.AgregarLog("Acceso denegado: " & Cuenta)
            GoTo Usr_Problema
            
        End If
        
        If .Recordset("Usr_Bloqueado") = True Then
            'si el usuario esta bloqueado
            Motivo = "NOACCESO¯[®©]¤¤¤Tu cuenta ha sido bloqueada!!!" & Chr(10) & "Consulta al Administrador" & "¤¢©§¦[BOLA]¦§©¢¤Acceso Denegado!!!"
            Problema = True
            Call FrmLog.AgregarLog("Usuario bloqueado: " & Cuenta)
            GoTo Usr_Problema
        End If
        
        Usr_Nivel = .Recordset("Nivel")
        
    End If
    
End With

With DataMaq

        .DatabaseName = Direccion
        .RecordSource = "Select * from Tbl_Maquina"
        .Connect = ";Pwd=" & C_BD
        Call .Refresh
        
        If .Recordset.RecordCount = 0 Then
            'si es el primer usuario en la base de datos
            Motivo = "NOACCESO¯[®©]¤¤¤Maquina no registrada!!!" & Chr(10) & "Consulta al Administrador" & "¤¢©§¦[BOLA]¦§©¢¤Acceso Denegado!!!"
            Problema = True
            Call FrmLog.AgregarLog("Maquina no registrada: " & Maquina)
            GoTo Usr_Problema
        Else
        
            .Recordset.FindFirst "C_Maq='" & Maquina & "'"
            If .Recordset("C_Maq") <> Maquina Then
                'si la maquina no es correcta
                Motivo = "NOACCESO¯[®©]¤¤¤Clave de Maquina Incorrecta!!!" & Chr(10) & "Consulta al Administrador" & "¤¢©§¦[BOLA]¦§©¢¤Acceso Denegado!!!"
                Problema = True
                Call FrmLog.AgregarLog("Maquina incorrecta: " & Maquina)
                GoTo Usr_Problema
            End If
            
            If .Recordset("Maq_Bloqueada") = True Then
                'si la maquina esta bloqueada
                Motivo = "NOACCESO¯[®©]¤¤¤Maquina Bloqueada Temporalmente!!!" & Chr(10) & "Consulta al Administrador" & "¤¢©§¦[BOLA]¦§©¢¤Acceso Denegado!!!"
                Problema = True
                Call FrmLog.AgregarLog("Maquina Bloqueada Temporalmente: " & Maquina)
                GoTo Usr_Problema
            End If
            
         
            If .Recordset("Maq_Ocupada") = True Then
                'si esta maquina ya esta registrada
                Motivo = "NOACCESO¯[®©]¤¤¤Maquina ya registrada!!!" & Chr(10) & "Consulta al Administrador" & "¤¢©§¦[BOLA]¦§©¢¤Acceso Denegado!!!"
                Problema = True
                Call FrmLog.AgregarLog("Esta maquina ya esta en uso: " & Maquina)
                GoTo Usr_Problema
            End If
            
            'checar si el usuario no es el amdministrador
            If Cuenta = TxtUsuario.Text Then
                Motivo = "NOACCESO¯[®©]¤¤¤Esta cuenta ya esta siendo utilizada!!!" & Chr(10) & "Consulta al Administrador" & "¤¢©§¦[BOLA]¦§©¢¤Acceso Denegado!!!"
                Problema = True
                Call FrmLog.AgregarLog("Este usuario ya a iniciado sesión (Administrador): " & Cuenta)
                GoTo Usr_Problema
                DoEvents
            End If
                
            '''''Agregamos el usuario al arbol
            For I2 = 4 To TVUsuarios.Nodes.Count
                If TVUsuarios.Nodes(I2).Text = Cuenta Then
                    Motivo = "NOACCESO¯[®©]¤¤¤Esta cuenta ya esta siendo utilizada!!!" & Chr(10) & "Consulta al Administrador" & "¤¢©§¦[BOLA]¦§©¢¤Acceso Denegado!!!"
                    Problema = True
                    Call FrmLog.AgregarLog("Este usuario ya a iniciado sesión: " & Cuenta)
                    GoTo Usr_Problema
                End If
                DoEvents
            Next I2
            
            If Usr_Nivel = 1 Then
                TVUsuarios.Nodes.Add "Alumno", tvwChild, Maquina & ":" & Usr_Index & ":" & Usr_Nivel & ":" & Cuenta, Cuenta, "Alumno"
            ElseIf Usr_Nivel = 2 Then
                TVUsuarios.Nodes.Add "Profr", tvwChild, Maquina & ":" & Usr_Index & ":" & Usr_Nivel & ":" & Cuenta, Cuenta, "Profr"
            ElseIf Usr_Nivel = 3 Then
                TVUsuarios.Nodes.Add "Admin", tvwChild, Maquina & ":" & Usr_Index & ":" & Usr_Nivel & ":" & Cuenta, Cuenta, "Admin"
            Else
                Motivo = "NOACCESO¯[®©]¤¤¤Error Inesperado!!!" & Chr(10) & "Consulta al Administrador" & "¤¢©§¦[BOLA]¦§©¢¤Acceso Denegado!!!"
                Problema = True
                Call FrmLog.AgregarLog("Error desconocido contacta a tu Proveedor: " & Maquina)
                GoTo Usr_Problema
            End If
            
        End If

End With

If Problema = False Then

    For I2 = 1 To W_Usr.UBound
        If W_Usr(I2).State = sckConnected And Left(W_Usr(I2).Tag, 8) = "ACCESADO" Then
            If I2 <> Usr_Index Then
                Mensaje = ("AGREGAR¯[®©]¤¤¤" & Maquina & "¤¢©§¦[BOLA]¦§©¢¤" & Usr_Index & "¤¢©§¦[BOLA]¦§©¢¤" & Usr_Nivel & "¤¢©§¦[BOLA]¦§©¢¤" & Cuenta)
                Call Enviar(Mensaje, I2)
                Mensaje = ""
            End If
        End If
        DoEvents
    Next I2
    
    With DataModif
        ''''//// aqui se hacen las modificaciones correspondientes al usuario actual
        '////modificamos la maquina con los datos de usuario actual
        .DatabaseName = Direccion
        .RecordSource = "SELECT * FROM [Tbl_Maquina]"
        .Connect = ";Pwd=" & C_BD
        Call .Refresh
        If .Recordset.RecordCount <> 0 Then
            .Recordset.FindFirst "C_Maq='" & Maquina & "'"
            .Recordset.Edit
            .Recordset("Maq_Ocupada") = True
            .Recordset("Maq_Inicio") = True
            .Recordset("Maq_Fin") = False
            .Recordset("C_Acceso") = Cuenta
            .Recordset.Update
        End If
    End With

        '//////agregamos al historial la entrada del usuario en la maquina correspondiente
    With DataModif
        .DatabaseName = Direccion
        .RecordSource = "SELECT * FROM [Tbl_Historial]"
        .Connect = ";Pwd=" & C_BD
        Call .Refresh
        
        If .Recordset.RecordCount = 0 Then
            .Recordset.AddNew
            .Recordset("C_Acceso") = Cuenta
            .Recordset("Fecha_Entrada") = TxtFecha.Text
            .Recordset("Hora_Entrada") = TxtHora.Text
            .Recordset("Hora_Salida") = Null
            .Recordset("C_Maq") = Maquina
            .Recordset.Update
        Else
            .Recordset.FindFirst "C_Acceso='" & Cuenta & "'"
            .Recordset.AddNew
            .Recordset("C_Acceso") = Cuenta
            .Recordset("Fecha_Entrada") = TxtFecha.Text
            .Recordset("Hora_Entrada") = TxtHora.Text
            .Recordset("Hora_Salida") = Null
            .Recordset("C_Maq") = Maquina
            .Recordset.Update
        End If
    End With

    Call EnviarConfiguracionUsr(Usr_Index, Cuenta)
    DoEvents
    Mensaje = ("ACCESO1¯[®©]¤¤¤Bienvenido al sistema " & Cuenta & " !!!¤¢©§¦[BOLA]¦§©¢¤Acceso Correcto!!!¤¢©§¦[BOLA]¦§©¢¤" & TxtUsuario.Text & "¤¢©§¦[BOLA]¦§©¢¤" & Usr_Nivel & "¤¢©§¦[BOLA]¦§©¢¤" & Usr_Index & "¤¢©§¦[BOLA]¦§©¢¤" & T_Reportes & "¤¢©§¦[BOLA]¦§©¢¤" & HU_Entrada)
    Call Enviar(Mensaje, Usr_Index)
    DoEvents
Else

    Motivo = "Error al Procesar tus datos!!!" & Chr(10) & "Consulta a tu Administrador!!!:Acceso Denegado!!!"
    GoTo Usr_Problema
End If

W_Usr(Usr_Index).Tag = "ACCESADO-" & Cuenta
ValidarAcceso = True

Exit Function

Usr_Problema:

        Call Enviar(Motivo, Usr_Index)
        DoEvents
        DoEvents
        ValidarAcceso = False
End Function
''terminan funciones independienes
Private Sub EnviarConfiguracionUsr(Index As Integer, Cuenta As String)
Repetir:
    With DataModif
        .DatabaseName = Direccion
        .RecordSource = "SELECT * FROM [Tbl_Config]"
        .Connect = ";Pwd=" & C_BD
        Call .Refresh
        If .Recordset.RecordCount > 0 Then
            Dim Combinacion As String
            T_Reportes = Val(CboReportes.Text)
            Combinacion = ChkAyuda.Value & "¤¢©§¦[BOLA]¦§©¢¤" & ChkBP.Value & "¤¢©§¦[BOLA]¦§©¢¤" & ChkAA.Value & "¤¢©§¦[BOLA]¦§©¢¤" & ChkAP.Value & "¤¢©§¦[BOLA]¦§©¢¤" & ChkPP.Value & "¤¢©§¦[BOLA]¦§©¢¤"
            Combinacion = Combinacion & Val(CboAmonestaciones.Text) & "¤¢©§¦[BOLA]¦§©¢¤" & Val(CboReportes) & "¤¢©§¦[BOLA]¦§©¢¤"
            ConexionPrincipal
            Sql = "select * from Tbl_Acceso where C_Acceso='" & Cuenta & "'"
            Rs.Open Sql, Conecta, adOpenStatic, adLockOptimistic
            If Not Rs.EOF Then
                Combinacion = Combinacion & Rs!Amonestaciones & "¤¢©§¦[BOLA]¦§©¢¤"
                DoEvents
            End If
            Rs.Close
            Combinacion = "CONFIGURACION¯[®©]¤¤¤" & Combinacion
            Call Enviar(Combinacion, Index)
            DoEvents
        Else
            .Recordset.AddNew
            .Recordset("Clave") = "Config"
            .Recordset.Update
            Call FrmLog.AgregarLog("Atención!!!" & vbNewLine & "Base de datos Alterada!!! >>> Tabla: Tbl_Config" & vbNewLine _
                   & "Se tomarán los datos predeterminados!!!"): DoEvents
            GoTo Repetir
        End If
    End With
End Sub

Public Function SQLDate(ConvertDate As Date) As String
    SQLDate = Format(ConvertDate, "mm/dd/yyyy")
End Function

Private Sub W_Usr_SendComplete(Index As Integer)
    Ocupado = False
End Sub

Private Sub AgregarLstVw(AgUsr As String, AgMaq As String, AgPuerto As Integer, AgIp As String, AgTime As Date)
    Dim LV_TempUsr As ListItem
    Set LV_TempUsr = LstVwUsrIn.ListItems.Add(, , AgUsr)
        LV_TempUsr.SubItems(1) = AgMaq
        LV_TempUsr.SubItems(2) = AgPuerto
        LV_TempUsr.SubItems(3) = AgIp
        LV_TempUsr.SubItems(4) = AgTime
End Sub

Private Sub UnlUsr(UnlUsrIn As String)
    Dim itmFound As ListItem

    Set itmFound = LstVwUsrIn. _
    FindItem(UnlUsrIn, lvwText, , 1)

    If itmFound Is Nothing Then   ' Si no hay coincidencia, informa al                                     ' usuario y sale.
        Call FrmLog.AgregarLog("Cerrar sesión usuario: " & UnlUsrIn & " [Atención!!! Usuario no econtrado]")
        Exit Sub
    Else
        itmFound.EnsureVisible    ' Desplaza ListView para mostrar el                                     ' ListItem hallado.
        itmFound.Selected = True   ' Selecciona el ListItem.
        LstVwUsrIn.SetFocus
    End If
    LstVwUsrIn.ListItems.Remove (LstVwUsrIn.SelectedItem.Index)
End Sub

Private Sub ObtenerInformacionTvUsr()
    Dim Temp1 As String
    Dim Temp2 As String

    Temp1 = ""
    Temp2 = ""

    InfoUsuario_T.Maquina_IU = ""
    InfoUsuario_T.Index_IU = 0
    InfoUsuario_T.Nivel_IU = 0
    InfoUsuario_T.Cuenta_IU = ""
    InfoUsuario_T.NivelS_IU = ""
    
    Temp1 = TVUsuarios.SelectedItem.Key
    'If TVUsuarios.SelectedItem.Text = TxtUsuario Then Exit Sub
    If Temp1 = "Admin" Or Temp1 = "Profr" Or Temp1 = "Alumno" Then Exit Sub
    Posicion2 = InStr(1, Temp1, ":")
    Usr_M1 = Mid(Temp1, 1, Posicion2 - 1)
    Temp1 = Mid(Temp1, Posicion2 + 1)
    Posicion2 = InStr(1, Temp1, ":")
    Usr_I1 = Mid(Temp1, 1, Posicion2 - 1)
    Usr_N1 = Mid(Temp1, Posicion2 + 1, 1)
    Temp1 = Mid(Temp1, Posicion2 + 1)
    Posicion2 = InStr(1, Temp1, ":")
    Usr_C1 = Mid(Temp1, Posicion2 + 1)
    
    If Usr_N1 = 1 Then
        Temp2 = "Alumno"
    ElseIf Usr_N1 = 2 Then
        Temp2 = "Profesor"
    ElseIf Usr_N1 = 3 Then
        Temp2 = "Administrador"
    End If
    
    InfoUsuario_T.Maquina_IU = Usr_M1
    InfoUsuario_T.Index_IU = Usr_I1
    InfoUsuario_T.Nivel_IU = Usr_N1
    InfoUsuario_T.Cuenta_IU = Usr_C1
    InfoUsuario_T.NivelS_IU = Temp2
End Sub

Private Sub WhoEnter(IndexC As Integer, WhosClave As String)
    Dim DB As Database
    Dim Rs As Recordset
    Dim Busca As String
    
    Busca = "SELECT * from Tbl_Acceso where C_Acceso=" & "'" + WhosClave + "'"
    Set DB = DBEngine.Workspaces(0).OpenDatabase(Direccion, False, False, ";Pwd=" & C_BD)
    Set Rs = DB.OpenRecordset(Busca, dbOpenDynaset)
    
    If Not Rs.BOF Or Not Rs.EOF Then
        Mensaje = ("NOMBREBOLA¯[®©]¤¤¤" & Rs!Nombre)
        Call Enviar(Mensaje, IndexC)
        DoEvents
    End If
End Sub

Public Function EnviarProcesos(IndexC As Integer)
    Dim StrProcesos As String
    ConexionPrincipal
    Sql = "select * from Tbl_Procesos"
    Rs.Open Sql, Conecta, adOpenDynamic, adLockBatchOptimistic
    StrProcesos = ""
    
    Do While Not Rs.EOF
        StrProcesos = StrProcesos & Rs!Programa & "¤¢©§¦[BOLA]¦§©¢¤"
        Rs.MoveNext
    Loop
    Rs.Close
    Call Enviar("PROCESOSR¯[®©]¤¤¤" & StrProcesos, IndexC)
End Function

Public Function EnviarFecha(IndexC As Integer)
    Call Enviar("FECHAR¯[®©]¤¤¤" & Date, IndexC)
End Function

Private Sub CargarConfiguracion()
    Dim C1 As Integer
    Dim Palabra
    
    CboReportes.Clear
    CboHistorial.Clear
    CboAmonestaciones.Clear
    
    For C1 = 1 To 15
        CboReportes.AddItem C1
        CboHistorial.AddItem C1
        CboAmonestaciones.AddItem C1
    Next
    
    With DataModif
        .DatabaseName = Direccion
        .RecordSource = "SELECT * FROM [Tbl_Config]"
        .Connect = ";Pwd=" & C_BD
        Call .Refresh
        If .Recordset.RecordCount = 0 Then
            .Recordset.AddNew
            .Recordset("Clave") = "Config"
            .Recordset.Update
            Call FrmLog.AgregarLog("Atención!!!" & vbNewLine & "Base de datos Alterada!!! >>> Tabla: Tbl_Config" & vbNewLine _
                   & "Se tomarán los datos predeterminados!!!")
        End If
    End With
    
    ConexionPrincipal
    Sql = "select * from Tbl_Config where Clave='Config'"
    Rs.Open Sql, Conecta, adOpenDynamic, adLockBatchOptimistic
    
    If Not Rs.EOF Then
        CboReportes.Text = Rs!Max_T_Reporte
        CboHistorial.Text = Rs!Dias_Eliminar
        CboAmonestaciones.Text = Rs!Max_Amonestaciones
        If Rs!AA = True Then
            ChkAA.Value = 1
        Else
            ChkAA.Value = 0
        End If
        If Rs!AP = True Then
            ChkAP.Value = 1
        Else
            ChkAP.Value = 0
        End If
        If Rs!PP = True Then
            ChkPP.Value = 1
        Else
            ChkPP.Value = 0
        End If
        
        If Rs!Ayuda_Linea = True Then
            ChkAyuda.Value = 1
        Else
            ChkAyuda.Value = 0
        End If
        
        If Rs!Contestadora = True Then
            ChkContesta.Value = 1
        Else
            ChkContesta.Value = 0
        End If
        
        TxtMensaje.Text = Rs!Mensaje
        
        If Rs!Bloquear_P = True Then
            ChkBP.Value = 1
        Else
            ChkBP.Value = 0
        End If
    End If
    Rs.Close
End Sub

Private Sub WUPD_DataArrival(ByVal bytesTotal As Long)
    On Error Resume Next
    Dim IdHost As String
    Dim IdHostFinal As String
    
    WUPD.GetData IdHost
    
    If InStr(1, IdHost, "ID¯[®©]¤¤¤[") Then
        IdHostFinal = Mid(IdHost, InStr(1, IdHost, "ID¯[®©]¤¤¤") + 4)
        IdHostFinal = Mid(IdHostFinal, 1, InStr(1, IdHostFinal, "]") - 1)
        If IdHostFinal = "SOLICITUD" Then
            WUPD.SendData "ID¯[®©]¤¤¤[" & W_Usr(0).LocalHostName & "]"
        End If
    ElseIf InStr(1, IdHost, "bola") = 0 Then
        Call FrmLog.AgregarLog("Atención!!!" & vbNewLine & "Alguien ajeno a la institución o alguien no registrado esta tratando accesar a la red!!!" & vbNewLine _
        & "Host remoto: " & WUPD.RemoteHost & vbNewLine _
         & "IP remota: " & WUPD.RemoteHostIP & vbNewLine _
         & "Puerto: " & WUPD.LocalPort)
    End If
End Sub

Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then
        Visible = False
    End If
End Sub

Private Sub AgregarFilaRPA(DatosC2 As String)
    Dim CRAT() As String
    Dim IRP As Integer
    CRAT = Split(DatosC2, "¤¢©§¦[BOLA]¦§©¢¤", , vbTextCompare)
    With FGP
        .Rows = .Rows + 1
        .Row = .Rows - 1
        .TextMatrix(.Row, 0) = CRAT(0)
        .TextMatrix(.Row, 1) = CRAT(1)
        .TextMatrix(.Row, 2) = Date
        .TextMatrix(.Row, 3) = Time
        .TextMatrix(.Row, 4) = CRAT(2)
         For IRP = 0 To 4
            .Col = IRP
            If .Row / 2 <> Int(.Row / 2) Then
                .CellBackColor = RGB(194, 208, 252)
            End If
        Next
        .TopRow = .Row
        .FixedRows = 1
    End With
    Call GuardarProcesos(CRAT(0), CRAT(1), CRAT(2))
    Erase CRAT
End Sub

Private Sub GuardarProcesos(CU_P As String, CM_P As String, CP_P As String)
    If Len(CP_P) > 250 Then CP_P = Mid$(CP_P, 1, 249)
    ConexionProcesos
    SqlCP = "insert into Tbl_Procesos_Reg (C_Acceso,C_Maq,Fecha,Hora,Proceso) " _
    & "values('" & CU_P & "','" & CM_P & "','" & Date & "','" & Time & "','" & CP_P & "')"
    RsCP.Open SqlCP, Conecta, adOpenDynamic, adLockBatchOptimistic
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
            Call Enviar("AYUDAR¯[®©]¤¤¤" & LstMensajes.List(0), UsrE_I2)
        Else
            If ChkAyuda.Value = 1 Then
                If ChkContesta.Value = 0 Then
                    Mensaje = ("AYUDAR¯[®©]¤¤¤" & UsrE_I1 & "¤¢©§¦[BOLA]¦§©¢¤" & UsrE_N1 & "¤¢©§¦[BOLA]¦§©¢¤" & UsrE_C1 & "¤¢©§¦[BOLA]¦§©¢¤" & UsrE_I2 & "¤¢©§¦[BOLA]¦§©¢¤" & UsrE_N2 & "¤¢©§¦[BOLA]¦§©¢¤" & UsrE_C2 & "¤¢©§¦[BOLA]¦§©¢¤" & TxtMensaje.Text)
                    Call Enviar(Mensaje, UsrE_I1)
                    DoEvents
                Else
                    If ExisteVentana(UsrE_C1, UsrE_F1) = False Then
                        Call NuevaVentana(UsrE_I1, UsrE_N1, UsrE_C1, UsrE_I2, UsrE_N2, UsrE_C2, UsrE_F1, ParseLevel(UsrE_N2), ParseLevel(UsrE_N1))
                        DoEvents
                    End If
                End If
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
    For IVl = 1 To 100
        LstVL.AddItem IVl
    Next
End Sub

Public Sub QuitarVentana(VRI As Integer)
    If LstVO.ListCount = 0 Then Exit Sub
    Dim IRV As Integer
    For IRV = 0 To LstVO.ListCount - 1
        If LstVO.List(IRV) = VRI Then LstVL.AddItem VRI: LstVO.RemoveItem (IRV): Exit Sub
    Next
End Sub

Private Sub CerrarVentana(CVUsuario As String)
    Dim ICV As Integer
    For ICV = 0 To LstVO.ListCount - 1
        If PMensaje(LstVO.List(ICV)).Tag = CVUsuario Then Unload PMensaje(LstVO.List(ICV)): Exit Sub
    Next
End Sub

Private Function ExisteVentanaII(CMUsuario As String) As Boolean
    If LstVO.ListCount = 0 Then ExisteVentanaII = False
    Dim IEVII As Integer
    For IEVII = 0 To LstVO.ListCount - 1
        If PMensaje(LstVO.List(IEVII)).Tag = CMUsuario Then ExisteVentanaII = True: Exit Function
    Next
    ExisteVentanaII = False
End Function




Private Sub CargarReportes()
    Dim Titulos() As Variant
    Dim i As Integer
    Dim Ancho As Long
    Titulos = Array("Usuario", "Maquina", "Fecha", "Hora", "Titulo", "Reporte")
    FGR.Row = 0
    For i = 0 To 5
        FGR.Col = i
        FGR.ColAlignment(i) = flexAlignLeftCenter
        FGR.Text = Titulos(i)
        FGR.ColWidth(i) = CInt(TextWidth(FGR.Text) + 100) + 400
        Ancho = Ancho + FGR.ColWidth(i)
    Next
    FGR.ColWidth(5) = 4000 - 175
End Sub

Private Sub AgregarFilaRR(A1 As String, B1 As String, C1 As String, D1 As String)
    Dim IRP As Integer
    With FGR
        .Rows = .Rows + 1
        .Row = .Rows - 1
        .TextMatrix(.Row, 0) = A1
        .TextMatrix(.Row, 1) = B1
        .TextMatrix(.Row, 2) = Date
        .TextMatrix(.Row, 3) = Time
        .TextMatrix(.Row, 4) = C1
        .TextMatrix(.Row, 5) = D1
         For IRP = 0 To 5
            .Col = IRP
            If .Row / 2 <> Int(.Row / 2) Then
                .CellBackColor = RGB(194, 208, 252)
            End If
        Next
        .TopRow = .Row
        .FixedRows = 1
    End With
End Sub

Private Sub GuardarReportes(RP_C As String, RP_M As String, RP_T As String, RP_Msj As String)
    If Len(RP_Msj) > 250 Then RP_Msj = Mid$(RP_Msj, 1, 249)
    ConexionProcesos
    SqlCP = "insert into Tbl_Reportes (C_Acceso,C_Maq,Fecha_Reporte,Hora_Reporte,Titulo,Reporte) " _
    & "values('" & RP_C & "','" & RP_M & "','" & Date & "','" & Time & "','" & RP_T & "','" & RP_Msj & "')"
    RsCP.Open SqlCP, Conecta, adOpenDynamic, adLockBatchOptimistic
End Sub

Private Sub BuscarDisponibles(Bloqueada As Integer, Ocupada As Integer, Index As Integer)
    ConexionDisponibles
    SqlCD = "select *  from Tbl_Maquina where Maq_Bloqueada =" & Bloqueada & " and Maq_Ocupada=" & Ocupada
    RsCD.Open SqlCD, Conecta, adOpenDynamic, adLockBatchOptimistic
    Do While Not RsCD.EOF
        LVM.ListItems.Add , Index & RsCD!C_Maq, RsCD!C_Maq, Index
        LVM.Refresh
        RsCD.MoveNext
    Loop
End Sub

Public Function AgregarMaquinasLVM()
    LVM.ListItems.Clear
    Call BuscarDisponibles(0, 0, 1)
    Call BuscarDisponibles(0, -1, 2)
    Call BuscarDisponibles(-1, 1, 3)
    Call BuscarDisponibles(-1, 0, 3)
End Function

Private Sub CargarEscuela()
    ConexionEscuelas
    SQLCE = "select * from Tbl_Escuela where IDESC = 'CBT2'"
    RsCE.Open SQLCE, Conecta, adOpenDynamic, adLockBatchOptimistic
    If RsCE.EOF = True Then
        RsCE.Close
        SQLCE = "insert into Tbl_Escuela ([IDESC],[C_Escuela],[Escuela],[Domicilio]) VALUES ('CBT2','CECControl','By Felipe Zaldivar de la O','http://www.paginasprodigy.com/zaldivar610818')"
        RsCE.Open SQLCE, Conecta, adOpenDynamic, adLockBatchOptimistic
    End If
End Sub

Private Sub AgregarAmonestacion(UsuarioAmo As String, IUsuarioAmo As Integer)
    ConexionA
    SqlCA = "Select * From Tbl_Acceso Where C_Acceso='" & UsuarioAmo & "'"
    RsCA.Open SqlCA, Conecta, adOpenStatic, adLockPessimistic
    
    If Not RsCA.EOF Then
        RsCA!Amonestaciones = RsCA!Amonestaciones + 1
        If RsCA!Amonestaciones > Val(CboAmonestaciones.Text) Then
            RsCA!Usr_Bloqueado = -1
            RsCA!Amonestaciones = 0
            RsCA!N_Bloqueos = RsCA!N_Bloqueos + 1
            Call Enviar("BLOQUEARUSUARIO¯[®©]¤¤¤", IUsuarioAmo)
        End If
        RsCA.Update
        RsCA.Close
    End If
End Sub
