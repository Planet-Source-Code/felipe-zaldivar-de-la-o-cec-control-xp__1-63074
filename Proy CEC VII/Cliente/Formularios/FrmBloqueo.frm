VERSION 5.00
Begin VB.Form FrmBloqueo 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000007&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4275
   ClientLeft      =   3435
   ClientTop       =   3690
   ClientWidth     =   7035
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4275
   ScaleWidth      =   7035
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox PicAccion 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3495
      Left            =   0
      ScaleHeight     =   3495
      ScaleWidth      =   6255
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   6255
      Begin VB.CommandButton Command1 
         Caption         =   "Salir"
         Height          =   255
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MAQUINA BLOQUEADA CONSULTA AL ADMINISTRADOR"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   2760
         TabIndex        =   2
         Top             =   2475
         Width           =   10245
      End
   End
   Begin VB.Timer TmrOnTop 
      Interval        =   1
      Left            =   6480
      Top             =   3720
   End
End
Attribute VB_Name = "FrmBloqueo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public FotoBloqueo As String

Private Sub Command1_Click()
    Call HabilitarRegistro
    Unload Me
End Sub

Private Sub Form_Load()
    Me.WindowState = 2
    Label1.Top = Screen.Height / 2 - Label1.Height / 2
    Label1.Left = Screen.Width / 2 - Label1.Width / 2
    PicAccion.Height = Screen.Height
    PicAccion.Width = Screen.Width
    PicAccion.Picture = LoadPicture(PicFolder & "\" & FotoBloqueo & ".jpg")
    PicAccion.Visible = True
    MaxTop Me.hWnd, True
End Sub

Private Sub TmrOnTop_Timer()
    MaxTop Me.hWnd, True
End Sub


