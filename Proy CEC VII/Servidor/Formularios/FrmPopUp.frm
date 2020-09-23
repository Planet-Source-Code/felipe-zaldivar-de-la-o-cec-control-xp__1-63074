VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Begin VB.Form FrmPopUp 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   2040
   ClientLeft      =   5730
   ClientTop       =   4845
   ClientWidth     =   2775
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2040
   ScaleWidth      =   2775
   ShowInTaskbar   =   0   'False
   Begin VB.Timer TmrAbrir 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   2160
      Top             =   1200
   End
   Begin VB.Timer TmrEspera 
      Enabled         =   0   'False
      Interval        =   4000
      Left            =   2160
      Top             =   600
   End
   Begin VB.Timer TmrCerrar 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   2160
      Top             =   0
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash ShockwaveFlash1 
      Height          =   1785
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2685
      _cx             =   4736
      _cy             =   3149
      FlashVars       =   ""
      Movie           =   ""
      Src             =   ""
      WMode           =   "Transparent"
      Play            =   -1  'True
      Loop            =   -1  'True
      Quality         =   "High"
      SAlign          =   ""
      Menu            =   0   'False
      Base            =   ""
      AllowScriptAccess=   "always"
      Scale           =   "ExactFit"
      DeviceFont      =   0   'False
      EmbedMovie      =   0   'False
      BGColor         =   ""
      SWRemote        =   ""
      MovieData       =   ""
      SeamlessTabbing =   -1  'True
      Profile         =   0   'False
      ProfileAddress  =   ""
      ProfilePort     =   0
   End
   Begin VB.Image Image1 
      Height          =   180
      Left            =   1680
      MouseIcon       =   "FrmPopUp.frx":0000
      MousePointer    =   99  'Custom
      Picture         =   "FrmPopUp.frx":0152
      Top             =   120
      Width           =   195
   End
   Begin VB.Label LblUsrInicio 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Programado por:  Felipe Zaldivar de la O"
      Height          =   795
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   1815
   End
   Begin VB.Image ImgBack 
      Height          =   1755
      Left            =   0
      Picture         =   "FrmPopUp.frx":0568
      Top             =   0
      Width           =   2055
   End
End
Attribute VB_Name = "FrmPopUp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public MiID As Integer

Dim resto As Long
Dim taskbar As Long
Dim HeightInit As Long

Private Sub Image1_Click()
    Unload Me
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image1.Picture = LoadPicture(App.Path + "\Graficos\x2.jpg")
End Sub

Private Sub ImgBack_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image1.Picture = LoadPicture(App.Path + "\Graficos\x1.jpg")
End Sub

Public Function Iniciar(Texto As String, Index As Integer)
    Dim WindowRect As Rect
    Dim Bandera As Long, yo As Long
    
    ShockwaveFlash1.Height = ImgBack.Height
    ShockwaveFlash1.Width = ImgBack.Width
    ShockwaveFlash1.Movie = App.Path & "\Graficos\PopUp.swf"
    ShockwaveFlash1.Play
    
    Me.Height = ShockwaveFlash1.Height
    Me.Width = ShockwaveFlash1.Width
    HeightInit = Me.Height
    
    SystemParametersInfo SPI_GETWORKAREA, 0, WindowRect, 0
    taskbar = ((Screen.Height / Screen.TwipsPerPixelX) - WindowRect.PAbj) * Screen.TwipsPerPixelX
    
    Me.Left = Screen.Width - Me.ScaleWidth
    Me.Top = Screen.Height
    resto = Me.Top - ((Me.Height * (Index)) + taskbar)
    Me.Top = resto + Me.ScaleHeight
    Me.Height = 0
    LblUsrInicio = Texto
    
    Me.Hide
    Me.Show
    
    Bandera = SND_ASYNC Or SND_NODEFAULT
    yo = sndPlaySound(App.Path & "\Graficos\Inicio.wav", Bandera)
    
    Call ImgBack.Refresh
    
    MaxTop Me.hwnd, True
    
    TmrAbrir.Enabled = True
End Function

Private Sub TmrAbrir_Timer()
    If Me.Height > HeightInit Then
        Me.Height = HeightInit
        Me.Top = resto
        TmrAbrir.Enabled = False
        TmrEspera.Enabled = True
    Else
        Me.Top = Me.Top - 30
        Me.Height = Me.Height + 30
    End If
End Sub

Private Sub TmrCerrar_Timer()
    If (Me.Height >= HeightInit Or Me.Height <= HeightInit) And Me.Height > 0 And Me.Height >= 30 Then
        Me.Height = Me.Height - 30
        Me.Top = Me.Top + 30
    Else
        TmrCerrar.Enabled = False
        NoPopup = NoPopup - 1
        Unload Me
    End If
End Sub

Private Sub TmrEspera_Timer()
    TmrAbrir.Enabled = False
    TmrCerrar.Enabled = True
End Sub

