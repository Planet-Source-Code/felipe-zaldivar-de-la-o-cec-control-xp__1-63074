VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form FrmLog 
   BackColor       =   &H00FFCEBB&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "::: LOG :::"
   ClientHeight    =   5550
   ClientLeft      =   3720
   ClientTop       =   3660
   ClientWidth     =   8055
   Icon            =   "FrmBusquedaRep.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5550
   ScaleWidth      =   8055
   Begin VB.TextBox TxtBusqueda 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   120
      MaxLength       =   50
      TabIndex        =   0
      Top             =   4560
      Width           =   7815
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00FFCEBB&
      Caption         =   "&Palabra Exacta"
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   5160
      Width           =   1455
   End
   Begin VB.CommandButton FindButton 
      BackColor       =   &H00FFCEBB&
      Caption         =   "&Buscar"
      Height          =   315
      Left            =   4920
      TabIndex        =   3
      Top             =   5130
      Width           =   1455
   End
   Begin VB.CommandButton CmdSiguiente 
      BackColor       =   &H00FFCEBB&
      Caption         =   "B&uscar Siguiente"
      Height          =   315
      Left            =   6480
      TabIndex        =   4
      Top             =   5130
      Width           =   1455
   End
   Begin VB.CheckBox Check2 
      BackColor       =   &H00FFCEBB&
      Caption         =   "P&alabra Completa ( Caso Exacto )"
      Height          =   315
      Left            =   1680
      TabIndex        =   2
      Top             =   5160
      Width           =   2775
   End
   Begin RichTextLib.RichTextBox RTLog 
      Height          =   4335
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   7646
      _Version        =   393217
      BorderStyle     =   0
      HideSelection   =   0   'False
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"FrmBusquedaRep.frx":08CA
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "FrmLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim DirPath As String
Dim DirLog As String
Dim Posicion As Integer

Private Sub CmdSiguiente_Click()
    Dim FindFlags As Long
    FindFlags = Check1.Value * 4 + Check2.Value * 2
    Posicion = RTLog.Find(TxtBusqueda.Text, Posicion + 1, , FindFlags)
    If Posicion < 0 Then
        MsgBox "La palabra: " & TxtBusqueda.Text & " no se pudo encontrar!!!", , "Atención!!!"
    End If
End Sub

Private Sub FindButton_Click()
    Dim FindFlags As Long
    Posicion = 0
    FindFlags = Check1.Value * 4 + Check2.Value * 2
    Posicion = RTLog.Find(TxtBusqueda.Text, Posicion, , FindFlags)
    If Posicion < 0 Then
        MsgBox "La palabra: " & TxtBusqueda.Text & " no se pudo encontrar!!!", , "Atención!!!"
    End If
End Sub

Private Sub Form_Load()
    MDIPrincipal.AgregarVentana Me, "LOG", "Registro de actividades..."
    DirPath = App.Path & "\LOG"
    DirLog = DirPath & "\" & Replace(Date, "/", "-") & ".txt"
    RTLog.Text = ""
    TxtBusqueda.Text = ""
    Call CargarLog(DirLog)
    PosicionInicial Me
End Sub

Private Sub CargarLog(Archivo As String)
    Dim NArch As Integer
    Dim CadenaLog As String
    
    NArch = FreeFile
    
    If Dir(Archivo) = "" Then
        Exit Sub
    End If

    Open Archivo For Input As NArch
    While Not EOF(NArch)
        Line Input #NArch, CadenaLog
        RTLog.Text = RTLog.Text & CadenaLog & vbNewLine
    Wend
    Close NArch
End Sub

Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then
        Visible = False
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set FrmPrincipal = Nothing
    MDIPrincipal.RemoverVentana Me, "LOG"
End Sub

Public Sub AgregarLog(MsjLog As String)
    If MsjLog = "" Then Exit Sub
    MsjLog = Date & " | " & Time & vbNewLine & MsjLog & vbNewLine
    RTLog.Text = RTLog.Text & MsjLog & vbNewLine
    If DirExist(DirPath) = True Then
        Call GuardarLog(MsjLog)
    End If
End Sub

Private Sub GuardarLog(MsjLogII As String)
    Open DirLog For Append As #1
    Print #1, MsjLogII
    Close #1
End Sub

Private Sub RTLog_Change()
    RTLog.SelStart = Len(RTLog.Text)
End Sub

Private Sub TxtBusqueda_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And TxtBusqueda.Text <> "" Then
        FindButton.SetFocus
    End If
End Sub
