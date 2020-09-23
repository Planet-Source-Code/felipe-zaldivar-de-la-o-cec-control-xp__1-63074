VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form FrmAyudar 
   BackColor       =   &H00FFCEBB&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "::: Ayuda Directa ::: Administrador: | Usuario: :::"
   ClientHeight    =   4935
   ClientLeft      =   4170
   ClientTop       =   3375
   ClientWidth     =   6990
   Icon            =   "FrmAyudar.frx":0000
   LinkTopic       =   "FrmPrincipal"
   MaxButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   6990
   Begin VB.TextBox Txt_Frase 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   120
      MaxLength       =   100
      MultiLine       =   -1  'True
      TabIndex        =   2
      Text            =   "FrmAyudar.frx":08CA
      Top             =   3360
      Width           =   5415
   End
   Begin VB.CommandButton Cmd_Enviar 
      BackColor       =   &H00FFCEBB&
      Caption         =   "Enviar"
      Height          =   615
      Left            =   5640
      TabIndex        =   1
      Top             =   3360
      Width           =   1215
   End
   Begin MSComctlLib.StatusBar SBEM 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   4680
      Width           =   6990
      _ExtentX        =   12330
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   12277
            Key             =   "EM"
         EndProperty
      EndProperty
   End
   Begin RichTextLib.RichTextBox Txt_Respuesta 
      Height          =   3135
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   5530
      _Version        =   393217
      BorderStyle     =   0
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      Appearance      =   0
      TextRTF         =   $"FrmAyudar.frx":08D4
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
   Begin VB.Label LblMensajeI 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"FrmAyudar.frx":0960
      Height          =   615
      Left            =   120
      TabIndex        =   4
      Top             =   4080
      Width           =   6735
   End
End
Attribute VB_Name = "FrmAyudar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Puerto1 As Integer
Public Puerto2 As Integer
Public Nivel1 As Integer
Public Nivel2 As Integer
Public Cuenta1 As String
Public Cuenta2 As String
Public NVIndex As Integer

Dim Contador1 As Integer

Private Sub Cmd_Enviar_Click()
    If Txt_Frase.Text <> "" Then
        Call FrmPrincipal.Enviar(("AYUDAR¯[®©]¤¤¤" & Puerto2 & "¤¢©§¦[BOLA]¦§©¢¤" & Nivel2 & "¤¢©§¦[BOLA]¦§©¢¤" & Cuenta2 & "¤¢©§¦[BOLA]¦§©¢¤" & Puerto1 & "¤¢©§¦[BOLA]¦§©¢¤" & Nivel1 & "¤¢©§¦[BOLA]¦§©¢¤" & Cuenta1 & "¤¢©§¦[BOLA]¦§©¢¤" & Txt_Frase.Text))
        Txt_Respuesta.Text = Txt_Respuesta.Text & Cuenta2 & ":" & Txt_Frase.Text & Chr(10): Txt_Frase.Text = ""
    End If
    Txt_Frase.SetFocus
End Sub

Private Sub Form_Activate()
    MaxTop Me.hWnd, True
End Sub

Private Sub Form_Load()
    Txt_Respuesta.Text = ""
    Txt_Frase.Text = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call FrmPrincipal.RemoverVentana(NVIndex)
End Sub

Private Sub Txt_Frase_Change()
    If Contador1 = 0 Then
        If Len(Txt_Frase.Text) = 0 Then
            Contador1 = 1
            Mensaje = "ESC¯[®©]¤¤¤" & Cuenta2 & "0¯-_[††]_-¯" & Puerto1
            Call FrmPrincipal.Enviar(Mensaje)
            DoEvents
            Exit Sub
        Else
            Exit Sub
        End If
    Else
        If Len(Txt_Frase.Text) <> 0 Then
            Contador1 = 0
            Mensaje = "ESC¯[®©]¤¤¤" & Cuenta2 & "1¯-_[††]_-¯" & Puerto1
            Call FrmPrincipal.Enviar(Mensaje)
            DoEvents
            Exit Sub
        End If
        Exit Sub
    End If
End Sub

Private Sub Txt_Frase_KeyPress(KeyAscii As Integer)
    If LTrim(Txt_Frase.Text) <> "" Or KeyAscii <> 255 Then
        If KeyAscii = 13 Then
            Call Cmd_Enviar_Click
            KeyAscii = 0
        End If
    Else
        Exit Sub
    End If
End Sub

Private Sub Txt_Respuesta_Change()
    If Len(Txt_Respuesta.Text) > 15000 Then
        Txt_Respuesta.Text = Right(Txt_Respuesta.Text, 5000)
    End If
    Txt_Respuesta.SelStart = Len(Txt_Respuesta.Text)
End Sub
