VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form FrmReporte 
   BackColor       =   &H00FFCEBB&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "::: Reportes :::"
   ClientHeight    =   4335
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6975
   Icon            =   "FrmReporte.frx":0000
   LinkTopic       =   "FrmPrincipal"
   MaxButton       =   0   'False
   ScaleHeight     =   4335
   ScaleWidth      =   6975
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Txt_Frase 
      Appearance      =   0  'Flat
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
      Height          =   405
      Left            =   120
      MaxLength       =   60
      TabIndex        =   0
      Top             =   360
      Width           =   6735
   End
   Begin VB.CommandButton Cmd_Enviar 
      BackColor       =   &H00FFCEBB&
      Caption         =   "Enviar"
      Height          =   615
      Left            =   5640
      TabIndex        =   2
      Top             =   3600
      Width           =   1215
   End
   Begin RichTextLib.RichTextBox Txt_Respuesta 
      Height          =   2415
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   4260
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   -1  'True
      ScrollBars      =   3
      MaxLength       =   4000
      Appearance      =   0
      TextRTF         =   $"FrmReporte.frx":08CA
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
      Caption         =   $"FrmReporte.frx":094C
      Height          =   615
      Left            =   120
      TabIndex        =   5
      Top             =   3600
      Width           =   5295
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Descripción del Reporte:"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   1755
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Título del Reporte:"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "FrmReporte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Quien_Reporta As String
Public Maquina_Reporte As String

Private Sub Cmd_Enviar_Click()
    If Txt_Frase.Text <> "" And Txt_Respuesta.Text <> "" Then
        Respuesta = MsgBox("¿Los datos son correctos?", 4 + 32 + 0, "Atención!!!")
        If Respuesta = vbYes Then
            Mensaje = ("REPORTE¯[®©]¤¤¤" & Quien_Reporta & "¤¢©§¦[BOLA]¦§©¢¤" & Maquina_Reporte & "¤¢©§¦[BOLA]¦§©¢¤" & Txt_Frase.Text & "¤¢©§¦[BOLA]¦§©¢¤" & Txt_Respuesta.Text)
            FrmPrincipal.Enviar Mensaje
            Txt_Respuesta.Text = ""
            Txt_Frase.Text = ""
            Max_ReportesY = Max_ReportesY + 1
            Unload Me
        Else
            Txt_Frase.SetFocus
        End If
    End If
End Sub

Private Sub Form_Activate()
    MaxTop Me.hWnd, True
End Sub

Private Sub Form_Load()
    Txt_Frase.Text = ""
    Txt_Respuesta.Text = ""
End Sub

