VERSION 5.00
Begin VB.Form FrmError 
   BackColor       =   &H00FFCEBB&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "::: Error en Tiempo de Ejecución :::"
   ClientHeight    =   4230
   ClientLeft      =   4395
   ClientTop       =   3825
   ClientWidth     =   6735
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmError.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4230
   ScaleWidth      =   6735
   Begin VB.CommandButton CmdImprimir 
      BackColor       =   &H00FFCEBB&
      Caption         =   "&Imprimir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   5
      Top             =   3720
      Width           =   1335
   End
   Begin VB.CommandButton CmdEnviar 
      BackColor       =   &H00FFCEBB&
      Caption         =   "&Enviar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3840
      TabIndex        =   0
      Top             =   3720
      Width           =   1335
   End
   Begin VB.CommandButton CmdCerrar 
      BackColor       =   &H00FFCEBB&
      Caption         =   "&Cerrar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5280
      TabIndex        =   1
      Top             =   3720
      Width           =   1335
   End
   Begin VB.TextBox TxtDescrip 
      BorderStyle     =   0  'None
      Height          =   2535
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   960
      Width           =   6495
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Descripción: "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   1380
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A  ocurrido un error inesperado en Tiempo de Ejecución !!!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   337
      TabIndex        =   3
      Top             =   120
      Width           =   6060
   End
End
Attribute VB_Name = "FrmError"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const Modulo As String = "FrmError"
Dim NChild As Long

Private Sub CmdCerrar_Click()
    FrmPrincipal.Enabled = True
    Unload Me
End Sub

Public Sub ErrorHandler(Numero As Long, Descripcion As String, Fuente As String _
                        , Modulo As String, Evento As String)
                        
    If Err.Number = 0 Then Exit Sub
    
    FrmPrincipal.Enabled = False
    TxtDescrip.Text = ""
    
    TxtDescrip.Text = "Sistema Operativo: " & getVersion & vbNewLine _
                      & "Versión del Programa: " & App.Major & "." & App.Minor & "." & App.Revision & vbNewLine _
                      & "Fecha y Hora del Error: " & Format(Now, "dd/mm/yy hh:mm:ss AMPM") & vbNewLine _
                      & "Número del Error: " & Numero & vbNewLine _
                      & "Descripción del Error: " & Descripcion & vbNewLine _
                      & "Fuente del Error: " & Fuente & vbNewLine _
                      & "Modulo del Error: " & Modulo & vbNewLine _
                      & "Evento del Error: " & Evento & vbNewLine ' _
                      & "Linea del Error: " & Linea
    Err.Clear
    Me.Show
    
End Sub

Private Function EnviarMail(ByVal Mail As String, Optional CCMail As String = "", Optional Titulo As String = " ", Optional Mensaje As String = " ")
    On Error GoTo Error
    Const Evento As String = "EnviarMail"
    
    Const CONST_EMAIL_COMMAND = "MailTo:{Mail}&CC={CCMail}&Subject={Titulo}&Body={Mensaje}"
    Dim RC As Variant
    Dim sCommand As String

    sCommand = "MailTo:" & Mail
    sCommand = sCommand & "?CC=" & CCMail
    sCommand = sCommand & "&Subject=" & Titulo
 
    Mensaje = Replace(Mensaje, vbTab, Space(5))
    Mensaje = Replace(Mensaje, " ", "%20")
    Mensaje = Replace(Mensaje, "&", "%26")
    Mensaje = Replace(Mensaje, vbCrLf, "%0D%0A")
    
    sCommand = sCommand & "&Body=" & Mensaje
    
    RC = ShellExecute(GetDesktopWindow(), "Open", sCommand, "", App.Path, 1)
    
    Exit Function

Error:
    Call FrmError.ErrorHandler(Err.Number, Err.Description, Err.Source, Modulo, Evento)
    Resume Next
    
End Function


Private Sub CmdEnviar_Click()
    EnviarMail "ate_fel@hotmail.com", "ate_fel@yahoo.com.mx", "CEC - CONTROL ERROR", TxtDescrip.Text
End Sub

Private Sub Form_Activate()
    MaxTop Me.hwnd, True
End Sub

Private Sub Form_Load()
    MDIPrincipal.AgregarVentana Me, "Error!!!", "Errores en tiempo de ejecución..."
    PosicionInicial Me
End Sub

Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then
        Visible = False
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MDIPrincipal.RemoverVentana Me, "Error!!!"
End Sub
