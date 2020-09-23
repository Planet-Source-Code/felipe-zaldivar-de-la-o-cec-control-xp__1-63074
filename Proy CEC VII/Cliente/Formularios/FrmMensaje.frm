VERSION 5.00
Begin VB.Form FrmMensaje 
   BackColor       =   &H00FFCEBB&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "::: MENSAJES DE ULTIMA HORA :::"
   ClientHeight    =   3720
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   6495
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3720
   ScaleWidth      =   6495
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdAceptar 
      BackColor       =   &H00FFCEBB&
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   4680
      TabIndex        =   1
      Top             =   3240
      Width           =   1695
   End
   Begin VB.TextBox TextMensaje 
      Height          =   3015
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   120
      Width           =   6255
   End
End
Attribute VB_Name = "FrmMensaje"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdAceptar_Click()
    Unload Me
End Sub
