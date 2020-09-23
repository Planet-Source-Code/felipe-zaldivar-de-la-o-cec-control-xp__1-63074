VERSION 5.00
Begin VB.Form FrmIpMaquina 
   BackColor       =   &H00FFCEBB&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "::: Buscando Servidor :::"
   ClientHeight    =   1065
   ClientLeft      =   5775
   ClientTop       =   5445
   ClientWidth     =   4185
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1065
   ScaleWidth      =   4185
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   4320
      Top             =   120
   End
   Begin VB.TextBox TxtMaquina 
      Height          =   375
      Left            =   840
      Locked          =   -1  'True
      TabIndex        =   2
      Text            =   "------------------------------------------------------"
      Top             =   600
      Width           =   3255
   End
   Begin VB.TextBox TxtHostName 
      Height          =   375
      Left            =   840
      Locked          =   -1  'True
      TabIndex        =   0
      Text            =   "------------------------------------------------------"
      Top             =   120
      Width           =   3255
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Maquina:"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   660
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Servidor:"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   630
   End
End
Attribute VB_Name = "FrmIpMaquina"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Contador As Integer

Private Sub Form_Load()
    MaxTop Me.hWnd, True
    
End Sub

Private Sub Timer1_Timer()
    Contador = Contador + 1
    Me.Caption = Contador
    If Contador = 2 Then
        Unload Me
        FrmPrincipal.Enabled = True
    End If
End Sub

