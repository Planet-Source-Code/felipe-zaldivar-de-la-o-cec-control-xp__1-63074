VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmPb 
   BackColor       =   &H00FFCEBB&
   BorderStyle     =   0  'None
   ClientHeight    =   3000
   ClientLeft      =   4815
   ClientTop       =   5070
   ClientWidth     =   4695
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   Picture         =   "FrmPb.frx":0000
   ScaleHeight     =   3000
   ScaleWidth      =   4695
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ProgressBar PB 
      Height          =   375
      Left            =   840
      TabIndex        =   0
      Top             =   1320
      Width           =   3090
      _ExtentX        =   5450
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Inicializando"
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   840
      TabIndex        =   1
      Top             =   720
      Width           =   3135
   End
End
Attribute VB_Name = "FrmPb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    PosicionInicial Me
End Sub
