Attribute VB_Name = "Mod_Variables"
Option Explicit

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function GetSystemMenu Lib "user32" (ByVal hWnd As Long, ByVal bRevert As Long) As Long
Public Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function GetDesktopWindow Lib "user32" () As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long
Public Declare Function sndPlaySound Lib "WINMM.DLL" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

Public Const SPI_GETWORKAREA = 48
Public Const SC_CLOSE = &HF060&
Public Const MF_BYCOMMAND = &H0&

Public Const SND_SYNC = &H0
Public Const SND_ASYNC = &H1
Public Const SND_NODEFAULT = &H2
Public Const SND_LOOP = &H8
Public Const SND_NOSTOP = &H10

Public Const HWND_TOPMOST = -1
Public Const HWND_NOTTOPMOST = -2
Public Const SWP_NOSIZE = 1
Public Const SWP_NOMOVE = 2

Public Type Rect
    PIzq As Long
    PDer As Long
    PArr As Long
    PAbj As Long
End Type

Public Type POINTAPI
    X As Long
    Y As Long
End Type

Public PMensaje() As New FrmAyudar
Public PopUp(1 To 15) As FrmPopUp '////Ventanas de aviso de sesion
Global N_Ventanas As Integer
Global Mensaje As String '///guarda las cadenas de texto a enviar al usuario
Global Frase As String
Global Llave As String ''///guarda usuario a desconectar
Global T_Reportes As Integer
Global HU_Entrada As Date
Public Direccion As String ''guarda temporalmente la direccion de la base de datos
Public C_BD As String ''guarda la clave de la base de datos
Public Primera As Boolean '' permite el bucle para validar la base de datos
Public NoPopup As Integer
Public NoPopup2 As Integer

Public Sub InicioSesion(MensajeUsrIni As String)
    If NoPopup >= 5 Or NoPopup < 0 Then NoPopup = 0
    NoPopup = NoPopup + 1
    Set PopUp(NoPopup) = New FrmPopUp
    Call PopUp(NoPopup).Iniciar(MensajeUsrIni, NoPopup)
    PopUp(NoPopup).MiID = NoPopup
    NoPopup2 = Screen.Height / FrmPopUp.ImgBack.Height
End Sub

Public Sub MaxTop(VentanaHwnd As Long, Optional MaxT As Boolean = True)
    Dim Bandera As Integer
    Dim OpTop As String
    Bandera = SWP_NOSIZE Or SWP_NOMOVE
    If MaxT Then
        OpTop = SetWindowPos(VentanaHwnd, HWND_TOPMOST, 0, 0, 0, 0, Bandera)
    Else
        OpTop = SetWindowPos(VentanaHwnd, HWND_NOTTOPMOST, 0, 0, 0, 0, Bandera)
    End If
End Sub
