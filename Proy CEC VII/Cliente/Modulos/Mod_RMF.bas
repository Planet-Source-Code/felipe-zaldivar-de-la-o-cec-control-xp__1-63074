Attribute VB_Name = "Mod_RMF"
Option Explicit

Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long
Public Declare Function sndPlaySound Lib "WINMM.DLL" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Const conSwNormal = 1

Public Const SPI_GETWORKAREA = 48
Public Const SPI_SCREENSAVERRUNNING = 97

Public Const SND_SYNC = &H0
Public Const SND_ASYNC = &H1
Public Const SND_NODEFAULT = &H2
Public Const SND_LOOP = &H8
Public Const SND_NOSTOP = &H10

Public Const HWND_TOPMOST = -1
Public Const HWND_NOTTOPMOST = -2
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_SHOWWINDOW = &H40

Public Type RECT
    PIzq As Long
    PDer As Long
    PArr As Long
    PAbj As Long
End Type

Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const REG_SZ = 1
Public Const REG_DWORD = 4
Public Const ERROR_SUCCESS = 0&

Public Sub MaxTop(VentanaHwnd As Long, Optional MaxT As Boolean = True)
    Dim Bandera As Integer
    Dim OpTop As String
    
    Bandera = SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
    
    If MaxT Then
        OpTop = SetWindowPos(VentanaHwnd, HWND_TOPMOST, 0, 0, 0, 0, Bandera)
    Else
        OpTop = SetWindowPos(VentanaHwnd, HWND_NOTTOPMOST, 0, 0, 0, 0, Bandera)
    End If
End Sub
