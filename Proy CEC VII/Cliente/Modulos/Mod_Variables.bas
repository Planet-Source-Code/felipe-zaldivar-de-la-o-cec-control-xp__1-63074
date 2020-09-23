Attribute VB_Name = "Mod_Variables"
Option Explicit

Global PMensajeC() As New FrmAyudar  '/////establecemos una matriz de 100 elementos como formulario frmayudar
Global PopUp(1 To 15) As FrmPopUp '////Ventanas de aviso de sesion
Global N_VentanasC As Integer '///guarda el numerode ventanas que vamos a abrir
Global Mensaje As String '///guarda las cadenas de texto a enviar al usuario
Global Frase As String '///guarda la frase que nos enviaron y la que vamos a enviar
Global Admin As String '//guarda el nombre del servidor
Global Mi_Puerto As Integer '' guarda el valor del puerto
Global Usr_Nivel As Integer '' guarda el nivel del usuario en integer
Global Usr_NivelS As String ''guarda el nivel del usuario en string
Global Respuesta As String ''guarda  la respuesta de un comando recibido
Global T_Reportes As Integer
Global Hora_Yo As Date
Global Hora_Dif As Date
Global R_Activo As Boolean
Global Max_ReportesY As Integer
Global Usr_Maquina As String
Global FechaEntrada As Date

Public NoPopup As Integer
Public NoPopup2 As Integer

Public Declare Function CreateRoundRectRgn Lib "gdi32" _
    (ByVal RectX1 As Long, ByVal RectY1 As Long, ByVal RectX2 As Long, _
    ByVal RectY2 As Long, ByVal EllipseWidth As Long, _
    ByVal EllipseHeight As Long) As Long

Public Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Public Declare Sub InitCommonControls Lib "Comctl32.dll" ()

Public Sub InicioSesion(MensajeUsrIni As String)
    If NoPopup >= 5 Or NoPopup < 0 Then NoPopup = 0
    NoPopup = NoPopup + 1
    Set PopUp(NoPopup) = New FrmPopUp
    Call PopUp(NoPopup).Iniciar(MensajeUsrIni, NoPopup)
    PopUp(NoPopup).MiID = NoPopup
    NoPopup2 = Screen.Height / FrmPopUp.ImgBack.Height
End Sub

Public Sub Redondear(Frm As Form)
    Dim Control As Control
    For Each Control In Frm
        If TypeOf Control Is PictureBox Then
            SetWindowRgn Control.hWnd, CreateRoundRectRgn(0, 0, Control.Width / 15, Control.Height / 15, 6, 6), True
        End If
    Next
End Sub

Function FileExists(ByVal sFileName As String) As Boolean
    Dim i As Integer
    FileExists = False
    On Error GoTo NotFound
    i = GetAttr(sFileName)
    FileExists = True
    Exit Function
NotFound:
    FileExists = False
End Function

Public Sub ChkManifest()
    Dim FileNum As Long
    Dim DataArray() As Byte
    Dim ArchivoManifest As String
    ArchivoManifest = App.Path & "\" & App.EXEName & ".exe.manifest"
    If FileExists(ArchivoManifest) = True Then Exit Sub
    DataArray = LoadResData("MANIFEST", "SKIN")
    FileNum = FreeFile
    Open ArchivoManifest For Binary As #FileNum
    Put #FileNum, 1, DataArray()
    Close #FileNum
    DoEvents
    Erase DataArray
End Sub
