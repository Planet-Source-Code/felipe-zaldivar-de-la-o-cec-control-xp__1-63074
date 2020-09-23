Attribute VB_Name = "Mod_Snapshot"
Option Explicit
Global PicFolder As String
Private Const Modulo As String = "Mod_Snapshot"
Public Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)

Public Sub GetScreenShot(SetObj As Object, NombreArchivo As String)
    On Error GoTo Error
    Const Evento As String = "GetScreenShot"
    Dim c  As New cDIBSection
    Dim FilePicture As String
    Dim Ret As Long
    Clipboard.Clear
    Call keybd_event(vbKeySnapshot, 0, 0, 0)
    DoEvents
    
    FrmSnapShot.Height = Screen.Height
    FrmSnapShot.Width = Screen.Width
    FrmSnapShot.Picture1.Height = Screen.Height
    FrmSnapShot.Picture1.Width = Screen.Width
    FrmSnapShot.Picture1.Top = 0
    FrmSnapShot.Picture1.Left = 0
    
    PicFolder = App.Path & "\PicFiles"
    FilePicture = PicFolder & "\" & NombreArchivo
    
    Ret = (Dir$(PicFolder, vbDirectory) <> "")
    
    If Not Ret Then
        MkDir PicFolder
    End If
    DoEvents
    
    SetObj.Cls
    SetObj.Picture = Clipboard.GetData
    SavePicture SetObj.Image, CStr((FilePicture & ".Bmp"))
    DoEvents
    
    c.CreateFromPicture FrmSnapShot.Picture1.Picture
    SaveJPG c, (FilePicture & ".jpg"), 70
    Kill (FilePicture & ".Bmp")
    DoEvents
    Exit Sub

Error:
    'Call FrmError.ErrorHandler(Err.Number, Err.Description, Err.Source, Modulo, Evento)
    Resume Next
End Sub
