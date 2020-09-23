Attribute VB_Name = "Mod_MainManifest"
Option Explicit

Public Declare Sub InitCommonControls Lib "Comctl32.dll" ()


'Public Type tagInitCommonControlsEx
   'lngSize As Long
   'lngICC As Long
'End Type
'Private Declare Function InitCommonControlsEx Lib "Comctl32.dll" (iccex As tagInitCommonControlsEx) As Boolean

'Public Const ICC_USEREX_CLASSES = &H200


'Public Sub Main()
'    On Error Resume Next
'   Dim iccex As tagInitCommonControlsEx

'    MServidor

'    With iccex
'        .lngSize = LenB(iccex)
'        .lngICC = ICC_USEREX_CLASSES
'    End With
'    InitCommonControlsEx iccex
'    MDIPrincipal.Show
'End Sub
'If Preguntar("No se encuentra la base de datos, desea crearla)") = False Then End

Public Sub ChkManifest()
    Dim FileNum     As Integer
    Dim DataArray() As Byte
    Dim ArchivoManifest As String
    
    ArchivoManifest = App.Path & "\" & App.EXEName & ".exe.manifest"
    If FileExists(ArchivoManifest) = True Then Exit Sub
    DataArray = LoadResData("MANIFEST", "SKIN")
    FileNum = FreeFile
    Open ArchivoManifest For Binary As #FileNum
    Put #FileNum, 1, DataArray()
    Close #FileNum
    Erase DataArray
    'MsgBox "Librerias creadas satisfactoriamente ...", , "Atención!!!"
End Sub

Public Sub ChkBasedeDatos()
    Dim FileNum     As Integer
    Dim DataArray() As Byte
    Dim ArchivoBD As String
    
    ArchivoBD = App.Path & "\BD\Bdcec1.mdb"
    If DirExist(App.Path & "\BD") = True Then
        If FileExists(ArchivoBD) = True Then Exit Sub
        MsgBox "La base de datos no fue encontrada se creará una nueva...", , "Atención!!!"
        DataArray = LoadResData("BDCEC1", "BD")
        FileNum = FreeFile
        Open ArchivoBD For Binary As #FileNum
        Put #FileNum, 1, DataArray()
        Close #FileNum
        Erase DataArray
        DoEvents
        'MsgBox "La base de datos fué creada satisfactoriamente ...", , "Atención!!!"
    End If
End Sub

Public Function FileExists(ByVal sFileName As String) As Boolean 'funcion para saber si existe un arhivo
    Dim i As Integer
    FileExists = False
    On Error GoTo NotFound
    i = GetAttr(sFileName)
    FileExists = True
    Exit Function
NotFound:
    FileExists = False
End Function

Public Function DirExist(Directorio As String) As Boolean 'funcion para crear direcotorios
    Dim Ret As Boolean
    DirExist = False
    Ret = (Dir$(Directorio, vbDirectory) <> "")
    If Not Ret Then
        MkDir Directorio
        DirExist = True
    Else
        DirExist = True
    End If
End Function

