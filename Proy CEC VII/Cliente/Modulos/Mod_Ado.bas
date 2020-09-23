Attribute VB_Name = "Mod_Ado"
Option Explicit

Public Conecta As ADODB.Connection
Public Rs As ADODB.Recordset
Public Sql As String
Dim DireccionMDB As String
Dim DireccionMDW As String

Public Sub Conexion()
    DireccionMDB = App.Path & "\Bd\RPP1.mdb"
    
    BDPROCESOS
    
    Set Conecta = New ADODB.Connection
    Set Rs = New ADODB.Recordset
    Conecta.Open "provider=microsoft.jet.oledb.4.0;data source=" & DireccionMDB & ";Jet OLEDB:Database Password=control"
End Sub

Private Sub BDPROCESOS()
    Dim FileNum    As Integer
    Dim DataArray() As Byte
    
    If FileExists(DireccionMDB) = False Then
        DataArray = LoadResData("MDB", "BD")
        FileNum = FreeFile
        Open DireccionMDB For Binary As #FileNum
        Put #FileNum, 1, DataArray()
        Close #FileNum
        Erase DataArray
    End If
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
