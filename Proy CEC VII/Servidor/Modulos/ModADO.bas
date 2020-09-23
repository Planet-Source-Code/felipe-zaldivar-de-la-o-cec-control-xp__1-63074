Attribute VB_Name = "Mod_ADO"
Option Explicit

Public Conecta As ADODB.Connection

Public Rs As ADODB.Recordset
Public Rs1 As ADODB.Recordset
Public RsCD As ADODB.Recordset
Public RsCE As ADODB.Recordset
Public RsCH As ADODB.Recordset
Public RsCP As ADODB.Recordset
Public RsCC As ADODB.Recordset
Public RsCM As ADODB.Recordset
Public RsCLogin As ADODB.Recordset
Public RsCCIP As ADODB.Recordset
Public RsCCon As ADODB.Recordset
Public RsCA As ADODB.Recordset
Public RsCVNU As ADODB.Recordset
Public RsCVCT As ADODB.Recordset
Public RsCNU As ADODB.Recordset

Public Sql, SqlCCon, SqlCVCT, SqlCVNU, _
SqlCNU, SqlCA, SqlRep, SqlCH, SqlCP, _
SqlCC, SqlCM, SqlCD, SqlCLogin, SqlCCIP, SQLCE As String

Public Sub Conexion()
    Set Conecta = New ADODB.Connection
    If Conecta.State <> adStateClosed Then Conecta.Close
    Conecta.Open "provider=microsoft.jet.oledb.4.0;data source=" & Direccion & ";Jet OLEDB:Database Password=" & C_BD
    'Conecta.Open "provider=microsoft.jet.oledb.4.0;data source=" _
    & Direccion & ";Persist Security Info=False;Jet OLEDB:System database=" _
    & Mid(Direccion, 1, Len(Direccion) - 4) & ".mdw" _
    & ";Jet OLEDB:Database Password=" & C_BD _
    & ";User Id=VBProgrammer00;Password=VBProgrammer11;"
End Sub

Public Sub CerrarConexiones()
    If Conecta.State <> adStateClosed Then Conecta.Close
    If RsCVCT.State <> adStateClosed Then RsCVCT.Close
    If RsCVNU.State <> adStateClosed Then RsCVNU.Close
    If RsCA.State <> adStateClosed Then RsCA.Close
    If RsCNU.State <> adStateClosed Then RsCNU.Close
    If RsCCon.State <> adStateClosed Then RsCCon.Close
    If Rs.State <> adStateClosed Then Rs.Close
    If Rs1.State <> adStateClosed Then Rs1.Close
    If RsCH.State <> adStateClosed Then RsCH.Close
    If RsCP.State <> adStateClosed Then RsCP.Close
    If RsCM.State <> adStateClosed Then RsCM.Close
    If RsCC.State <> adStateClosed Then RsCC.Close
    If RsCD.State <> adStateClosed Then RsCD.Close
    If RsCLogin.State <> adStateClosed Then RsCLogin.Close
    If RsCCIP.State <> adStateClosed Then RsCCIP.Close
    If RsCE.State <> adStateClosed Then RsCE.Close
End Sub

Public Sub ConexionVCT()
    '/// verificar la cuenta del nuevo usuario
    Conexion
    Set RsCVCT = New ADODB.Recordset
End Sub

Public Sub ConexionVCDU()
    '/// verificar la cuenta del nuevo usuario
    Conexion
    Set RsCVNU = New ADODB.Recordset
End Sub

Public Sub ConexionA()
    
    Conexion
    Set RsCA = New ADODB.Recordset
End Sub

Public Sub ConexionNuevoUsuario()
    '/// conexion para agregar un nuevo usuario
    Conexion
    Set RsCNU = New ADODB.Recordset
End Sub

Public Sub ConexionConfiguracion()
    'conexion para la configuracion
    Conexion
    Set RsCCon = New ADODB.Recordset
End Sub

Public Sub ConexionPrincipal()
    'conexion principal utilizada en varios procesos
    Conexion
    Set Rs = New ADODB.Recordset
End Sub

Public Sub ConexionReportes()
    'conexion para los remportes
    Conexion
    Set Rs1 = New ADODB.Recordset
End Sub

Public Sub ConexionHistorial()
    'conexion para el historial
    Conexion
    Set RsCH = New ADODB.Recordset
End Sub

Public Sub ConexionProcesos()
    'conexion para los procesos
    Conexion
    Set RsCP = New ADODB.Recordset
End Sub

Public Sub ConexionMaquinas()
    'conexion para altas bajas etc.. de las maquinas
    Conexion
    Set RsCM = New ADODB.Recordset
End Sub

Public Sub ConexionConsultas()
    'conexion para realizar consultas
    Conexion
    Set RsCC = New ADODB.Recordset
End Sub

Public Sub ConexionDisponibles()
    'conexion para verificar maquinas disponibles
    Conexion
    Set RsCD = New ADODB.Recordset
End Sub

Public Sub ConexionLogin()
    'conexion para accesar al sistema
    Conexion
    Set RsCLogin = New ADODB.Recordset
End Sub

Public Sub ConexionConsultasIP()
    Conexion
    Set RsCCIP = New ADODB.Recordset
End Sub

Public Sub ConexionEscuelas()
    'conexion para la info de la institucion
    Conexion
    Set RsCE = New ADODB.Recordset
End Sub

