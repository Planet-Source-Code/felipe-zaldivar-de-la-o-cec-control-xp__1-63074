Attribute VB_Name = "Mod_CuentaTemporal"
Option Explicit

Private Const Inicio = "T-"
Private cad As String
Private Num As Long

Public Function CrearCT() As String
    Dim Cont As Integer
    Randomize ' inicializamos el random
    Cont = 0
Again:
    Cont = Cont + 1
    If Cont >= 10 Then Exit Function 'evitamos cualquier #$%&
    Num = 0
    cad = ""
    Do While Len(cad) < 7 'hasta que la cadena tenga 7 o mas caracteres
        Num = Int((9999999 * Rnd) + 1) ''son 7 las posiciones que podemos almacenar
        cad = cad & Num 'almacenamos la cadena
    Loop
    cad = Inicio & Left(cad, 7)
    If VerificarCT(cad) = False Then GoTo Again
    CrearCT = cad
End Function

Public Function VerificarCT(CadII As String) As Boolean
    ConexionVCT
    SqlCVCT = "Select * From Tbl_Acceso Where C_Acceso='" & CadII & "'"
    RsCVCT.Open SqlCVCT, Conecta, adOpenDynamic, adLockBatchOptimistic
    
    If RsCVCT.EOF Then
        VerificarCT = True
    Else
        VerificarCT = False
    End If
End Function
