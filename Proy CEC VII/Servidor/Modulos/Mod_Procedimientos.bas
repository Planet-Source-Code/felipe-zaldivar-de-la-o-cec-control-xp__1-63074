Attribute VB_Name = "Mod_Procedimientos"
'Option Explicit

Public Function TerminarProg()
    Close 'Cierra Conexiones entre solicitantes y solicitados
    End 'Termina el programa
End Function

Public Function ValidarBD(base1 As String) As Boolean
'Vemos si la base de datos que se introdujo sea correcta
On Error GoTo Problema
    With FrmContraBD
        .DataChec.DatabaseName = base1
        .DataChec.RecordSource = "Select * from Tbl_Acceso"
        .DataChec.Connect = ";Pwd=" & C_BD
        Call .DataChec.Refresh
    End With
    ValidarBD = False
    Exit Function
Problema:
    If Err.Number <> 0 Then
        ValidarBD = True
        Exit Function
    End If
End Function

Public Function Preguntar(Pregunta As String) As Boolean
    RespuestaUsr1 = MsgBox(Pregunta, 4 + 32 + 0, "Atención!!!")
    If RespuestaUsr1 = vbYes Then
        Preguntar = True
    Else
        Preguntar = False
    End If
End Function

Public Sub PosicionInicial(frm As Form)
    frm.Left = ((MDIPrincipal.Width - MDIPrincipal.PicContenedor.Width) / 2) - (frm.Width / 2)
    frm.Top = ((MDIPrincipal.Height - MDIPrincipal.PicVentanas.Height) / 2) - (frm.Height / 2) - 350
End Sub

Public Sub Redondear(frm As Form)
    Dim Control As Control
    For Each Control In frm
        If TypeOf Control Is PictureBox Then
            SetWindowRgn Control.hwnd, CreateRoundRectRgn(0, 0, Control.Width / 15, Control.Height / 15, 6, 6), True
        End If
    Next
End Sub

Public Sub ActivarForm(NForm As Form)
    If NForm.WindowState = 0 Then
        NForm.Visible = True
        SendMessage (NForm.hwnd), WM_CHILDACTIVATE, 0, 0
        SendMessage (NForm.hwnd), WM_SETFOCUS, 0, 0
    Else
        NForm.WindowState = 0
        NForm.Visible = True
        NForm.SetFocus
    End If
End Sub
