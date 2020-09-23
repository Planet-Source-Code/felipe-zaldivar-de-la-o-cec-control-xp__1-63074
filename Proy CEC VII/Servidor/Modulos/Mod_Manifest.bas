Attribute VB_Name = "Mod_Manifest"
Option Explicit

Public Declare Function CreateRoundRectRgn Lib "gdi32" _
    (ByVal RectX1 As Long, ByVal RectY1 As Long, ByVal RectX2 As Long, _
    ByVal RectY2 As Long, ByVal EllipseWidth As Long, _
    ByVal EllipseHeight As Long) As Long

Public Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long

Public Sub InsertarPicture(frm As Form)

    Dim Control1 As Control
    Dim Control2 As Control
    Dim Picture1 As PictureBox
    Dim Cuantos As Integer
    Cuantos = 0
    
    frm.Visible = False

    For Each Control1 In frm.Controls
        If TypeOf Control1 Is Frame Then
            Cuantos = Cuantos + 1
            Load frm.P1(Cuantos)
            Set Picture1 = frm.P1(Cuantos)
            Set Picture1.Container = Control1
            Picture1.Move 25, 190, Control1.Width - 50, Control1.Height - 205
            Picture1.BackColor = Control1.BackColor
            Picture1.Visible = True
            SetWindowRgn Picture1.hWnd, CreateRoundRectRgn(0, 0, Picture1.Width / 15, Picture1.Height / 15, 6, 6), True
            For Each Control2 In frm.Controls
                If Not TypeOf Control2 Is CommonDialog _
                And Not TypeOf Control2 Is ImageList _
                And Not TypeOf Control2 Is Timer _
                And Not TypeOf Control2 Is Data _
                And Not TypeOf Control2 Is Winsock _
                And Not TypeOf Control2 Is Menu Then
                    If Control2.Container Is Control1 Then
                        If Not Control2 Is Picture1 Then
                            Set Control2.Container = Picture1
                            Control2.Left = Control2.Left - 20
                            Control2.Top = Control2.Top - 205
                            If TypeOf Control2 Is CommandButton Then
                                Control2.BackColor = Picture1.BackColor
                            End If
                            Control2.ZOrder
                        End If
                    End If
                End If
                DoEvents
            Next
        End If
        DoEvents
    Next
    frm.Visible = True
End Sub
