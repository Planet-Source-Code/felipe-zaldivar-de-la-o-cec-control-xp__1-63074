Attribute VB_Name = "Mod_StartUp"
Option Explicit

Public Const GW_HWNDPREV = 3

Declare Function OpenIcon Lib "user32" (ByVal hwnd As Long) As Long
Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long

Public Sub ModeStartUP()
    Dim Reg As Object
    Set Reg = CreateObject("Wscript.shell")
    Reg.RegWrite "HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\Run\" & App.EXEName, App.Path & "\" & App.EXEName & ".exe"
End Sub

Sub ActivatePrevInstance()
    Dim OldTitle As String
    Dim PrevHndl As Long
    Dim result As Long
    
    OldTitle = App.Title
    App.Title = "CEC_BOLA"
    PrevHndl = FindWindow("ThunderRTMain", OldTitle)
    
    If PrevHndl = 0 Then
        PrevHndl = FindWindow("ThunderRT5Main", OldTitle)
    End If
    
    If PrevHndl = 0 Then
        PrevHndl = FindWindow("ThunderRT6Main", OldTitle)
    End If
    
    If PrevHndl = 0 Then
        Exit Sub
    End If
    
    PrevHndl = GetWindow(PrevHndl, GW_HWNDPREV)
    result = OpenIcon(PrevHndl)
    result = SetForegroundWindow(PrevHndl)
    End
End Sub
