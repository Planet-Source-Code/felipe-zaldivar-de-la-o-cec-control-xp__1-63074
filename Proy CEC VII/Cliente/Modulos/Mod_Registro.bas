Attribute VB_Name = "Mod_Registro"
Option Explicit

Public Sub EditarRegistroI(Reg_Dir As String, Reg_Value As Integer)
    On Error Resume Next
    Dim Reg_Obj As Object
    Set Reg_Obj = CreateObject("wscript.shell")
    Reg_Obj.RegWrite Reg_Dir, Reg_Value, "REG_DWORD"
End Sub

Public Sub EditarRegistros(Reg_Dir As String, Reg_Value As String)
    On Error Resume Next
    Dim Reg_Obj As Object
    Set Reg_Obj = CreateObject("wscript.shell")
    Reg_Obj.RegWrite Reg_Dir, Reg_Value
End Sub

Public Sub HabilitarRegistro()
    EditarRegistroI "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\System\DisableTaskmgr", 0 'habilitar el Admon. de Tareas
    EditarRegistroI "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\System\DisableChangePassword", 0 'habilita el cambio de password
    EditarRegistroI "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\System\NoDispScrSavPage", 0 'habilita la opcion de contraseña para el screen saver
    EditarRegistroI "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\System\NoDispSettingsPage", 0 'habilita las opciones de pagina
    EditarRegistroI "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\System\NoDispCPL", 0 'deshabilita el panel de control
    EditarRegistroI "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\System\NoConfigPage", 0 'habilita las pag. de propiedades de hardware
    EditarRegistroI "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\System\DisableRegistryTools", 0 'deshabilita las herramientas de registro
    EditarRegistroI "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\System\DisableLockWorkstation", 0 'deshabilita el bloqueo de maquina
    
    EditarRegistroI "HKEY_CURRENT_USER\Software\Policies\Microsoft\Windows\System\DisableCMD", 0 'habilita el command promt
    
    EditarRegistroI "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\NoLogoff", 0 'habilita el cierre de sesion
    EditarRegistroI "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\NoClose", 0 'habilita el apagado de la maquina
    EditarRegistroI "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\NoPropertiesMyComputer", 0 'habilita las propiedades de mi Pc
End Sub

Public Sub DesHabilitarRegistro()
    EditarRegistroI "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\System\DisableTaskmgr", 1 'deshabilitar el Admon. de Tareas
    EditarRegistroI "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\System\DisableChangePassword", 1 'deshabilita el cambio de password
    EditarRegistroI "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\System\NoDispScrSavPage", 1 'deshabilita la opcion de contraseña para el screen saver
    EditarRegistroI "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\System\NoDispSettingsPage", 1 'deshabilita las opciones de pagina
    EditarRegistroI "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\System\NoDispCPL", 1 'deshabilita el panel de control
    EditarRegistroI "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\System\NoConfigPage", 1 'deshabilita las pag. de propiedades de hardware
    EditarRegistroI "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\System\DisableRegistryTools", 1 'deshabilita las herramientas de registro
    EditarRegistroI "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\System\DisableLockWorkstation", 1 'deshabilita el bloqueo de maquina
    
    
    EditarRegistroI "HKEY_CURRENT_USER\Software\Policies\Microsoft\Windows\System\DisableCMD", 1 'deshabilita el command promt
    
    EditarRegistroI "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\NoLogoff", 1 'deshabilita el cierre de sesion
    EditarRegistroI "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\NoClose", 1 'deshabilita el apagado de la maquina
    EditarRegistroI "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\NoPropertiesMyComputer", 1 'deshabilita las propiedades de mi Pc
End Sub
