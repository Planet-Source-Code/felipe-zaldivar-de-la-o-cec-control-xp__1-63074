Attribute VB_Name = "Mod_POrient"
Private Const CCHDEVICENAME = 32
Private Const CCHFORMNAME = 32

'cONSTANTES DEL NT sECURITY
Private Const STANDARD_RIGHTS_REQUIRED = &HF0000
Private Const PRINTER_ACCESS_ADMINISTER = &H4
Private Const PRINTER_ACCESS_USE = &H8
Private Const PRINTER_ALL_ACCESS = (STANDARD_RIGHTS_REQUIRED Or PRINTER_ACCESS_ADMINISTER Or PRINTER_ACCESS_USE)

Private Const DM_MODIFY = 8
Private Const DM_IN_BUFFER = DM_MODIFY
Private Const DM_COPY = 2
Private Const DM_OUT_BUFFER = DM_COPY
Private Const DM_DUPLEX = &H1000&
Private Const DMDUP_SIMPLEX = 1
Private Const DMDUP_VERTICAL = 2
Private Const DMDUP_HORIZONTAL = 3
Private Const DM_ORIENTATION = &H1&
Private Const DM_PaperSize = vbPRPSA4

Private PageDirection As Integer


Private Type DEVMODE
    dmDeviceName As String * CCHDEVICENAME
    dmSpecVersion As Integer
    dmDriverVersion As Integer
    dmSize As Integer
    dmDriverExtra As Integer
    dmFields As Long
    dmOrientation As Integer
    dmPaperSize As Integer
    dmPaperLength As Integer
    dmPaperWidth As Integer
    dmScale As Integer
    dmCopies As Integer
    dmDefaultSource As Integer
    dmPrintQuality As Integer
    dmColor As Integer
    dmDuplex As Integer
    dmYResolution As Integer
    dmTTOption As Integer
    dmCollate As Integer
    dmFormName As String * CCHFORMNAME
    dmLogPixels As Integer
    dmBitsPerPel As Long
    dmPelsWidth As Long
    dmPelsHeight As Long
    dmDisplayFlags As Long
    dmDisplayFrequency As Long
    dmICMMethod As Long
    dmICMIntent As Long
    dmMediaType As Long
    dmDitherType As Long
    dmReserved1 As Long
    dmReserved2 As Long
End Type

Private Type PRINTER_DEFAULTS

    pDatatype As String
    pDevMode As Long
    DesiredAccess As Long
End Type




Private Declare Function OpenPrinter Lib "winspool.drv" Alias "OpenPrinterA" (ByVal pPrinterName As String, phPrinter As Long, pDefault As PRINTER_DEFAULTS) As Long
Private Declare Function SetPrinter Lib "winspool.drv" Alias "SetPrinterA" (ByVal hPrinter As Long, ByVal Level As Long, pPrinter As Any, ByVal Command As Long) As Long
Private Declare Function GetPrinter Lib "winspool.drv" Alias "GetPrinterA" (ByVal hPrinter As Long, ByVal Level As Long, pPrinter As Any, ByVal cbBuf As Long, pcbNeeded As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Private Declare Function ClosePrinter Lib "winspool.drv" (ByVal hPrinter As Long) As Long


Private Declare Function DocumentProperties Lib "winspool.drv" Alias "DocumentPropertiesA" (ByVal hwnd As Long, ByVal hPrinter As Long, ByVal pDeviceName As String, ByVal pDevModeOutput As Any, ByVal pDevModeInput As Any, ByVal fMode As Long) As Long

Private Sub SetOrientation(NewSetting As Long, chng As Integer, ByVal frm As Form)
    Dim PrinterHandle As Long
    Dim PrinterName As String
    Dim pd As PRINTER_DEFAULTS
    Dim MyDevMode As DEVMODE
    Dim result As Long
    Dim Needed As Long
    Dim pFullDevMode As Long
    Dim pi2_buffer() As Long     '
    
    PrinterName = Printer.DeviceName
    If PrinterName = "" Then
        Exit Sub
    End If
    
    pd.pDatatype = vbNullString
    pd.pDevMode = 0&

    pd.DesiredAccess = PRINTER_ALL_ACCESS
    
    result = OpenPrinter(PrinterName, PrinterHandle, pd)

    result = GetPrinter(PrinterHandle, 2, ByVal 0&, 0, Needed)
    ReDim pi2_buffer((Needed \ 4))
    result = GetPrinter(PrinterHandle, 2, pi2_buffer(0), Needed, Needed)
    
    pFullDevMode = pi2_buffer(7)
 
    Call CopyMemory(MyDevMode, ByVal pFullDevMode, Len(MyDevMode))

    MyDevMode.dmDuplex = NewSetting
    MyDevMode.dmFields = DM_DUPLEX Or DM_ORIENTATION Or DM_PaperSize
    MyDevMode.dmOrientation = chng
    MyDevMode.dmPaperSize = vbPRPSA4

    Call CopyMemory(ByVal pFullDevMode, MyDevMode, Len(MyDevMode))

    result = DocumentProperties(frm.hwnd, PrinterHandle, PrinterName, ByVal pFullDevMode, ByVal pFullDevMode, DM_IN_BUFFER Or DM_OUT_BUFFER)

    result = SetPrinter(PrinterHandle, 2, pi2_buffer(0), 0&)
    
    Call ClosePrinter(PrinterHandle)
    
    Dim p As Printer
    For Each p In Printers
        If p.DeviceName = PrinterName Then
            Set Printer = p
            Exit For
        End If
    Next p
    Printer.Duplex = MyDevMode.dmDuplex
    Printer.PaperSize = MyDevMode.dmPaperSize

End Sub

Public Sub POLandscape(ByVal frm As Form)
    PageDirection = 2
    Call SetOrientation(DMDUP_SIMPLEX, PageDirection, frm)
 
End Sub

Public Sub RPOrientation(ByVal frm As Form)

    If PageDirection = 1 Then
        PageDirection = 2
    Else
        PageDirection = 1
    End If

    Call SetOrientation(DMDUP_SIMPLEX, PageDirection, frm)
End Sub

Public Sub POPortrait(ByVal frm As Form)
    
    PageDirection = 1
 
    Call SetOrientation(DMDUP_SIMPLEX, PageDirection, frm)
End Sub

