VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form FrmVisor 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFCEBB&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "::: Imágenes :::"
   ClientHeight    =   6615
   ClientLeft      =   555
   ClientTop       =   585
   ClientWidth     =   10455
   FillStyle       =   0  'Solid
   Icon            =   "FrmVisor.frx":0000
   LinkTopic       =   "FrmVisor"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   441
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   697
   Begin VB.PictureBox PicDir 
      BackColor       =   &H00E89C78&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      ClipControls    =   0   'False
      FillStyle       =   0  'Solid
      HasDC           =   0   'False
      Height          =   6375
      Left            =   120
      ScaleHeight     =   6375
      ScaleMode       =   0  'User
      ScaleWidth      =   10215
      TabIndex        =   5
      Top             =   120
      Width           =   10215
      Begin VB.DirListBox DirLst 
         Height          =   5265
         Left            =   120
         TabIndex        =   1
         Top             =   600
         Width           =   2775
      End
      Begin VB.DriveListBox Dir 
         Height          =   315
         Left            =   120
         TabIndex        =   0
         Top             =   120
         Width           =   2775
      End
      Begin MSComctlLib.ListView LVImg 
         Height          =   5775
         Left            =   3000
         TabIndex        =   2
         Top             =   120
         Width           =   7095
         _ExtentX        =   12515
         _ExtentY        =   10186
         Arrange         =   2
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         OLEDragMode     =   1
         HotTracking     =   -1  'True
         _Version        =   393217
         Icons           =   "ImgList"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   0
         OLEDragMode     =   1
         NumItems        =   0
      End
      Begin MSComctlLib.ProgressBar PB1 
         Height          =   255
         Left            =   4680
         TabIndex        =   8
         Top             =   2880
         Visible         =   0   'False
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Filtrado:  .jpg"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   4680
         TabIndex        =   10
         Top             =   3240
         Width           =   3735
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Buscando Imágenes; aguarda unos instantes!!!"
         ForeColor       =   &H8000000E&
         Height          =   375
         Left            =   4680
         TabIndex        =   9
         Top             =   2520
         Width           =   3735
      End
      Begin VB.Label lblT 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total de imágenes:"
         ForeColor       =   &H8000000E&
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   6000
         Width           =   1350
      End
      Begin VB.Label lblPreview 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "::: Doble click para visualizar en tamaño real :::"
         ForeColor       =   &H8000000E&
         Height          =   195
         Left            =   5040
         TabIndex        =   6
         Top             =   6000
         Width           =   3300
      End
   End
   Begin VB.PictureBox picThumb 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      Height          =   1152
      Left            =   9240
      ScaleHeight     =   73
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   98
      TabIndex        =   4
      Top             =   6720
      Visible         =   0   'False
      Width           =   1536
   End
   Begin VB.PictureBox PicSrc 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   855
      Left            =   7800
      ScaleHeight     =   795
      ScaleWidth      =   1035
      TabIndex        =   3
      Top             =   6600
      Visible         =   0   'False
      Width           =   1095
   End
   Begin MSComctlLib.ImageList ImgList 
      Left            =   7200
      Top             =   6840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
End
Attribute VB_Name = "FrmVisor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long

Private Const FILE_ATTRIBUTE_READONLY = &H1
Private Const FILE_ATTRIBUTE_HIDDEN = &H2
Private Const FILE_ATTRIBUTE_SYSTEM = &H4
Private Const FILE_ATTRIBUTE_DIRECTORY = &H10
Private Const FILE_ATTRIBUTE_ARCHIVE = &H20
Private Const FILE_ATTRIBUTE_NORMAL = &H80
Private Const FILE_ATTRIBUTE_TEMPORARY = &H100
Private Const FILE_ATTRIBUTE_COMPRESSED = &H800
Private Const MAX_PATH = 260

Private Type FILETIME
       dwLowDateTime As Long
       dwHighDateTime As Long
End Type

Private Type WIN32_FIND_DATA
       dwFileAttributes As Long
       ftCreationTime As FILETIME
       ftLastAccessTime As FILETIME
       ftLastWriteTime As FILETIME
       nFileSizeHigh As Long
       nFileSizeLow As Long
       dwReserved0 As Long
       dwReserved1 As Long
       cFileName As String * MAX_PATH
       cAlternate As String * 14
End Type

Dim EnablePreview As Boolean
Dim FileName As String
Dim INIPath As String
Dim lstFilesFocus As Boolean

Dim LVlst As New Collection
Dim i As Long
Dim FN As String
Dim hHeight As Double, hWidth As Double

Private Sub DirLst_Change()
On Error GoTo Desbloquear

    
    For i = LVlst.Count To 1 Step -1
        LVlst.Remove (i)
    ImgList.ListImages.Clear
    Next
    
    LVImg.Icons = Nothing
    ImgList.ListImages.Clear
    
    LVImg.ListItems.Clear
    Call LVImg.Refresh
    
    GetFiles DirLst.Path
    
    For i = LVlst.Count To 1 Step -1
        FN = LCase$(Right$(LVlst.Item(i), 3))
        If FN <> "jpg" Then
            LVlst.Remove (i)
        End If
    Next
    
    PB1.Min = 0
    PB1.Value = 0
    MDIPrincipal.Enabled = False
    PB1.Visible = True
    LVImg.Visible = False
    For i = 1 To LVlst.Count
        PicSrc.Picture = LoadPicture(LVlst(i))
        hWidth = PicSrc.Width
        hHeight = PicSrc.Height
        If hHeight > 76.8 Then
            hWidth = 76.8 * PicSrc.Width / PicSrc.Height
            hHeight = 76.8
        End If
        If hWidth > 102.4 Then
            hHeight = 102.4 * PicSrc.Height / PicSrc.Width
            hWidth = 102.4
        End If
        picThumb.PaintPicture PicSrc, (picThumb.Width - hWidth) / 2, (picThumb.Height - hHeight) / 2, hWidth, hHeight
        ImgList.ListImages.Add , , picThumb.Image
        If LVImg.Icons Is Nothing Then LVImg.Icons = ImgList
        LVImg.ListItems.Add , , GetFileName(LVlst(i)), i
        DoEvents
        picThumb.Cls
        PB1.Max = LVlst.Count
        PB1.Value = PB1.Value + 1
        lblT.Caption = "Total de imágenes: " & i
        DoEvents
    Next
    PB1.Visible = False
    LVImg.Visible = True
    MDIPrincipal.Enabled = True
    LVImg.Arrange = lvwAutoTop
    DoEvents
    Exit Sub
    
Desbloquear:
    PB1.Visible = False
    LVImg.Visible = True
    MDIPrincipal.Enabled = True
End Sub

Private Sub dir_Change()
    On Error GoTo Err
    DirLst.Path = Dir.Drive

Exit Sub
Err:
If Err.Number = 68 Then
    Dir.Drive = "C:"
End If
End Sub

Private Sub GetFiles(Path As String)
   Dim WFD As WIN32_FIND_DATA
   Dim hFile As Long, fPath As String, fName As String
   Dim colFiles As Collection
   Dim varFile As Variant
   
   fPath = AddBackslash(Path)
   fName = fPath & "*.*"
   Set colFiles = New Collection
   
   hFile = FindFirstFile(fName, WFD)
   If (hFile > 0) And ((WFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) <> FILE_ATTRIBUTE_DIRECTORY) Then
       colFiles.Add fPath & StripNulls(WFD.cFileName)
   End If
   
   While FindNextFile(hFile, WFD)
       If ((WFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) <> FILE_ATTRIBUTE_DIRECTORY) Then
           colFiles.Add fPath & StripNulls(WFD.cFileName)
       End If
   Wend
   
   FindClose hFile
   
   For Each varFile In colFiles
       LVlst.Add varFile
   Next
   Set colFiles = Nothing
End Sub

Private Function StripNulls(F As String) As String
   StripNulls = Left$(F, InStr(1, F, Chr$(0)) - 1)
End Function

Private Function AddBackslash(S As String) As String
   If Len(S) Then
      If Right$(S, 1) <> "\" Then
         AddBackslash = S & "\"
      Else
         AddBackslash = S
      End If
   Else
      AddBackslash = "\"
   End If
End Function

Private Function GetFileName(File As String) As String
    Dim i As Integer
    For i = Len(File) To 1 Step -1
        If Mid$(File, i, 1) = "\" Then
            i = i + 1
            Exit For
        End If
    Next
    GetFileName = Mid$(File, i)
End Function

Private Sub LVImg_DblClick()
    Dim FileName As String
    FileName = AddBackslash(DirLst.Path)
    FileName = FileName & LVImg.SelectedItem
    ShellExecute FrmVisor.hwnd, "", FileName, "", DirLst.Path, 0
End Sub

Private Sub Form_Load()
    SetWindowRgn PicDir.hwnd, CreateRoundRectRgn(0, 0, PicDir.Width, PicDir.Height, 6, 6), True
    MDIPrincipal.AgregarVentana Me, "Visor", "Visor de imagenes .jpg"
    PosicionInicial Me
    DirLst.Path = App.Path
    Dir.Drive = App.Path
    DirLst_Change
End Sub


Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then
        Visible = False
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MDIPrincipal.RemoverVentana Me, "Visor"
End Sub

