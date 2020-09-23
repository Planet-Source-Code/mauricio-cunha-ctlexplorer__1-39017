VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl ctlExplorer 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4080
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5655
   ScaleHeight     =   4080
   ScaleWidth      =   5655
   ToolboxBitmap   =   "ctlExplorer.ctx":0000
   Begin VB.ListBox List1 
      Height          =   450
      Left            =   1320
      Sorted          =   -1  'True
      TabIndex        =   4
      Top             =   3000
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.DirListBox Dir1 
      Height          =   315
      Left            =   4080
      TabIndex        =   3
      Top             =   240
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox picBuffer 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   360
      Left            =   3600
      ScaleHeight     =   360
      ScaleWidth      =   360
      TabIndex        =   2
      Top             =   600
      Visible         =   0   'False
      Width           =   360
   End
   Begin MSComctlLib.ImageList ImgFiles32 
      Left            =   4560
      Top             =   1200
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImgFiles16 
      Left            =   3960
      Top             =   1200
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   1095
      Left            =   240
      TabIndex        =   0
      Top             =   1320
      Visible         =   0   'False
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   1931
      View            =   3
      Arrange         =   2
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      Icons           =   "ImgFiles32"
      SmallIcons      =   "ImgFiles16"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   1
         Text            =   "Size"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Type"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Modified"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Created in"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   1095
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   1931
      _Version        =   393217
      Indentation     =   34
      LineStyle       =   1
      Sorted          =   -1  'True
      Style           =   7
      ImageList       =   "imgFolder"
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OLEDragMode     =   1
      OLEDropMode     =   1
   End
   Begin MSComctlLib.ImageList imgFolder 
      Left            =   120
      Top             =   3240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
End
Attribute VB_Name = "ctlExplorer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Const MAX_PATH As Long = 260
Private Const LVM_FIRST = &H1000

Private Const FILE_ATTRIBUTE_ARCHIVE = &H20
Private Const FILE_ATTRIBUTE_COMPRESSED = &H800
Private Const FILE_ATTRIBUTE_DIRECTORY = &H10
Private Const FILE_ATTRIBUTE_HIDDEN = &H2
Private Const FILE_ATTRIBUTE_NORMAL = &H80
Private Const FILE_ATTRIBUTE_READONLY = &H1
Private Const FILE_ATTRIBUTE_SYSTEM = &H4
Private Const FILE_ATTRIBUTE_TEMPORARY = &H100

Private Const TEXT_RESOURCE_WORKSPACE = 4162
Private Const TEXT_RESOURCE_MYCOMPUTER = 9216
Private Const TEXT_RESOURCE_CONTROLPANEL = 4161

Private Const TEXT_RESOURCE_COL_NAME = 8976
Private Const TEXT_RESOURCE_COL_SIZE = 8978
Private Const TEXT_RESOURCE_COL_TYPE = 8979
Private Const TEXT_RESOURCE_COL_MODIFIED = 8980
Private Const TEXT_RESOURCE_COL_CREATED = 8996


Private Const ICON_RESOURCE_WORKSPACE = 34
Private Const ICON_RESOURCE_MYCOMPUTER = 16
Private Const ICON_RESOURCE_MYDOCUMENTS = 20
Private Const ICON_RESOURCE_NETWOORK = 17
Private Const ICON_RESOURCE_CONTROLPANEL = 35

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)
Private Declare Function DrawIconEx Lib "user32" (ByVal hdc As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
Private Declare Function ExtractIconEx Lib "shell32.dll" Alias "ExtractIconExA" (ByVal lpszFile As String, ByVal nIconIndex As Long, phiconLarge As Long, phiconSmall As Long, ByVal nIcons As Long) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Private Declare Function GetLogicalDriveStrings Lib "kernel32" Alias "GetLogicalDriveStringsA" (ByVal nBufferL As Long, ByVal lpBuffer As String) As Long
Private Declare Function GetDiskFreeSpace Lib "kernel32" Alias "GetDiskFreeSpaceA" (ByVal lpRootPathName As String, lpSectorsPerCluster As Long, lpBytesPerSector As Long, lpNumberOfFreeClusters As Long, lpTotalNumberOfClusters As Long) As Long
Private Declare Function ImageList_Draw Lib "comctl32.dll" (ByVal himl&, ByVal I&, ByVal hDCDest&, ByVal x&, ByVal y&, ByVal Flags&) As Long
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function LoadString Lib "user32" Alias "LoadStringA" (ByVal hInstance As Long, ByVal uID As Long, ByVal lpBuffer As String, ByVal nBufferMax As Long) As Long
Private Declare Function SearchPath Lib "kernel32" Alias "SearchPathA" (ByVal lpPath As String, ByVal lpFileName As String, ByVal lpExtension As String, ByVal nBufferL As Long, ByVal lpBuffer As String, ByVal lpFilePart As String) As Long
Private Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" (ByVal pszPath As String, ByVal dwFileAttributes As Long, psfi As typSHFILEINFO, ByVal cbSizeFileInfo As Long, ByVal uFlags As Long) As Long
Private Declare Function SHGetSpecialFolderLocation Lib "shell32" (ByVal hwndOwner As Long, ByVal nFolder As Long, ByRef pIdl As Long) As Long
Private Declare Function SHGetPathFromIDListA Lib "shell32" (ByVal pIdl As Long, ByVal pszPath As String) As Long
Private Declare Function ShellAbout Lib "shell32.dll" Alias "ShellAboutA" (ByVal hwnd As Long, ByVal szApp As String, ByVal szOtherStuff As String, ByVal hIcon As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Type typSHFILEINFO
  hIcon As Long
  iIcon As Long
  dwAttributes As Long
  szDisplayName As String * 260
  szTypeName As String * 80
End Type

Private Const SHGFI_DISPLAYNAME = &H200
Private Const SHGFI_EXETYPE = &H2000
Private Const SHGFI_SYSICONINDEX = &H4000
Private Const SHGFI_SHELLICONSIZE = &H4
Private Const SHGFI_TYPENAME = &H400
Private Const SHGFI_LARGEICON = &H0
Private Const SHGFI_SMALLICON = &H1
Private Const ILD_TRANSPARENT = &H1
Private Const Flags = SHGFI_TYPENAME Or SHGFI_SHELLICONSIZE Or SHGFI_SYSICONINDEX Or SHGFI_DISPLAYNAME Or SHGFI_EXETYPE

Private FSO As FileSystemObject

Private FileInfo As typSHFILEINFO
Private NbFile As Long
Private FileFSToOpen As String
Private StringToFind As String
Private ProgressCancel As Boolean
Private TypeView
Private SizeOn As Boolean
Private OldX
Private InitialFormWith
Private DriveError As Boolean

Private mShell32 As String
Private mShowFolders As Boolean
Private mFilter As String

Private mHideFavorites As Boolean
Private mHideMyDocuments As Boolean
Private mHideNetwork As Boolean
Private mHideControlPanel As Boolean

Private NodeX As Node
Private NodeY As Node
Private Node0 As Node
Private Node1 As Node

Private Type FileTime
  dwLowDateTime As Long
  dwHighDateTime As Long
End Type
   
Private Type WIN32_FIND_DATA
  dwFileAttributes As Long
  ftCreationTime As FileTime
  ftLastAccessTime As FileTime
  ftLastWriteTime As FileTime
  nFileSizeHigh As Long
  nFileSizeLow As Long
  dwReserved0 As Long
  dwReserved1 As Long
  cFileName As String * MAX_PATH
  cAlternate As String * 14
End Type

Private Type RANDYS_OWN_DRIVE_INFO
  DrvSectors As Long
  DrvBytesPerSector As Long
  DrvFreeClusters As Long
  DrvTotalClusters As Long
  DrvSpaceFree As Long
  DrvSpaceUsed As Long
  DrvSpaceTotal As Long
End Type

Public Enum EStyle
 ESTreeFolder = 0
 ESListFile = 1
End Enum

Public Enum EBorderStyle
 EBSNone = 0
 EBSFixedSingle = 1
End Enum

Public Enum EReportColumn
  ERCName = 1
  ERCSize = 2
  ERCType = 3
  ERCModified = 4
  ERCCreated = 5
End Enum

Public Enum ERootFolder
  ERFWorkSpace = 0
  ERFMyComputer = 1
  ERFControlPanel = 2
End Enum

Private mStyle As EStyle
Private mPath As String

Public Event AfterRename(Cancel As Integer, NewName As String)
Public Event BeforeRename(Cancel As Integer)
Public Event Change()
Public Event Click()
Public Event ColumnClick(Column As EReportColumn)
Public Event DblClick()
Public Event DriveNotReady(DriveName As String)
Public Event FolderCheck(FolderName As String, Value As Boolean)
Public Event FolderClick(FolderName As String)
Public Event FolderDblClick(FolderName As String)
Public Event FolderExpand(FolderName As String)
Public Event FolderCollapse(FolderName As String)
Public Event FileCheck(FileName As String, Value As Boolean)
Public Event FileClick(FileName As String)
Public Event FileDblClick(FileName As String)
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Public Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event OLECompleteDrag(Effect As Long)
Public Event OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event OLEDragOver(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
Public Event OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
Public Event OLESetData(Data As MSComctlLib.DataObject, DataFormat As Integer)
Public Event OLEStartDrag(Data As MSComctlLib.DataObject, AllowedEffects As Long)
Public Event Progress(Percent As Integer)
Public Event RootClick(RootType As ERootFolder)
Public Event RootDblClick(RootType As ERootFolder)
Public Event RootCheck(RootType As ERootFolder, Value As Boolean)
Public Event RootExpand(RootType As ERootFolder)
Public Event RootCollapse(RootType As ERootFolder)

Public Property Get ColumnWidth(Column As EReportColumn) As Long
  ColumnWidth = ListView1.ColumnHeaders(Column).Width
End Property
Public Property Let ColumnWidth(Column As EReportColumn, NewValue As Long)
  ListView1.ColumnHeaders(Column).Width = NewValue
End Property

Private Function CalcPercent(iValue As Integer, iMax As Integer) As Integer
 If iMax <= 0 Then
  CalcPercent = 0
 ElseIf iValue > iMax Then
  CalcPercent = 100
 Else
  CalcPercent = (iValue / iMax * 100)
 End If
End Function

Private Function GetSpecialPath(pFolder As Long) As String
Dim sPath As String
Dim IDL As Long
Dim strPath As String
Dim lngPos As Long
    
    If SHGetSpecialFolderLocation(0, pFolder, IDL) = 0 Then
        sPath = String(255, 0)
        SHGetPathFromIDListA IDL, sPath
        lngPos = InStr(sPath, Chr(0))
        If lngPos > 0 Then strPath = Left(sPath, lngPos - 1)
      GetSpecialPath = sPath
    End If
End Function

Private Function GetResourceStringFromFile(sModule As String, idString As Long) As String
Dim hModule As Long
Dim nChars As Long
Dim Buffer As String * 260

   hModule = LoadLibrary(sModule)
   If hModule Then
      nChars = LoadString(hModule, idString, Buffer, 260)
      If nChars Then GetResourceStringFromFile = Left(Buffer, nChars)
      FreeLibrary hModule
   End If
End Function

Private Function ExtractIcon(FileName As String, AddtoImageList As ImageList, PictureBox As PictureBox, PixelsXY As Integer, Optional IsLibrary As Boolean, Optional IconIDInLibrary As Long = -1, Optional IsShell32 As Boolean = True) As Long
Dim SmallIcon As Long
Dim NewImage As ListImage
Dim IconIndex As Integer
Dim ImageCheck As ListImage
Dim IconCount As Long
Dim LibLargeIcon As Long
Dim LibSmallIcon As Long
    
    For Each ImageCheck In AddtoImageList.ListImages
     If LCase(ImageCheck.Key) = LCase(FileName) Then
      ExtractIcon = ImageCheck.Index
      Exit Function
     End If
    Next
    
    If IsLibrary = False Then
     GoSub Extract_Normal_Icon
    Else
     GoSub Extract_Library_Icon
    End If
    Exit Function
    
    
Extract_Normal_Icon:
    If PixelsXY = 16 Then
        SmallIcon = SHGetFileInfo(FileName, 0&, FileInfo, Len(FileInfo), Flags Or SHGFI_SMALLICON)
    Else
        SmallIcon = SHGetFileInfo(FileName, 0&, FileInfo, Len(FileInfo), Flags Or SHGFI_LARGEICON)
    End If
    
    If SmallIcon <> 0 Then
      With PictureBox
        .Height = 15 * PixelsXY
        .Width = 15 * PixelsXY
        .ScaleHeight = 15 * PixelsXY
        .ScaleWidth = 15 * PixelsXY
        .Picture = LoadPicture("")
        .AutoRedraw = True
        .Refresh
        SmallIcon = ImageList_Draw(SmallIcon, FileInfo.iIcon, .hdc, 0, 0, ILD_TRANSPARENT)
      End With
      
      IconIndex = AddtoImageList.ListImages.Count + 1
      Set NewImage = AddtoImageList.ListImages.Add(IconIndex, FileName, PictureBox.Image)
      ExtractIcon = IconIndex
    End If
Return


Extract_Library_Icon:
 Dim FileToWorK As String
 
 If IsShell32 = True Then FileToWorK = mShell32 Else FileToWorK = FileName
 
 IconCount = ExtractIconEx(FileToWorK, -1, 0, 0, 0)
 If IconCount < 0 Then
  ExtractIcon = 0
  Exit Function
 End If
 
 
 Call ExtractIconEx(FileToWorK, IconIDInLibrary, LibLargeIcon, LibSmallIcon, 1)
 With PictureBox
    .Height = 15 * PixelsXY
    .Width = 15 * PixelsXY
    .ScaleHeight = 15 * PixelsXY
    .ScaleWidth = 15 * PixelsXY
    .Picture = LoadPicture("")
    .AutoRedraw = True
    Call DrawIconEx(.hdc, 0, 0, IIf(PixelsXY = 16, LibSmallIcon, LibLargeIcon), PixelsXY, PixelsXY, 0, 0, 3)
    .Refresh
 End With
 
 IconIndex = AddtoImageList.ListImages.Count + 1
 Set NewImage = AddtoImageList.ListImages.Add(IconIndex, FileName, PictureBox.Image)
 ExtractIcon = IconIndex
Return
End Function

Private Function FormatFileSize(lFileSize As Double) As String
Select Case lFileSize
    Case 0 To 1023
     FormatFileSize = Format(lFileSize, "##0") & " Bytes"
    Case 1024 To 1048575
     FormatFileSize = Format(lFileSize / 1024#, "#,##0") & " KB"
    Case 1024# ^ 2 To 1073741823
     FormatFileSize = Format(lFileSize / (1024# ^ 2), "#,##0.00") & " MB"
    Case Is > 1073741823#
    FormatFileSize = Format(lFileSize / (1024# ^ 3), "#,###,##0.00") & " GB"
End Select
End Function

Public Sub ColumnsAutoSize(Column As EReportColumn)
If ListView1.View = lvwReport Then
  Call SendMessage(ListView1.hwnd, LVM_FIRST + 30, Column - 1, -2)
End If
End Sub

Public Property Get SelectedFile() As String
If mStyle = ESListFile And ListView1.SelectedItem.Tag <> 0 Then
   SelectedFile = IIf(Not (ListView1.SelectedItem Is Nothing), ListView1.SelectedItem.Key, Empty)
End If
End Property

Public Property Get SelectedFolder() As String
If mStyle = ESListFile And ListView1.SelectedItem.Tag = 0 Then
   SelectedFolder = IIf(Not (ListView1.SelectedItem Is Nothing), ListView1.SelectedItem.Key, Empty)
End If
 
If mStyle = ESTreeFolder Then
 SelectedFolder = IIf(Not (TreeView1.SelectedItem Is Nothing), TreeView1.SelectedItem.Key, Empty)
End If
End Property

Private Sub LoadTreeView(zPath As String)
On Error GoTo err1

Dim PrevDir As String
Dim TempFile As String
Dim I As Integer
Dim R As Long
Dim D As Drive
Dim DriveName As String
Dim WinPath As String

If TreeView1.Nodes.Count = 0 Then
  WinPath = FSO.GetSpecialFolder(WindowsFolder)
    
  PrevDir = "Root_WorkSpace"
  Set Node0 = TreeView1.Nodes.Add(, , PrevDir, GetResourceStringFromFile(mShell32, TEXT_RESOURCE_WORKSPACE))
  Node0.Image = ExtractIcon(PrevDir, imgFolder, picBuffer, 16, True, ICON_RESOURCE_WORKSPACE)
    
  TempFile = FixPath(WinPath) & "explorer.exe"
  PrevDir = "Root_MyComputer"
  Set Node1 = TreeView1.Nodes.Add(Node0.Key, tvwChild, PrevDir, GetResourceStringFromFile(mShell32, TEXT_RESOURCE_MYCOMPUTER))
  Node1.Image = ExtractIcon(TempFile, imgFolder, picBuffer, ICON_RESOURCE_MYCOMPUTER)
  
  For Each D In FSO.Drives
   DriveName = ""
   Select Case D.DriveType
    Case Removable
     DriveName = GetResourceStringFromFile(mShell32, 9220)
    Case Fixed
     DriveName = D.VolumeName
     If DriveName = "" Then DriveName = GetResourceStringFromFile(mShell32, 9397)
    Case CDRom
     If D.IsReady = True Then DriveName = D.VolumeName
    Case Remote
     DriveName = D.ShareName
   End Select
   
   Set NodeX = TreeView1.Nodes.Add(Node1.Key, tvwChild, FixPath(D.Path), DriveName & " (" & D.Path & ")", ExtractIcon(FixPath(D.Path), imgFolder, picBuffer, 16))
       NodeX.Tag = 0
   Set NodeX = TreeView1.Nodes.Add(FixPath(D.Path), tvwChild, FixPath(D.Path) & "_empty_child")
  Next
  
  Node0.Expanded = True
  Node1.Expanded = True

  If mHideFavorites = False Then
    PrevDir = GetSpecialPath(&H6)
    Set NodeX = TreeView1.Nodes.Add(Node1.Key, tvwChild, PrevDir, FSO.GetBaseName(PrevDir), ExtractIcon(PrevDir, imgFolder, picBuffer, 16))
        NodeX.Tag = 0
    Set NodeX = TreeView1.Nodes.Add(PrevDir, tvwChild, PrevDir & "_empty_child")
  End If
  
  If mHideMyDocuments = False Then
    PrevDir = GetSpecialPath(&H5)
    Set NodeX = TreeView1.Nodes.Add(Node1.Key, tvwChild, PrevDir, FSO.GetBaseName(PrevDir), ExtractIcon(PrevDir, imgFolder, picBuffer, 16, True, ICON_RESOURCE_MYDOCUMENTS))
        NodeX.Tag = 0
    Set NodeX = TreeView1.Nodes.Add(PrevDir, tvwChild, PrevDir & "_empty_child")
  End If
  
  If mHideNetwork = False Then
    PrevDir = GetSpecialPath(&H13)
    Set NodeX = TreeView1.Nodes.Add(Node0.Key, tvwChild, PrevDir, FSO.GetBaseName(PrevDir), ExtractIcon(PrevDir, imgFolder, picBuffer, 16, True, ICON_RESOURCE_NETWOORK))
        NodeX.Tag = 0
    Set NodeX = TreeView1.Nodes.Add(PrevDir, tvwChild, PrevDir & "_empty_child")
  End If
  
  If mHideControlPanel = False Then
    PrevDir = "Root_ControlPanel"
    Set NodeX = TreeView1.Nodes.Add(Node0.Key, tvwChild, PrevDir, GetResourceStringFromFile(FixPath(FSO.GetSpecialFolder(SystemFolder)) & "shell32.dll", TEXT_RESOURCE_CONTROLPANEL), ExtractIcon(PrevDir, imgFolder, picBuffer, 16, True, ICON_RESOURCE_CONTROLPANEL))
        NodeX.Tag = 0
  End If
  
  GetFolders zPath
End If
Exit Sub


err1:
If Err.Number > 0 Then
 Err.Raise Err.Number, Ambient.DisplayName & "_LoadTreeView", Err.Description
 Exit Sub
End If
End Sub

Public Sub ListRootFolder(RootType As ERootFolder)
Dim PrevDir As String
Dim TempFile As String
Dim I As Integer
Dim R As Long
Dim D As Drive
Dim DriveName As String
Dim WinPath As String
Dim zPath As String
Dim szPath As String
Dim hFile As Long
Dim Result As Long
Dim WFD As WIN32_FIND_DATA
Dim TMP As ListItem
Dim TS As String
Dim FinalPath As String
Dim StrTemp As String
       
  ListView1.ListItems.Clear
  If RootType = ERFWorkSpace Then GoSub Show_WorkSpace
  If RootType = ERFMyComputer Then GoSub Show_MyComputer
  If RootType = ERFControlPanel Then GoSub Show_ControlPanel
  Exit Sub
       
Show_WorkSpace:
  WinPath = FSO.GetSpecialFolder(WindowsFolder)
  
  TempFile = FixPath(WinPath) & "explorer.exe"
  PrevDir = "Root_MyComputer"
  Set TMP = ListView1.ListItems.Add(, PrevDir, GetResourceStringFromFile(mShell32, TEXT_RESOURCE_MYCOMPUTER), ExtractIcon(TempFile, ImgFiles32, picBuffer, 32), ExtractIcon(TempFile, ImgFiles16, picBuffer, 16))
      TMP.Tag = 2
  
  If mHideNetwork = False Then
    PrevDir = GetSpecialPath(&H13)
    Set TMP = ListView1.ListItems.Add(, PrevDir, FSO.GetBaseName(PrevDir), ExtractIcon(PrevDir, ImgFiles32, picBuffer, 32), ExtractIcon(PrevDir, ImgFiles16, picBuffer, 16, True, ICON_RESOURCE_NETWOORK))
        TMP.Tag = 2
  End If
  
  If mHideControlPanel = False Then
    PrevDir = "Root_ControlPanel"
    Set TMP = ListView1.ListItems.Add(, PrevDir, GetResourceStringFromFile(mShell32, TEXT_RESOURCE_CONTROLPANEL), ExtractIcon(PrevDir, ImgFiles32, picBuffer, 32, True, ICON_RESOURCE_CONTROLPANEL), ExtractIcon(PrevDir, ImgFiles16, picBuffer, 16, True, ICON_RESOURCE_CONTROLPANEL))
        TMP.Tag = 2
  End If
Return
       
Show_MyComputer:
  I = 0
  For Each D In FSO.Drives
   I = I + 1
   RaiseEvent Progress(CalcPercent(I, FSO.Drives.Count))
   Select Case D.DriveType
    Case Removable
      DriveName = GetResourceStringFromFile(mShell32, 9220)
    Case Fixed
     DriveName = D.VolumeName
     If DriveName = "" Then DriveName = GetResourceStringFromFile(mShell32, 9397)
    Case CDRom
     If D.IsReady = True Then DriveName = D.VolumeName
    Case Remote
     DriveName = D.ShareName
   End Select
   
  Set TMP = ListView1.ListItems.Add(, FixPath(D.Path), DriveName & " (" & D.Path & ")", ExtractIcon(FixPath(D.Path), ImgFiles32, picBuffer, 32), ExtractIcon(FixPath(D.Path), ImgFiles16, picBuffer, 16))
      If D.IsReady Then TMP.SubItems(1) = FormatFileSize(D.TotalSize)
      TMP.Tag = 0
 Next
  
  If mHideFavorites = False Then
    PrevDir = GetSpecialPath(&H6)
    Set TMP = ListView1.ListItems.Add(, PrevDir, FSO.GetBaseName(PrevDir), ExtractIcon(PrevDir, ImgFiles32, picBuffer, 32), ExtractIcon(PrevDir, ImgFiles16, picBuffer, 16))
        TMP.Tag = 0
  End If
  
  If mHideMyDocuments = False Then
    PrevDir = GetSpecialPath(&H5)
    Set TMP = ListView1.ListItems.Add(, PrevDir, FSO.GetBaseName(PrevDir), ExtractIcon(PrevDir, ImgFiles32, picBuffer, 32), ExtractIcon(PrevDir, ImgFiles16, picBuffer, 16, True, ICON_RESOURCE_MYDOCUMENTS))
        TMP.Tag = 0
  End If
  
  RaiseEvent Progress(0)
Return
       
       
Show_ControlPanel:
  zPath = FixPath(FSO.GetSpecialFolder(SystemFolder))
  szPath = FixPath(FSO.GetSpecialFolder(SystemFolder)) & "*.cpl" & Chr(0)
  hFile = FindFirstFile(szPath, WFD)
  RaiseEvent Progress(50)
  Do
     TS = StripNull(WFD.cFileName)
     FinalPath = FixPath(zPath) & TS
     If WFD.dwFileAttributes <> FILE_ATTRIBUTE_DIRECTORY Then
      If FSO.FileExists(FinalPath) = True Then
       StrTemp = ""
       If InStr(1, FinalPath, "access.cpl") > 0 Then StrTemp = GetResourceStringFromFile(FinalPath, 4)
       If InStr(1, FinalPath, "appwiz.cpl") > 0 Then StrTemp = GetResourceStringFromFile(FinalPath, 2001)
       If InStr(1, FinalPath, "desk.cpl") > 0 Then StrTemp = GetResourceStringFromFile(FinalPath, 41)
       If InStr(1, FinalPath, "hdwwiz.cpl") > 0 Then StrTemp = GetResourceStringFromFile(FinalPath, 1000)
       If InStr(1, FinalPath, "inetcpl.cpl") > 0 Then StrTemp = GetResourceStringFromFile(FinalPath, 4312)
       If InStr(1, FinalPath, "intl.cpl") > 0 Then StrTemp = GetResourceStringFromFile(FinalPath, 1)
       If InStr(1, FinalPath, "irprops.cpl") > 0 Then StrTemp = GetResourceStringFromFile(FinalPath, 2)
       If InStr(1, FinalPath, "joy.cpl") > 0 Then StrTemp = GetResourceStringFromFile(FinalPath, 1076)
       If InStr(1, FinalPath, "main.cpl") > 0 Then StrTemp = GetResourceStringFromFile(FinalPath, 100)
       If InStr(1, FinalPath, "mmsys.cpl") > 0 Then StrTemp = GetResourceStringFromFile(FinalPath, 5659)
       If InStr(1, FinalPath, "ncpa.cpl") > 0 Then StrTemp = GetResourceStringFromFile(FinalPath, 2001)
       If InStr(1, FinalPath, "odbccp32.cpl") > 0 Then StrTemp = GetResourceStringFromFile(FinalPath, 1310)
       If InStr(1, FinalPath, "powercfg.cpl") > 0 Then StrTemp = GetResourceStringFromFile(FinalPath, 90)
       If InStr(1, FinalPath, "sticpl.cpl") > 0 Then StrTemp = GetResourceStringFromFile(FinalPath, 1)
       If InStr(1, FinalPath, "sysdm.cpl") > 0 Then StrTemp = GetResourceStringFromFile(FinalPath, 1)
       If InStr(1, FinalPath, "telephon.cpl") > 0 Then StrTemp = GetResourceStringFromFile(FinalPath, 1)
       If InStr(1, FinalPath, "timedate.cpl") > 0 Then StrTemp = GetResourceStringFromFile(FinalPath, 300)
       
       'Try again...
       If StrTemp = "" Then StrTemp = GetResourceStringFromFile(FinalPath, 1)

       If Trim(StrTemp) <> "" Then
        Set TMP = ListView1.ListItems.Add(, FinalPath, StrTemp, ExtractIcon(FinalPath, ImgFiles32, picBuffer, 32, True, 0, False), ExtractIcon(FinalPath, ImgFiles16, picBuffer, 16, True, 0, False))
        TMP.Tag = 1
       End If
      End If
     End If
     WFD.cFileName = ""
     Result = FindNextFile(hFile, WFD)
   Loop Until Result = 0
  RaiseEvent Progress(100)
  FindClose hFile
  With ListView1
   .SortKey = 0
   .SortOrder = lvwAscending
   .Sorted = True
  End With
  RaiseEvent Progress(0)
Return
End Sub

Private Sub LoadListView(zPath As String)
  Dim hFile As Long
  Dim Result As Long
  Dim PathToFind As String
  Dim WFD As WIN32_FIND_DATA
  Dim TMP As ListItem
  Dim MyPos As Long
  Dim TS As String
  Dim FinalPath As String
  Dim StrTemp As Variant
  
       With ListView1
        .ListItems.Clear
        .ColumnHeaders(1).Text = GetResourceStringFromFile(mShell32, TEXT_RESOURCE_COL_NAME)
        .ColumnHeaders(2).Text = GetResourceStringFromFile(mShell32, TEXT_RESOURCE_COL_SIZE)
        .ColumnHeaders(3).Text = GetResourceStringFromFile(mShell32, TEXT_RESOURCE_COL_TYPE)
        .ColumnHeaders(4).Text = GetResourceStringFromFile(mShell32, TEXT_RESOURCE_COL_MODIFIED)
        .ColumnHeaders(5).Text = GetResourceStringFromFile(mShell32, TEXT_RESOURCE_COL_CREATED)
       End With
       
       If mShowFolders = True Then
        List1.Clear
        PathToFind = zPath & "*.*" & Chr(0)
        hFile = FindFirstFile(PathToFind, WFD)
        Do
           TS = StripNull(WFD.cFileName)
           FinalPath = FixPath(zPath) & TS
           If (WFD.dwFileAttributes = 16 Or WFD.dwFileAttributes = 17 Or WFD.dwFileAttributes = 20) And TS <> "." And TS <> ".." Then List1.AddItem Trim(WFD.cFileName)
           WFD.cFileName = ""
           Result = FindNextFile(hFile, WFD)
        Loop Until Result = 0
        FindClose hFile
        
        For MyPos = 0 To List1.ListCount - 1
         RaiseEvent Progress(CalcPercent(CInt(MyPos), List1.ListCount - 1))
         FinalPath = FixPath(zPath) & List1.List(MyPos)
         Set TMP = ListView1.ListItems.Add(, FinalPath, List1.List(MyPos), ExtractIcon(FinalPath, ImgFiles32, picBuffer, 32), ExtractIcon(FinalPath, ImgFiles16, picBuffer, 16))
           TMP.SubItems(1) = FormatFileSize(FSO.GetFolder(FinalPath).Size)
           TMP.SubItems(2) = FSO.GetFolder(FinalPath).Type
           TMP.SubItems(3) = FSO.GetFolder(FinalPath).DateLastModified
           TMP.SubItems(4) = FSO.GetFolder(FinalPath).DateCreated
           TMP.Ghosted = IIf(FSO.GetFolder(FinalPath).Attributes = Hidden, True, False)
           TMP.Tag = 0
        Next
        RaiseEvent Progress(0)
       End If
       
       
       
       List1.Clear
       PathToFind = zPath & mFilter & Chr(0)
       hFile = FindFirstFile(PathToFind, WFD)
       Do
          TS = StripNull(WFD.cFileName)
          FinalPath = FixPath(zPath) & TS
          If WFD.dwFileAttributes <> FILE_ATTRIBUTE_DIRECTORY Then
           If FSO.FileExists(FinalPath) = True Then List1.AddItem FinalPath & vbTab & Trim(TS)
          End If
          WFD.cFileName = ""
          Result = FindNextFile(hFile, WFD)
       Loop Until Result = 0
       FindClose hFile
          
       For MyPos = 0 To List1.ListCount - 1
           RaiseEvent Progress(CalcPercent(CInt(MyPos), List1.ListCount - 1))
           FinalPath = Split(List1.List(MyPos), vbTab)(0)
           If FSO.FileExists(FinalPath) = True Then
            Set TMP = ListView1.ListItems.Add(, FinalPath, Split(List1.List(MyPos), vbTab)(1), ExtractIcon(FinalPath, ImgFiles32, picBuffer, 32), ExtractIcon(FinalPath, ImgFiles16, picBuffer, 16))
             TMP.SubItems(1) = FormatFileSize(FSO.GetFile(FinalPath).Size)
             TMP.SubItems(2) = FSO.GetFile(FinalPath).Type
             TMP.SubItems(3) = FSO.GetFile(FinalPath).DateCreated
             TMP.SubItems(4) = FSO.GetFile(FinalPath).DateCreated
             If FSO.GetFile(FinalPath).Attributes = Hidden Then TMP.Ghosted = True
             TMP.Tag = 1
           End If
      Next
      RaiseEvent Progress(0)
      List1.Clear
End Sub

Private Function StripNull(ByVal WhatStr As String) As String
Dim pos As Integer
pos = InStr(WhatStr, Chr(0))
If pos > 0 Then StripNull = Left(WhatStr, pos - 1) Else StripNull = WhatStr
End Function

Public Property Get Style() As EStyle
 Style = mStyle
End Property
Public Property Let Style(ByVal NewValue As EStyle)
 mStyle = NewValue
 PropertyChanged "Style"
 Call UserControl_Resize
 Call Refresh
End Property

Private Sub Refresh()
 If mStyle = ESTreeFolder Then LoadTreeView FixPath(mPath)
 If mStyle = ESListFile Then LoadListView FixPath(mPath)
End Sub

Private Sub ListView1_AfterLabelEdit(Cancel As Integer, NewString As String)
  RaiseEvent AfterRename(Cancel, NewString)
End Sub

Private Sub ListView1_BeforeLabelEdit(Cancel As Integer)
  RaiseEvent BeforeRename(Cancel)
End Sub

Private Sub ListView1_Click()
 RaiseEvent Click
End Sub

Private Sub ListView1_DblClick()
 RaiseEvent DblClick
 
 If Not (ListView1.SelectedItem Is Nothing) Then
  If ListView1.SelectedItem.Tag = 0 Then
   RaiseEvent FolderDblClick(ListView1.SelectedItem.Key)
  ElseIf ListView1.SelectedItem.Tag = 2 Then
   If InStr(1, ListView1.SelectedItem.Key, "MyComputer") > 0 Then
    RaiseEvent RootDblClick(ERFMyComputer)
   ElseIf InStr(1, ListView1.SelectedItem.Key, "ControlPanel") > 0 Then
    RaiseEvent RootDblClick(ERFControlPanel)
   Else
    RaiseEvent FolderDblClick(ListView1.SelectedItem.Key)
   End If
  Else
   RaiseEvent FileDblClick(ListView1.SelectedItem.Key)
  End If
 End If
End Sub

Private Sub ListView1_ItemCheck(ByVal Item As MSComctlLib.ListItem)
  RaiseEvent FileCheck(Item.Key, Item.Checked)
End Sub

Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
 If Item.Tag = 0 Then
  RaiseEvent FolderClick(Item.Key)
 Else
  RaiseEvent FileClick(Item.Key)
 End If
End Sub

Private Sub ListView1_KeyDown(KeyCode As Integer, Shift As Integer)
  RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub ListView1_KeyPress(KeyAscii As Integer)
  RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub ListView1_KeyUp(KeyCode As Integer, Shift As Integer)
  RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub ListView1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  RaiseEvent MouseDown(Button, Shift, x, y)
End Sub

Private Sub ListView1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  RaiseEvent MouseMove(Button, Shift, x, y)
End Sub

Private Sub ListView1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  RaiseEvent MouseUp(Button, Shift, x, y)
End Sub

Private Sub ListView1_OLECompleteDrag(Effect As Long)
  RaiseEvent OLECompleteDrag(Effect)
End Sub

Private Sub ListView1_OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
  RaiseEvent OLEDragDrop(Data, Effect, Button, Shift, x, y)
End Sub

Private Sub ListView1_OLEDragOver(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
  RaiseEvent OLEDragOver(Data, Effect, Button, Shift, x, y, State)
End Sub

Private Sub ListView1_OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
  RaiseEvent OLEGiveFeedback(Effect, DefaultCursors)
End Sub

Private Sub ListView1_OLESetData(Data As MSComctlLib.DataObject, DataFormat As Integer)
  RaiseEvent OLESetData(Data, DataFormat)
End Sub

Private Sub ListView1_OLEStartDrag(Data As MSComctlLib.DataObject, AllowedEffects As Long)
  RaiseEvent OLEStartDrag(Data, AllowedEffects)
End Sub

Private Sub TreeView1_AfterLabelEdit(Cancel As Integer, NewString As String)
 RaiseEvent AfterRename(Cancel, NewString)
End Sub

Private Sub TreeView1_BeforeLabelEdit(Cancel As Integer)
 RaiseEvent BeforeRename(Cancel)
End Sub

Private Sub TreeView1_Click()
 RaiseEvent Click
End Sub

Private Sub TreeView1_Collapse(ByVal Node As MSComctlLib.Node)
  If Node.Image = 1 Or Node.Image = 2 Then Exit Sub
  If InStr(1, Node.Key, "Root_") > 0 Then
   If InStr(1, LCase(Node.Key), "computer") Then RaiseEvent RootCollapse(ERFMyComputer)
   If InStr(1, LCase(Node.Key), "workspace") Then RaiseEvent RootCollapse(ERFWorkSpace)
   If InStr(1, LCase(Node.Key), "controlpanel") Then RaiseEvent RootCollapse(ERFControlPanel)
  Else
    RaiseEvent FolderCollapse(Node.Key)
  End If
  
  Node.Tag = 0
  While Node.Children > 1
    TreeView1.Nodes.Remove (Node.Child.Index)
  Wend
End Sub

Private Sub TreeView1_DblClick()
 RaiseEvent DblClick
End Sub

Private Sub TreeView1_Expand(ByVal Node As MSComctlLib.Node)
If Node.Tag = 0 Then
    TreeView1.Nodes.Remove (Node.Child.Index)
    Call GetFolders(Node.Key)
    If InStr(1, Node.Key, "Root_") > 0 Then
     If InStr(1, LCase(Node.Key), "computer") Then RaiseEvent RootExpand(ERFMyComputer)
     If InStr(1, LCase(Node.Key), "workspace") Then RaiseEvent RootExpand(ERFWorkSpace)
     If InStr(1, LCase(Node.Key), "controlpanel") Then RaiseEvent RootExpand(ERFControlPanel)
    Else
      RaiseEvent FolderExpand(Node.Key)
    RaiseEvent Change
    End If
    
    
    If DriveError = False Then
     Node.Tag = 1
    Else
     Node.Tag = 0
     DriveError = False
     Node.Parent.Expanded = False
    End If
End If
End Sub

Private Sub GetFolders(DirName As String)
Dim DirText As String
Dim DirPath As String
Dim I As Integer
Dim J As Integer
Dim L As Integer
Dim COL As New Collection

   On Error GoTo Errorhandle
   
   Dir1.Path = DirName
   Dir1.Refresh
   
   For I = 0 To Dir1.ListCount - 1
        COL.Add Dir1.List(I)
   Next I
   
   For I = 1 To COL.Count
         RaiseEvent Progress(CalcPercent(I, COL.Count))
         DirText = ""
         L = Len(COL.Item(I))
         
         For J = L To 0 Step -1
          If Mid(COL.Item(I), J, 1) = "\" Then Exit For
         Next J
         
         DirText = Right(COL.Item(I), L - J)
         DirPath = FixPath(DirName) & DirText
         
         If DirName = GetSpecialPath(&H13) Then
          Set NodeY = TreeView1.Nodes.Add(DirName, tvwChild, COL.Item(I), DirText, ExtractIcon(DirPath, imgFolder, picBuffer, 16, True, 85))
         Else
          Set NodeY = TreeView1.Nodes.Add(DirName, tvwChild, COL.Item(I), DirText, ExtractIcon(DirPath, imgFolder, picBuffer, 16))
         End If
         Dir1.Path = COL.Item(I)
         Dir1.Refresh
         If Dir1.ListCount > 0 Then
             NodeY.Tag = 0
             Set NodeY = TreeView1.Nodes.Add(COL.Item(I), tvwChild, COL.Item(I) & "\a")
         End If
   Next I
   RaiseEvent Progress(0)
Exit Sub

Errorhandle:
    RaiseEvent DriveNotReady(DirName)
    RaiseEvent Progress(0)
    DriveError = True
    Resume Next
End Sub

Private Sub SetFolder(PathFolder As String)
On Error Resume Next
Dim H As Integer
Dim T As Integer
Dim StrPT As String
Dim StrOM As String

H = 3
MainProcess:
Do
 StrPT = TreeView1.Nodes(H).Key
  If InStr(1, StrPT, ":") > 0 Then
    StrOM = LCase(Mid(StrPT, InStr(1, StrPT, ":") - 1, 2) & Mid(StrPT, InStr(1, StrPT, ":") + 1))
    If Left(LCase(PathFolder), Len(StrOM)) = LCase(StrOM) Then
      TreeView1.Nodes(H).Expanded = True
      TreeView1.Nodes(H).Selected = True
      If TreeView1.Nodes(H).Children > 0 Then H = TreeView1.Nodes(H).Child.Index Else Exit Do
      GoTo MainProcess
      Exit Do
    End If
  If H = TreeView1.Nodes(H).LastSibling.Index Then Exit Do
  H = TreeView1.Nodes(H).Next.Index
 End If
Loop
TreeView1.Nodes(H).EnsureVisible
End Sub

Private Sub TreeView1_KeyDown(KeyCode As Integer, Shift As Integer)
  RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub TreeView1_KeyPress(KeyAscii As Integer)
  RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub TreeView1_KeyUp(KeyCode As Integer, Shift As Integer)
  RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub TreeView1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  RaiseEvent MouseDown(Button, Shift, x, y)
End Sub

Private Sub TreeView1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  RaiseEvent MouseMove(Button, Shift, x, y)
End Sub

Private Sub TreeView1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  RaiseEvent MouseUp(Button, Shift, x, y)
End Sub

Private Sub TreeView1_NodeCheck(ByVal Node As MSComctlLib.Node)
If InStr(1, Node.Key, "Root_") > 0 Then
 If InStr(1, LCase(Node.Key), "computer") Then RaiseEvent RootCheck(ERFMyComputer, Node.Checked)
 If InStr(1, LCase(Node.Key), "workspace") Then RaiseEvent RootCheck(ERFWorkSpace, Node.Checked)
 If InStr(1, LCase(Node.Key), "controlpanel") Then RaiseEvent RootCheck(ERFControlPanel, Node.Checked)
Else
  RaiseEvent FolderCheck(Node.Key, Node.Checked)
  RaiseEvent Change
End If
End Sub

Private Sub TreeView1_NodeClick(ByVal Node As MSComctlLib.Node)
If InStr(1, Node.Key, "Root_") > 0 Then
 If InStr(1, LCase(Node.Key), "computer") Then RaiseEvent RootClick(ERFMyComputer)
 If InStr(1, LCase(Node.Key), "workspace") Then RaiseEvent RootClick(ERFWorkSpace)
 If InStr(1, LCase(Node.Key), "controlpanel") Then RaiseEvent RootClick(ERFControlPanel)
Else
  RaiseEvent FolderClick(Node.Key)
  RaiseEvent Change
  mPath = Node.Key
End If
End Sub

Public Sub AboutBox()
Attribute AboutBox.VB_UserMemId = -552
 ShellAbout UserControl.hwnd, "Explorer Controls by Mauricio Cunha", "Control to show files and folder like Windows Explorer." & vbCrLf & "Developed by Mauricio Cunha mcunha98@terra.com.br", Empty
End Sub

Private Sub TreeView1_OLECompleteDrag(Effect As Long)
 RaiseEvent OLECompleteDrag(Effect)
End Sub

Private Sub TreeView1_OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
  RaiseEvent OLEDragDrop(Data, Effect, Button, Shift, x, y)
End Sub

Private Sub TreeView1_OLEDragOver(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
  RaiseEvent OLEDragOver(Data, Effect, Button, Shift, x, y, State)
End Sub

Private Sub TreeView1_OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
  RaiseEvent OLEGiveFeedback(Effect, DefaultCursors)
End Sub

Private Sub TreeView1_OLESetData(Data As MSComctlLib.DataObject, DataFormat As Integer)
  RaiseEvent OLESetData(Data, DataFormat)
End Sub

Private Sub TreeView1_OLEStartDrag(Data As MSComctlLib.DataObject, AllowedEffects As Long)
  RaiseEvent OLEStartDrag(Data, AllowedEffects)
End Sub

Private Sub UserControl_Initialize()
 Set FSO = New FileSystemObject
 mShell32 = FixPath(FSO.GetSpecialFolder(SystemFolder)) & "shell32.dll"
 Call Refresh
End Sub

Private Sub UserControl_InitProperties()
 Path = CurDir
 Filter = "*.*"
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
 UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
 UserControl.BorderStyle = PropBag.ReadProperty("BorderStyle", 1)
 
 ListView1.Arrange = PropBag.ReadProperty("Arrange", 0)
 ListView1.CheckBoxes = PropBag.ReadProperty("CheckBoxes", False)
 TreeView1.CheckBoxes = PropBag.ReadProperty("CheckBoxes", False)
 Set ListView1.Font = PropBag.ReadProperty("Font", Ambient.Font)
 Set TreeView1.Font = PropBag.ReadProperty("Font", Ambient.Font)
 ListView1.FullRowSelect = PropBag.ReadProperty("FullRowSelect", False)
 TreeView1.FullRowSelect = PropBag.ReadProperty("FullRowSelect", False)
 ListView1.GridLines = PropBag.ReadProperty("GridLines", False)
 ListView1.HideColumnHeaders = PropBag.ReadProperty("HideColumnHeaders", False)
 ListView1.HideSelection = PropBag.ReadProperty("HideSelection", True)
 TreeView1.HideSelection = PropBag.ReadProperty("HideSelection", True)
 ListView1.HotTracking = PropBag.ReadProperty("HotTracking", False)
 TreeView1.HotTracking = PropBag.ReadProperty("HotTracking", False)
 ListView1.LabelEdit = PropBag.ReadProperty("LabelEdit", 0)
 TreeView1.LabelEdit = PropBag.ReadProperty("LabelEdit", 0)
 ListView1.LabelWrap = PropBag.ReadProperty("LabelWrap", False)
 ListView1.MultiSelect = PropBag.ReadProperty("MultiSelect", False)
 Set ListView1.MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
 Set TreeView1.MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
 ListView1.MousePointer = PropBag.ReadProperty("MousePointer", 0)
 TreeView1.MousePointer = PropBag.ReadProperty("MousePointer", 0)
 ListView1.OleDragMode = PropBag.ReadProperty("OleDragMode", 0)
 TreeView1.OleDragMode = PropBag.ReadProperty("OleDragMode", 0)
 ListView1.OleDropMode = PropBag.ReadProperty("OleDropMode", 0)
 TreeView1.OleDropMode = PropBag.ReadProperty("OleDropMode", 0)
 TreeView1.SingleSel = PropBag.ReadProperty("SingleSel", False)
 ListView1.Sorted = PropBag.ReadProperty("Sorted", False)
 ListView1.SortKey = PropBag.ReadProperty("SortKey", 0)
 ListView1.SortOrder = PropBag.ReadProperty("SortOrder", 0)
 HideControlPanel = PropBag.ReadProperty("HideControlPanel", False)
 HideFavorites = PropBag.ReadProperty("HideFavorites", False)
 HideMyDocuments = PropBag.ReadProperty("HideMyDocuments", False)
 HideNetwork = PropBag.ReadProperty("HideNetwork", False)
 Filter = PropBag.ReadProperty("Filter", "*.*")
 Path = PropBag.ReadProperty("Path", CurDir)
 Style = PropBag.ReadProperty("Style", 0)
 ShowFolders = PropBag.ReadProperty("ShowFolders", True)
 ListView1.View = PropBag.ReadProperty("View", 0)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
 Call PropBag.WriteProperty("Arrange", ListView1.Arrange, 0)
 Call PropBag.WriteProperty("CheckBoxes", ListView1.CheckBoxes, False)
 Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
 Call PropBag.WriteProperty("GridLines", ListView1.GridLines, False)
 Call PropBag.WriteProperty("HideColumnHeaders", ListView1.HideColumnHeaders, False)
 Call PropBag.WriteProperty("HideSelection", ListView1.HideSelection, True)
 Call PropBag.WriteProperty("HotTracking", ListView1.HotTracking, False)
 Call PropBag.WriteProperty("LabelEdit", ListView1.LabelEdit, 0)
 Call PropBag.WriteProperty("LabelWrap", ListView1.LabelWrap, False)
 Call PropBag.WriteProperty("MultiSelect", ListView1.MultiSelect, False)
 Call PropBag.WriteProperty("MouseIcon", ListView1.MouseIcon, Nothing)
 Call PropBag.WriteProperty("MousePointer", ListView1.MousePointer, 0)
 Call PropBag.WriteProperty("OleDragMode", ListView1.OleDragMode, 0)
 Call PropBag.WriteProperty("OleDropMode", ListView1.OleDropMode, 0)
 Call PropBag.WriteProperty("SingleSel", TreeView1.SingleSel, False)
 Call PropBag.WriteProperty("ShowFolders", mShowFolders, True)
 Call PropBag.WriteProperty("Filter", mFilter, "*.*")
 Call PropBag.WriteProperty("Path", mPath, CurDir)
 Call PropBag.WriteProperty("Style", mStyle, 0)
 Call PropBag.WriteProperty("Sorted", ListView1.Sorted, False)
 Call PropBag.WriteProperty("SortKey", ListView1.SortKey, 0)
 Call PropBag.WriteProperty("SortOrder", ListView1.SortOrder, 0)
 Call PropBag.WriteProperty("View", ListView1.View, 0)
 Call PropBag.WriteProperty("Font", ListView1.Font, Ambient.Font)
 Call PropBag.WriteProperty("FullRowSelect", ListView1.FullRowSelect, False)
 Call PropBag.WriteProperty("BorderStyle", UserControl.BorderStyle, 1)
 Call PropBag.WriteProperty("HideControlPanel", mHideControlPanel, False)
 Call PropBag.WriteProperty("HideFavorites", mHideFavorites, False)
 Call PropBag.WriteProperty("HideMyDocuments", mHideMyDocuments, False)
 Call PropBag.WriteProperty("HideNetwork", mHideNetwork, False)
End Sub

Private Sub UserControl_Resize()
TreeView1.Visible = False
ListView1.Visible = False

Select Case mStyle
  Case ESTreeFolder
    TreeView1.Visible = True
    TreeView1.Move 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
  Case ESListFile
    ListView1.Visible = True
    ListView1.Move 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
End Select
End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As ColumnHeader)
RaiseEvent ColumnClick(ColumnHeader.Index)
End Sub

Private Function FixPath(sPath As String) As String
 FixPath = sPath & IIf(Right(sPath, 1) <> "\", "\", "")
End Function

Public Property Get BorderStyle() As EBorderStyle
 BorderStyle = UserControl.BorderStyle
End Property
Public Property Let BorderStyle(ByVal NewValue As EBorderStyle)
 UserControl.BorderStyle = NewValue
 PropertyChanged "BorderStyle"
End Property

Public Property Get Path() As String
 If Not (TreeView1.SelectedItem Is Nothing) Then
  Path = TreeView1.SelectedItem.Key
 Else
  Path = mPath
 End If
End Property
Public Property Let Path(ByVal NewValue As String)
 mPath = NewValue
 Call Refresh
 SetFolder NewValue
 PropertyChanged "Path"
 RaiseEvent Change
 If Style = ESTreeFolder Then
  RaiseEvent FolderClick(NewValue)
 Else
  RaiseEvent FileClick(NewValue)
 End If
End Property

Public Property Get View() As ListViewConstants
 View = ListView1.View
End Property
Public Property Let View(ByVal NewValue As ListViewConstants)
 ListView1.View = NewValue
 PropertyChanged "View"
End Property

Public Property Get Font() As StdFont
 Set Font = ListView1.Font
End Property
Public Property Set Font(ByVal NewValue As StdFont)
 Set ListView1.Font = NewValue
 Set TreeView1.Font = NewValue
 PropertyChanged "Font"
End Property

Public Property Get FullRowSelect() As Boolean
 FullRowSelect = ListView1.FullRowSelect
End Property
Public Property Let FullRowSelect(ByVal NewValue As Boolean)
 ListView1.FullRowSelect = NewValue
 TreeView1.FullRowSelect = NewValue
 PropertyChanged "FullRowSelect"
End Property

Public Property Get Arrange() As ListArrangeConstants
 Arrange = ListView1.Arrange
End Property
Public Property Let Arrange(ByVal NewValue As ListArrangeConstants)
 ListView1.Arrange = NewValue
 ListView1.Refresh
 PropertyChanged "Arrange"
End Property

Public Property Get CheckBoxes() As Boolean
 CheckBoxes = ListView1.CheckBoxes
End Property
Public Property Let CheckBoxes(ByVal NewValue As Boolean)
 ListView1.CheckBoxes = NewValue
 TreeView1.CheckBoxes = NewValue
 PropertyChanged "CheckBoxes"
 If mStyle = ESListFile Then ListView1.Refresh Else TreeView1.Refresh
End Property

Public Property Get Enabled() As Boolean
 Enabled = UserControl.Enabled
End Property
Public Property Let Enabled(ByVal NewValue As Boolean)
 UserControl.Enabled = NewValue
 ListView1.Enabled = NewValue
 TreeView1.Enabled = NewValue
 PropertyChanged "Enabled"
End Property

Public Property Get GridLines() As Boolean
 GridLines = ListView1.GridLines
End Property
Public Property Let GridLines(ByVal NewValue As Boolean)
 ListView1.GridLines = NewValue
 PropertyChanged "GridLines"
End Property

Public Property Get HideColumnHeaders() As Boolean
 HideColumnHeaders = ListView1.HideColumnHeaders
End Property
Public Property Let HideColumnHeaders(ByVal NewValue As Boolean)
 ListView1.HideColumnHeaders = NewValue
 PropertyChanged "HideColumnHeaders"
End Property

Public Property Get HideSelection() As Boolean
 HideSelection = ListView1.HideSelection
End Property
Public Property Let HideSelection(ByVal NewValue As Boolean)
 ListView1.HideSelection = NewValue
 PropertyChanged "HideSelection"
End Property

Public Property Get HotTracking() As Boolean
 HotTracking = ListView1.HotTracking
End Property
Public Property Let HotTracking(ByVal NewValue As Boolean)
 ListView1.HotTracking = NewValue
 PropertyChanged "HotTracking"
End Property

Public Property Get HoverSelection() As Boolean
 HoverSelection = ListView1.HoverSelection
End Property
Public Property Let HoverSelection(ByVal NewValue As Boolean)
 ListView1.HoverSelection = NewValue
 PropertyChanged "HoverSelection"
End Property

Public Property Get LabelWrap() As Boolean
 LabelWrap = ListView1.LabelWrap
End Property
Public Property Let LabelWrap(ByVal NewValue As Boolean)
 ListView1.LabelWrap = NewValue
 PropertyChanged "LabelWrap"
End Property

Public Property Get LabelEdit() As LabelEditConstants
 LabelEdit = ListView1.LabelEdit
End Property
Public Property Let LabelEdit(ByVal NewValue As LabelEditConstants)
 ListView1.LabelEdit = NewValue
 TreeView1.LabelEdit = NewValue
 PropertyChanged "LabelEdit"
End Property

Public Property Get MultiSelect() As Boolean
 MultiSelect = ListView1.MultiSelect
End Property
Public Property Let MultiSelect(ByVal NewValue As Boolean)
 ListView1.MultiSelect = NewValue
 PropertyChanged "MultiSelect"
End Property

Public Property Get ShowFolders() As Boolean
 ShowFolders = mShowFolders
End Property
Public Property Let ShowFolders(ByVal NewValue As Boolean)
 mShowFolders = NewValue
 PropertyChanged "ShowFolders"
 If mStyle = ESListFile Then Call Refresh
End Property

Public Property Get Filter() As String
 Filter = mFilter
End Property
Public Property Let Filter(ByVal NewValue As String)
 mFilter = NewValue
 PropertyChanged "Filter"
 Call Refresh
End Property

Public Property Get FileCount() As Long
 FileCount = ListView1.ListItems.Count
End Property

Public Property Get ColumnCount() As Long
 ColumnCount = ListView1.ColumnHeaders.Count
End Property

Public Sub Clear()
 If mStyle = ESListFile Then ListView1.ListItems.Clear
 If mStyle = ESTreeFolder Then TreeView1.Nodes.Clear
End Sub

Public Property Get FileChecked(Index As Long) As Boolean
 If ValidIndex(Index) = False Then Exit Property
 FileChecked = ListView1.ListItems(Index).Checked
End Property
Public Property Let FileChecked(Index As Long, Value As Boolean)
 If ValidIndex(Index) = False Then Exit Property
 ListView1.ListItems(Index).Checked = Value
End Property

Public Property Get FileSelected(Index As Long) As Boolean
 If ValidIndex(Index) = False Then Exit Property
 FileSelected = ListView1.ListItems(Index).Selected
End Property
Public Property Let FileSelected(Index As Long, Value As Boolean)
 If ValidIndex(Index) = False Then Exit Property
 ListView1.ListItems(Index).Selected = Value
End Property

Public Property Get FolderChecked(Index As Long) As Boolean
 If ValidIndex(Index) = False Then Exit Property
 FolderChecked = TreeView1.Nodes(Index).Checked
End Property
Public Property Let FolderChecked(Index As Long, Value As Boolean)
 If ValidIndex(Index) = False Then Exit Property
 TreeView1.Nodes(Index).Checked = Value
End Property

Public Property Get FileIndex(FileName As String) As Long
On Error GoTo err1
 FileIndex = ListView1.ListItems(FileName).Index
 Exit Property
 
err1:
 If Err > 0 Then
  FileIndex = 0
  Exit Property
 End If
End Property

Private Function ValidIndex(V As Long) As Boolean
If mStyle = ESListFile Then
 If ListView1.ListItems.Count = 0 Then
  ValidIndex = False
  Exit Function
 ElseIf V > ListView1.ListItems.Count Then
  ValidIndex = False
  Exit Function
 ElseIf V <= 0 Then
  ValidIndex = False
  Exit Function
 Else
  ValidIndex = True
 End If
End If


If mStyle = ESTreeFolder Then
 If TreeView1.Nodes.Count = 0 Then
  ValidIndex = False
  Exit Function
 ElseIf V > TreeView1.Nodes.Count Then
  ValidIndex = False
  Exit Function
 ElseIf V < 0 Then
  ValidIndex = False
  Exit Function
 Else
  ValidIndex = True
 End If
End If
End Function

Public Property Get SelectedCount() As Long
Dim SC As Long
Dim I As Long

SC = 0
If mStyle = ESListFile Then
  If ListView1.ListItems.Count >= 1 Then
   For I = 1 To ListView1.ListItems.Count
    If ListView1.ListItems(I).Selected = True Then SC = SC + 1
   Next I
  End If
End If

If mStyle = ESTreeFolder Then
  If TreeView1.Nodes.Count >= 1 Then
   For I = 1 To TreeView1.Nodes.Count
    If TreeView1.Nodes(I).Selected = True Then SC = SC + 1
   Next I
  End If
End If

SelectedCount = SC
End Property

Public Property Get CheckedCount() As Long
Dim SC As Long
Dim I As Long

SC = 0
If mStyle = ESListFile Then
  If ListView1.ListItems.Count >= 1 Then
   For I = 1 To ListView1.ListItems.Count
    If ListView1.ListItems(I).Checked = True Then SC = SC + 1
   Next I
  End If
End If

If mStyle = ESTreeFolder Then
  If TreeView1.Nodes.Count >= 1 Then
   For I = 1 To TreeView1.Nodes.Count
    If TreeView1.Nodes(I).Checked = True Then SC = SC + 1
   Next I
  End If
End If

CheckedCount = SC
End Property

Public Property Get FolderCount() As Long
 If mStyle = ESListFile Then
  Dim R As Long
  Dim F As Long
  F = 0
  For R = 1 To ListView1.ListItems.Count
   If ListView1.ListItems(R).Tag = 1 Then F = F + 1
  Next R
  FolderCount = F
 End If
 
 If mStyle = ESTreeFolder Then
  If Not (TreeView1.SelectedItem Is Nothing) Then
   FolderCount = TreeView1.Nodes.Count
  End If
 End If
End Property

Public Property Get SubFolderCount() As Long
 If mStyle = ESTreeFolder Then
  SubFolderCount = TreeView1.SelectedItem.Children
 End If
End Property

Public Property Get MousePointer() As MSComctlLib.MousePointerConstants
 MousePointer = ListView1.MousePointer
End Property
Public Property Let MousePointer(ByVal NewValue As MSComctlLib.MousePointerConstants)
 ListView1.MousePointer = NewValue
 TreeView1.MousePointer = NewValue
 PropertyChanged "MousePointer"
End Property

Public Property Get MouseIcon() As StdPicture
 Set MouseIcon = ListView1.MouseIcon
End Property
Public Property Set MouseIcon(ByVal NewValue As StdPicture)
 Set ListView1.MouseIcon = NewValue
 Set TreeView1.MouseIcon = NewValue
 PropertyChanged "MouseIcon"
End Property

Public Property Get OleDragMode() As MSComctlLib.OLEDragConstants
  OleDragMode = ListView1.OleDragMode
End Property
Public Property Let OleDragMode(ByVal NewValue As MSComctlLib.OLEDragConstants)
  ListView1.OleDragMode = NewValue
  TreeView1.OleDragMode = NewValue
  PropertyChanged "OleDragMode"
End Property

Public Property Get OleDropMode() As MSComctlLib.OLEDropConstants
  OleDropMode = ListView1.OleDropMode
End Property
Public Property Let OleDropMode(ByVal NewValue As MSComctlLib.OLEDropConstants)
  ListView1.OleDropMode = NewValue
  TreeView1.OleDropMode = NewValue
  PropertyChanged "OleDropMode"
End Property

Public Property Get SingleSel() As Boolean
  SingleSel = TreeView1.SingleSel
End Property
Public Property Let SingleSel(ByVal NewValue As Boolean)
  TreeView1.SingleSel = NewValue
  PropertyChanged "SingleSel"
End Property

Public Property Get SortOrder() As ListSortOrderConstants
 SortOrder = ListView1.SortOrder
End Property
Public Property Let SortOrder(ByVal NewValue As ListSortOrderConstants)
 ListView1.SortOrder = NewValue
 PropertyChanged "SortOrder"
End Property

Public Property Get Sorted() As Boolean
  Sorted = ListView1.Sorted
End Property
Public Property Let Sorted(ByVal NewValue As Boolean)
  ListView1.Sorted = NewValue
  PropertyChanged "Sorted"
End Property

Public Property Get SortKey() As EReportColumn
 SortKey = (ListView1.SortKey + 1)
End Property
Public Property Let SortKey(ByVal NewValue As EReportColumn)
 ListView1.SortKey = (NewValue - 1)
 PropertyChanged "SortKey"
End Property

Public Property Get HideFavorites() As Boolean
  HideFavorites = mHideFavorites
End Property
Public Property Let HideFavorites(ByVal NewValue As Boolean)
  mHideFavorites = NewValue
  PropertyChanged "HideFavorites"
  TreeView1.Nodes.Clear
  Call Refresh
End Property

Public Property Get HideMyDocuments() As Boolean
  HideMyDocuments = mHideMyDocuments
End Property
Public Property Let HideMyDocuments(ByVal NewValue As Boolean)
  mHideMyDocuments = NewValue
  PropertyChanged "HideMyDocuments"
  TreeView1.Nodes.Clear
  Call Refresh
End Property

Public Property Get HideNetwork() As Boolean
  HideNetwork = mHideNetwork
End Property
Public Property Let HideNetwork(ByVal NewValue As Boolean)
  mHideNetwork = NewValue
  PropertyChanged "HideNetwork"
  TreeView1.Nodes.Clear
  Call Refresh
End Property

Public Property Get HideControlPanel() As Boolean
  HideControlPanel = mHideControlPanel
End Property
Public Property Let HideControlPanel(ByVal NewValue As Boolean)
  mHideControlPanel = NewValue
  PropertyChanged "HideControlPanel"
  TreeView1.Nodes.Clear
  Call Refresh
End Property
