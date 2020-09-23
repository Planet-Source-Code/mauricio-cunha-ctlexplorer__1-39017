VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmTest 
   Caption         =   "Test of CtlExplorer by Mauricio Cunha"
   ClientHeight    =   3705
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   6885
   LinkTopic       =   "Form1"
   ScaleHeight     =   3705
   ScaleWidth      =   6885
   StartUpPosition =   3  'Windows Default
   Begin ctlexplorerLib.ctlExplorer ctlExplorer2 
      Height          =   2655
      Left            =   3120
      TabIndex        =   3
      Top             =   240
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   4683
      Arrange         =   2
      LabelWrap       =   -1  'True
      MouseIcon       =   "FrmTest.frx":0000
      Path            =   "D:\mp3\Nacionais"
      Style           =   1
      View            =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ProgressBar PBar 
      Height          =   375
      Left            =   3120
      TabIndex        =   2
      Top             =   2880
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSComctlLib.StatusBar SBar 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   3450
      Width           =   6885
      _ExtentX        =   12144
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   6588
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin ctlexplorerLib.ctlExplorer ctlExplorer1 
      Height          =   3255
      Left            =   0
      TabIndex        =   0
      Top             =   240
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   5741
      Arrange         =   2
      LabelWrap       =   -1  'True
      MouseIcon       =   "FrmTest.frx":001C
      ShowFolders     =   0   'False
      Filter          =   ""
      Path            =   "D:\mp3\Nacionais"
      View            =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu mnuFileTop 
      Caption         =   "&File"
      Begin VB.Menu mnuFileSub 
         Caption         =   "Show F&olders"
         Index           =   0
      End
      Begin VB.Menu sp0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewTop 
         Caption         =   "View"
         Begin VB.Menu mnuViewSub 
            Caption         =   "Icons"
            Index           =   0
         End
         Begin VB.Menu mnuViewSub 
            Caption         =   "Small Icons"
            Index           =   1
         End
         Begin VB.Menu mnuViewSub 
            Caption         =   "List"
            Index           =   2
         End
         Begin VB.Menu mnuViewSub 
            Caption         =   "Details"
            Index           =   3
         End
      End
      Begin VB.Menu mnuSortTop 
         Caption         =   "Sort by"
         Begin VB.Menu mnuSortSub 
            Caption         =   "File"
            Index           =   0
         End
         Begin VB.Menu mnuSortSub 
            Caption         =   "Size"
            Index           =   1
         End
         Begin VB.Menu mnuSortSub 
            Caption         =   "Type"
            Index           =   2
         End
         Begin VB.Menu mnuSortSub 
            Caption         =   "Modified"
            Index           =   3
         End
         Begin VB.Menu mnuSortSub 
            Caption         =   "Created"
            Index           =   4
         End
      End
      Begin VB.Menu sp1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSelectAll 
         Caption         =   "&Select All"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuFileMultiSelect 
         Caption         =   "&Multi-Select"
      End
      Begin VB.Menu sp2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileAutoSize 
         Caption         =   "&Auto Size Columns"
      End
      Begin VB.Menu mnuFileCheckBoxes 
         Caption         =   "CheckBoxes"
      End
      Begin VB.Menu sp3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileArrangeTop 
         Caption         =   "&Arrange"
         Begin VB.Menu mnuFileArrangeSub 
            Caption         =   "&None"
            Index           =   0
         End
         Begin VB.Menu mnuFileArrangeSub 
            Caption         =   "&Left"
            Index           =   1
         End
         Begin VB.Menu mnuFileArrangeSub 
            Caption         =   "&Top"
            Index           =   2
         End
      End
      Begin VB.Menu mnuFileFullRow 
         Caption         =   "F&ull row select"
      End
   End
   Begin VB.Menu mnuFolderTop 
      Caption         =   "F&older"
      Begin VB.Menu mnuHideControlPanel 
         Caption         =   "Hide Control Panel"
      End
      Begin VB.Menu mnuHideMyDocuments 
         Caption         =   "Hide My Documents"
      End
      Begin VB.Menu mnuHideNetwoork 
         Caption         =   "Hide Netwoork"
      End
      Begin VB.Menu mnuHideFavorites 
         Caption         =   "Hide Favorites"
      End
      Begin VB.Menu st1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFolderCheckBoxes 
         Caption         =   "&Checkboxes"
      End
   End
   Begin VB.Menu mnuHelpTop 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpSub 
         Caption         =   "&About..."
         Index           =   0
      End
      Begin VB.Menu mnuHelpSub 
         Caption         =   "&Home-page of developer"
         Index           =   1
      End
   End
End
Attribute VB_Name = "FrmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'For hyperlink...
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Const SW_SHOWNORMAL = 1

Private Sub ctlExplorer1_Change()
 SBar.Panels(1).Text = ctlExplorer1.Path
 SBar.Panels(3).Text = ctlExplorer1.FolderCount & " folder(s)"
End Sub

Private Sub ctlExplorer1_DriveNotReady(DriveName As String)
 SBar.Panels(1).Text = "*** Drive " & DriveName & " not ready ***"
End Sub

Private Sub ctlExplorer1_FolderCheck(FolderName As String, Value As Boolean)
Dim M As Long
 For M = 1 To ctlExplorer1.FolderCount
  ctlExplorer1.FolderChecked(M) = False
 Next M
 
  
 
 ctlExplorer2.Path = FolderName
 If ctlExplorer2.CheckBoxes = True And ctlExplorer2.FileCount > 0 Then
  For M = 1 To ctlExplorer2.FileCount
   ctlExplorer2.FileChecked(M) = Value
  Next M
 End If
End Sub

Private Sub ctlExplorer1_FolderClick(FolderName As String)
 With ctlExplorer2
  .Sorted = False
  .Path = FolderName
 End With
End Sub

Private Sub ctlExplorer1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  If Button = 2 Then PopupMenu mnuFolderTop
End Sub

Private Sub ctlExplorer1_Progress(Percent As Integer)
  PBar.Value = Percent
  Call Form_Resize
End Sub

Private Sub ctlExplorer1_RootClick(RootType As ERootFolder)
 With ctlExplorer2
  .Clear
  .Sorted = False
  .ListRootFolder RootType
  .LabelEdit = tvwManual
  SBar.Panels(2).Text = ""
  SBar.Panels(3).Text = ""
 End With

Select Case RootType
 Case ERFControlPanel
  SBar.Panels(1).Text = "Change your settings in this machine using Control Panel"
 Case ERFMyComputer
  SBar.Panels(1).Text = "List all avaliables dispositives on local machine"
 Case ERFWorkSpace
  SBar.Panels(1).Text = "All dispositives for this desktop"
End Select
End Sub

Private Sub ctlExplorer1_RootCollapse(RootType As ERootFolder)
  Select Case RootType
   Case ERFControlPanel
    SBar.Panels(1).Text = "Change your settings in this machine"
   Case ERFMyComputer
    SBar.Panels(1).Text = "List all avaliables dispositives on local machine"
   Case ERFWorkSpace
    SBar.Panels(1).Text = "All dispositives for this desktop"
  End Select
End Sub

Private Sub ctlExplorer2_Change()
 SBar.Panels(2).Text = ctlExplorer2.FileCount & " file(s)"
End Sub

Private Sub ctlExplorer2_ColumnClick(Column As EReportColumn)
  With ctlExplorer2
   .SortOrder = IIf(.SortOrder = 0, 1, 0)
   .Sorted = True
  End With
End Sub

Private Sub ctlExplorer2_FileClick(FileName As String)
If ctlExplorer2.MultiSelect = False Then
 SBar.Panels(1).Text = FileName
Else
 SBar.Panels(1).Text = ctlExplorer2.SelectedCount & " file(s) selected(s)"
End If
End Sub

Private Sub ctlExplorer2_FolderClick(FolderName As String)
If ctlExplorer2.MultiSelect = False Then
 SBar.Panels(1).Text = FolderName
Else
 SBar.Panels(1).Text = ctlExplorer2.SelectedCount & " file(s) selected(s)"
End If
End Sub

Private Sub ctlExplorer2_FolderDblClick(FolderName As String)
 ctlExplorer1.Path = FolderName
End Sub

Private Sub ctlExplorer2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
 If Button = 2 Then PopupMenu mnuFileTop
End Sub

Private Sub ctlExplorer2_Progress(Percent As Integer)
 PBar.Value = Percent
End Sub

Private Sub ctlExplorer2_RootDblClick(RootType As ERootFolder)
 Call ctlExplorer1_RootClick(RootType)
End Sub

Private Sub Form_Load()
Dim M As Long
 
 mnuViewSub(ctlExplorer2.View).Checked = True
 mnuFileSub(0).Checked = ctlExplorer2.ShowFolders
 mnuSortSub(ctlExplorer2.SortKey - 1).Checked = True
 ctlExplorer1.Path = CurDir
 ctlExplorer2.Path = ctlExplorer1.Path
 mnuFileMultiSelect.Checked = ctlExplorer2.MultiSelect
 mnuFileArrangeSub(ctlExplorer1.Arrange).Checked = True
 mnuFileFullRow.Checked = ctlExplorer2.FullRowSelect
 
 mnuHideControlPanel.Checked = ctlExplorer1.HideControlPanel
 mnuHideFavorites.Checked = ctlExplorer1.HideFavorites
 mnuHideMyDocuments.Checked = ctlExplorer1.HideMyDocuments
 mnuHideNetwoork.Checked = ctlExplorer1.HideNetwork
 
 For M = 1 To ctlExplorer2.ColumnCount
  ctlExplorer2.ColumnWidth(M) = GetSetting(App.EXEName, "column" & M, "width", 1200)
 Next M
End Sub

Private Sub Form_Resize()
'On Error Resume Next
 ctlExplorer1.Move 0, 0, ctlExplorer1.Width, Me.ScaleHeight - SBar.Height
 ctlExplorer2.Move ctlExplorer1.Width, ctlExplorer1.Top, Me.ScaleWidth - ctlExplorer1.Width, Me.ScaleHeight - (SBar.Height + PBar.Height)
 PBar.Move ctlExplorer2.Left, ctlExplorer2.Top + ctlExplorer2.Height, ctlExplorer2.Width

End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
  ChDir ctlExplorer1.Path

 Dim M As Long
 For M = 1 To ctlExplorer2.ColumnCount
  SaveSetting App.EXEName, "column" & M, "width", ctlExplorer2.ColumnWidth(M)
 Next M
End Sub

Private Sub mnuFileArrangeSub_Click(Index As Integer)
Dim L As Long

For L = 0 To mnuFileArrangeSub.Count - 1
 mnuFileArrangeSub(L).Checked = False
Next L

mnuFileArrangeSub(Index).Checked = True
ctlExplorer1.Arrange = Index
End Sub

Private Sub mnuFileAutoSize_Click()
 Me.ctlExplorer2.ColumnsAutoSize ERCName
End Sub

Private Sub mnuFileCheckBoxes_Click()
 mnuFileCheckBoxes.Checked = Not (mnuFileCheckBoxes.Checked)
 ctlExplorer2.CheckBoxes = mnuFileCheckBoxes.Checked
End Sub

Private Sub mnuFileFullRow_Click()
 mnuFileFullRow.Checked = Not (mnuFileFullRow.Checked)
 ctlExplorer2.FullRowSelect = mnuFileFullRow.Checked
End Sub

Private Sub mnuFileMultiSelect_Click()
 mnuFileMultiSelect.Checked = Not (mnuFileMultiSelect.Checked)
 ctlExplorer2.MultiSelect = mnuFileMultiSelect.Checked
 mnuFileSelectAll.Enabled = mnuFileMultiSelect.Checked
End Sub

Private Sub mnuFileSelectAll_Click()
Dim L As Long

For L = 1 To ctlExplorer2.FileCount
 ctlExplorer2.FileSelected(L) = True
Next L
End Sub

Private Sub mnuFileSub_Click(Index As Integer)
Select Case Index
 Case 0
  mnuFileSub(Index).Checked = Not (mnuFileSub(Index).Checked)
  ctlExplorer2.ShowFolders = mnuFileSub(Index).Checked
End Select
End Sub

Private Sub mnuFolderCheckBoxes_Click()
 mnuFolderCheckBoxes.Checked = Not (mnuFolderCheckBoxes.Checked)
 ctlExplorer1.CheckBoxes = mnuFolderCheckBoxes.Checked
End Sub

Private Sub mnuHelpSub_Click(Index As Integer)
If Index = 0 Then ctlExplorer1.AboutBox
If Index = 1 Then
 Dim iret As Long
 iret = ShellExecute(Me.hwnd, vbNullString, "http://www.mcunha98.cjb.net", vbNullString, CurDir, SW_SHOWNORMAL)
End If
End Sub

Private Sub mnuHideControlPanel_Click()
 mnuHideControlPanel.Checked = Not (mnuHideControlPanel.Checked)
 ctlExplorer1.HideControlPanel = mnuHideControlPanel.Checked
End Sub

Private Sub mnuHideFavorites_Click()
 mnuHideFavorites.Checked = Not (mnuHideFavorites.Checked)
 ctlExplorer1.HideFavorites = mnuHideFavorites.Checked
End Sub

Private Sub mnuHideMyDocuments_Click()
 mnuHideMyDocuments.Checked = Not (mnuHideMyDocuments.Checked)
 ctlExplorer1.HideMyDocuments = mnuHideMyDocuments.Checked
End Sub

Private Sub mnuHideNetwoork_Click()
 mnuHideNetwoork.Checked = Not (mnuHideNetwoork.Checked)
 ctlExplorer1.HideNetwork = mnuHideNetwoork.Checked
End Sub

Private Sub mnuSortSub_Click(Index As Integer)
Dim M As Integer

For M = 0 To mnuSortSub.Count - 1
 mnuSortSub(M).Checked = False
Next M

With ctlExplorer2
 .SortKey = Index + 1
 .SortOrder = IIf(.SortOrder = lvwAscending, lvwDescending, lvwAscending)
 .Sorted = True
End With

mnuSortSub(Index).Checked = True
End Sub

Private Sub mnuViewSub_Click(Index As Integer)
Dim M As Integer
 For M = 0 To mnuViewSub.Count - 1
  mnuViewSub(M).Checked = False
 Next M
 ctlExplorer2.View = Index
 mnuViewSub(Index).Checked = True
End Sub

