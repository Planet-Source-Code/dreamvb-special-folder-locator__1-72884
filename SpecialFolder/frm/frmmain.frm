VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmmain 
   Caption         =   "Special Folder Locater"
   ClientHeight    =   3825
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   8145
   LinkTopic       =   "Form1"
   ScaleHeight     =   3825
   ScaleWidth      =   8145
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ListView LstV 
      Height          =   750
      Left            =   -15
      TabIndex        =   2
      Top             =   405
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   1323
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Folder Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Read Only"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Hidden"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Folder Path"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Folder CLSID"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   8145
      _ExtentX        =   14367
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "OPEN"
            Object.ToolTipText     =   "Open Selected Folders"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "PROP"
            Object.ToolTipText     =   "Folder Properties"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "FIND"
            Object.ToolTipText     =   "Find"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "SCUT"
            Object.ToolTipText     =   "Create Shortcut On Desktop"
            ImageIndex      =   5
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   255
      Top             =   2085
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":0352
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":06A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":09F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":0D48
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar sBar1 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   0
      Top             =   3525
      Width           =   8145
      _ExtentX        =   14367
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   13864
         EndProperty
      EndProperty
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   1
      X1              =   0
      X2              =   405
      Y1              =   375
      Y2              =   375
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      Index           =   0
      X1              =   0
      X2              =   405
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuOpenFol 
         Caption         =   "Open Selected Folders"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuProp 
         Caption         =   "Folder Properties"
      End
      Begin VB.Menu Mnushortcut 
         Caption         =   "Create Shortcut On Desktop"
      End
      Begin VB.Menu mnuBlank1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "Edit"
      Begin VB.Menu mnuFolCpy 
         Caption         =   "Copy Folder Path"
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuCpyItem 
         Caption         =   "Copy Item"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuBlank3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFind 
         Caption         =   "Find"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuBlank2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSelAll 
         Caption         =   "Select &All"
         Shortcut        =   ^A
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuReadme 
         Caption         =   "Readme"
      End
      Begin VB.Menu mnuBlank4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SHGetFolderPath Lib "shell32.dll" Alias "SHGetFolderPathA" (ByVal hwnd As Long, ByVal csidl As Long, ByVal hToken As Long, ByVal dwFlags As Long, ByVal pszPath As String) As Long

Private Enum SpecilFolders
    CSIDL_ADMINTOOLS = &H30
    CSIDL_APPDATA = &H1A
    CSIDL_CDBURN_AREA = &H3B
    CSIDL_COMMON_ADMINTOOLS = &H2F
    CSIDL_COMMON_APPDATA = &H23
    CSIDL_COMMON_DESKTOPDIRECTORY = &H19
    CSIDL_COMMON_DOCUMENTS = &H2E
    CSIDL_COMMON_FAVORITES = &H1F
    CSIDL_COMMON_MUSIC = &H35
    CSIDL_COMMON_PICTURES = &H36
    CSIDL_COMMON_STARTMENU = &H16
    CSIDL_COMMON_PROGRAMS = &H17
    CSIDL_COMMON_STARTUP = &H18
    CSIDL_COMMON_TEMPLATES = &H2D
    CSIDL_COMMON_VIDEO = &H37
    CSIDL_COOKIES = &H21
    CSIDL_DESKTOPDIRECTORY = &H10
    CSIDL_FAVORITES = &H6
    CSIDL_FONTS = &H14
    CSIDL_HISTORY = &H22
    CSIDL_LOCAL_APPDATA = &H1C
    CSIDL_PERSONAL = &H5
    CSIDL_MYMUSIC = &HD
    CSIDL_MYPICTURES = &H27
    CSIDL_MYVIDEO = &HE
    CSIDL_NETHOOD = &H13
    CSIDL_PRINTHOOD = &H1B
    CSIDL_PROFILE = &H28
    CSIDL_PROGRAM_FILES = &H26
    CSIDL_PROGRAM_FILES_COMMON = &H2B
    CSIDL_RECENT = &H8
    CSIDL_SENDTO = &H9
    CSIDL_STARTMENU = &HB
    CSIDL_PROGRAMS = &H2
    CSIDL_STARTUP = &H7
    CSIDL_SYSTEM = &H25
    CSIDL_TEMPLATES = &H15
    CSIDL_INTERNET_CACHE = &H20
    CSIDL_WINDOWS = &H24
End Enum

Private Type SPFolderInfo
    sFolderName As String
    sFolderPath As String
    sCLSID As String
End Type

Private mFolderInfo(1 To 39) As SPFolderInfo
Private sFolder As String
Private mButton As MouseButtonConstants

Private Sub ShortCut()
Dim oWs
Dim oLink
Dim sLinkFile As String

    'Create a Windows short cut
    Set oWs = CreateObject("WScript.Shell")
    'Create shortcut name
    sLinkFile = FixPath(GetSpFolder(CSIDL_DESKTOPDIRECTORY)) & LstV.SelectedItem.Text & ".Lnk"
    'Create the shortcut
    Set oLink = oWs.CreateShortcut(sLinkFile)
    'shortcut location
    oLink.TargetPath = sFolder
    'Create the shortcut
    Call oLink.Save
    
    Set oWs = Nothing
    Set oLink = Nothing
End Sub

Private Sub CopyListItem()
Dim lItem As ListItem
Dim sItem As String
    'Set the item to copy
    Set lItem = LstV.ListItems(LstV.SelectedItem.Index)
    'Store the item data
    sItem = lItem.Text & vbTab & lItem.SubItems(1) & vbTab & lItem.SubItems(2) & vbTab & lItem.SubItems(3) & vbTab & lItem.SubItems(4)
    'Copy to clipboard
    Call Clipboard.Clear
    Call Clipboard.SetText(sItem, vbCFText)
    Set lItem = Nothing
End Sub

Private Function BoolToYesNo(ByVal iVal As Boolean) As String
    'Turns true / false into Yes / No
    BoolToYesNo = IIf(iVal, "Yes", "No")
End Function

Private Sub SelectAllItems(iSelect As Boolean)
Dim lItem As ListItem
    'Select all items
    For Each lItem In LstV.ListItems
        lItem.Selected = iSelect
    Next lItem
End Sub

Private Sub ClickItem(ByVal Index As Integer)
Dim lItem As ListItem
On Error Resume Next
    Set lItem = LstV.ListItems(Index)
    lItem.Selected = True
    Call LstV_ItemClick(lItem)
    Call LstV.SetFocus
    'Destroy Listitem
    Set lItem = Nothing
End Sub

Private Function GetSpFolder(ByVal CLSID As SpecilFolders) As String
Dim Ret As Long
Dim sBuff As String
    
    'Create space to hold folder location
    sBuff = Space$(260)
    
    Ret = SHGetFolderPath(frmmain.hwnd, CLSID, 0, 1, sBuff)
    'Strip away chr 0
    GetSpFolder = Left(sBuff, InStr(sBuff, Chr(0)) - 1)
    'Clear up
    sBuff = vbNullString
    
End Function

Private Sub Form_Load()
Dim Count As Integer

    'Set up mFolderInfo array with Special folder info
    mFolderInfo(1).sFolderName = "Administrative Tools"
    mFolderInfo(1).sFolderPath = GetSpFolder(CSIDL_ADMINTOOLS)
    mFolderInfo(1).sCLSID = "CSIDL_ADMINTOOLS"
    '
    mFolderInfo(2).sFolderName = "Application Data"
    mFolderInfo(2).sFolderPath = GetSpFolder(CSIDL_APPDATA)
    mFolderInfo(2).sCLSID = "CSIDL_APPDATA"
    '
    mFolderInfo(3).sFolderName = "CD Burning"
    mFolderInfo(3).sFolderPath = GetSpFolder(CSIDL_CDBURN_AREA)
    mFolderInfo(3).sCLSID = "CSIDL_CDBURN_AREA"
    '
    mFolderInfo(4).sFolderName = "Common Administrative Tools"
    mFolderInfo(4).sFolderPath = GetSpFolder(CSIDL_COMMON_ADMINTOOLS)
    mFolderInfo(4).sCLSID = "CSIDL_COMMON_ADMINTOOLS"
    '
    mFolderInfo(5).sFolderName = "Common Application Data"
    mFolderInfo(5).sFolderPath = GetSpFolder(CSIDL_COMMON_APPDATA)
    mFolderInfo(5).sCLSID = "CSIDL_COMMON_APPDATA"
    '
    mFolderInfo(6).sFolderName = "Common Desktop"
    mFolderInfo(6).sFolderPath = GetSpFolder(CSIDL_COMMON_DESKTOPDIRECTORY)
    mFolderInfo(6).sCLSID = "CSIDL_COMMON_DESKTOPDIRECTORY"
    '
    mFolderInfo(7).sFolderName = "Common Documents"
    mFolderInfo(7).sFolderPath = GetSpFolder(CSIDL_COMMON_DOCUMENTS)
    mFolderInfo(7).sCLSID = "CSIDL_COMMON_DOCUMENTS"
    '
    mFolderInfo(8).sFolderName = "Common Favorites"
    mFolderInfo(8).sFolderPath = GetSpFolder(CSIDL_COMMON_FAVORITES)
    mFolderInfo(8).sCLSID = "CSIDL_COMMON_FAVORITES"
    '
    mFolderInfo(9).sFolderName = "Common Music"
    mFolderInfo(9).sFolderPath = GetSpFolder(CSIDL_COMMON_MUSIC)
    mFolderInfo(9).sCLSID = "CSIDL_COMMON_MUSIC"
    '
    mFolderInfo(10).sFolderName = "Common Pictures"
    mFolderInfo(10).sFolderPath = GetSpFolder(CSIDL_COMMON_PICTURES)
    mFolderInfo(10).sCLSID = "CSIDL_COMMON_PICTURES"
    '
    mFolderInfo(11).sFolderName = "Common Start Menu"
    mFolderInfo(11).sFolderPath = GetSpFolder(CSIDL_COMMON_STARTMENU)
    mFolderInfo(11).sCLSID = "CSIDL_COMMON_STARTMENU"
    '
    mFolderInfo(12).sFolderName = "Common Start Menu Programs"
    mFolderInfo(12).sFolderPath = GetSpFolder(CSIDL_COMMON_PROGRAMS)
    mFolderInfo(12).sCLSID = "CSIDL_COMMON_PROGRAMS"
    '
    mFolderInfo(13).sFolderName = "Common Startup"
    mFolderInfo(13).sFolderPath = GetSpFolder(CSIDL_COMMON_STARTUP)
    mFolderInfo(13).sCLSID = "CSIDL_COMMON_STARTUP"
    '
    mFolderInfo(14).sFolderName = "Common Templates"
    mFolderInfo(14).sFolderPath = GetSpFolder(CSIDL_COMMON_TEMPLATES)
    mFolderInfo(14).sCLSID = "CSIDL_COMMON_TEMPLATES"
    '
    mFolderInfo(15).sFolderName = "Common Video"
    mFolderInfo(15).sFolderPath = GetSpFolder(CSIDL_COMMON_VIDEO)
    mFolderInfo(15).sCLSID = "CSIDL_COMMON_VIDEO"
    '
    mFolderInfo(16).sFolderName = "Cookies"
    mFolderInfo(16).sFolderPath = GetSpFolder(CSIDL_COOKIES)
    mFolderInfo(16).sCLSID = "CSIDL_COOKIES"
    '
    mFolderInfo(17).sFolderName = "Desktop"
    mFolderInfo(17).sFolderPath = GetSpFolder(CSIDL_DESKTOPDIRECTORY)
    mFolderInfo(17).sCLSID = "CSIDL_DESKTOPDIRECTORY"
    '
    mFolderInfo(18).sFolderName = "Favorites"
    mFolderInfo(18).sFolderPath = GetSpFolder(CSIDL_FAVORITES)
    mFolderInfo(18).sCLSID = "CSIDL_FAVORITES"
    '
    mFolderInfo(19).sFolderName = "Fonts"
    mFolderInfo(19).sFolderPath = GetSpFolder(CSIDL_FONTS)
    mFolderInfo(19).sCLSID = "CSIDL_FONTS"
    '
    mFolderInfo(20).sFolderName = "History"
    mFolderInfo(20).sFolderPath = GetSpFolder(CSIDL_HISTORY)
    mFolderInfo(20).sCLSID = "CSIDL_HISTORY"
    '
    mFolderInfo(21).sFolderName = "Local Application Data"
    mFolderInfo(21).sFolderPath = GetSpFolder(CSIDL_LOCAL_APPDATA)
    mFolderInfo(21).sCLSID = "CSIDL_LOCAL_APPDATA"
    '
    mFolderInfo(22).sFolderName = "My Documents"
    mFolderInfo(22).sFolderPath = GetSpFolder(CSIDL_PERSONAL)
    mFolderInfo(22).sCLSID = "CSIDL_PERSONAL"
    '
    mFolderInfo(23).sFolderName = "My Music"
    mFolderInfo(23).sFolderPath = GetSpFolder(CSIDL_MYMUSIC)
    mFolderInfo(23).sCLSID = "CSIDL_MYMUSIC"
    '
    mFolderInfo(24).sFolderName = "My Pictures"
    mFolderInfo(24).sFolderPath = GetSpFolder(CSIDL_MYPICTURES)
    mFolderInfo(24).sCLSID = "CSIDL_MYPICTURES"
    '
    mFolderInfo(25).sFolderName = "My Video"
    mFolderInfo(25).sFolderPath = GetSpFolder(CSIDL_MYVIDEO)
    mFolderInfo(25).sCLSID = "CSIDL_MYVIDEO"
    '
    mFolderInfo(26).sFolderName = "NetHood"
    mFolderInfo(26).sFolderPath = GetSpFolder(CSIDL_NETHOOD)
    mFolderInfo(26).sCLSID = "CSIDL_NETHOOD"
    '
    mFolderInfo(27).sFolderName = "PrintHood"
    mFolderInfo(27).sFolderPath = GetSpFolder(CSIDL_PRINTHOOD)
    mFolderInfo(27).sCLSID = "CSIDL_PRINTHOOD"
    '
    mFolderInfo(28).sFolderName = "Profile Folder"
    mFolderInfo(28).sFolderPath = GetSpFolder(CSIDL_PROFILE)
    mFolderInfo(28).sCLSID = "CSIDL_PROFILE"
    '
    mFolderInfo(29).sFolderName = "Program Files"
    mFolderInfo(29).sFolderPath = GetSpFolder(CSIDL_PROGRAM_FILES)
    mFolderInfo(29).sCLSID = "CSIDL_PROGRAM_FILES"
    '
    mFolderInfo(30).sFolderName = "Program Files Common"
    mFolderInfo(30).sFolderPath = GetSpFolder(CSIDL_PROGRAM_FILES_COMMON)
    mFolderInfo(30).sCLSID = "CSIDL_PROGRAM_FILES_COMMON"
    '
    mFolderInfo(31).sFolderName = "Recent"
    mFolderInfo(31).sFolderPath = GetSpFolder(CSIDL_RECENT)
    mFolderInfo(31).sCLSID = "CSIDL_RECENT"
    '
    mFolderInfo(32).sFolderName = "Send To"
    mFolderInfo(32).sFolderPath = GetSpFolder(CSIDL_SENDTO)
    mFolderInfo(32).sCLSID = "CSIDL_SENDTO"
    '
    mFolderInfo(33).sFolderName = "Start Menu"
    mFolderInfo(33).sFolderPath = GetSpFolder(CSIDL_STARTMENU)
    mFolderInfo(33).sCLSID = "CSIDL_STARTMENU"
    '
    mFolderInfo(34).sFolderName = "Start Menu Programs"
    mFolderInfo(34).sFolderPath = GetSpFolder(CSIDL_PROGRAMS)
    mFolderInfo(34).sCLSID = "CSIDL_PROGRAMS"
    '
    mFolderInfo(35).sFolderName = "Startup"
    mFolderInfo(35).sFolderPath = GetSpFolder(CSIDL_STARTUP)
    mFolderInfo(35).sCLSID = "CSIDL_STARTUP"
    '
    mFolderInfo(36).sFolderName = "System Directory"
    mFolderInfo(36).sFolderPath = GetSpFolder(CSIDL_SYSTEM)
    mFolderInfo(36).sCLSID = "CSIDL_SYSTEM"
    '
    mFolderInfo(37).sFolderName = "Templates"
    mFolderInfo(37).sFolderPath = GetSpFolder(CSIDL_TEMPLATES)
    mFolderInfo(37).sCLSID = "CSIDL_TEMPLATES"
    '
    mFolderInfo(38).sFolderName = "Temporary Internet Files"
    mFolderInfo(38).sFolderPath = GetSpFolder(CSIDL_INTERNET_CACHE)
    mFolderInfo(38).sCLSID = "CSIDL_INTERNET_CACHE"
    '
    mFolderInfo(39).sFolderName = "Windows Directory"
    mFolderInfo(39).sFolderPath = GetSpFolder(CSIDL_WINDOWS)
    mFolderInfo(39).sCLSID = "CSIDL_WINDOWS"
    
    For Count = 1 To 39
        'Add folder name
        LstV.ListItems.Add , , mFolderInfo(Count).sFolderName, 1, 1
        'Add ReadOnly
        LstV.ListItems(LstV.ListItems.Count).SubItems(1) = BoolToYesNo(CBool(GetAttr(mFolderInfo(Count).sFolderPath) And vbReadOnly))
        'Add Hidden
        LstV.ListItems(LstV.ListItems.Count).SubItems(2) = BoolToYesNo(CBool(GetAttr(mFolderInfo(Count).sFolderPath) And vbHidden))
        'Add Folder Path
        LstV.ListItems(LstV.ListItems.Count).SubItems(3) = mFolderInfo(Count).sFolderPath
        'Add CLSID
        LstV.ListItems(LstV.ListItems.Count).SubItems(4) = mFolderInfo(Count).sCLSID
    Next Count
    
    'Resize column headers
    Call lvSizeColumns(LstV)
    sBar1.Panels(1).Text = LstV.ListItems.Count & " Special Folders Found"
    'Click the first item
    Call ClickItem(1)
End Sub

Private Sub Form_Resize()
On Error Resume Next
    'Make fake 3d line
    Line1(0).X2 = frmmain.ScaleWidth
    Line1(1).X2 = frmmain.ScaleWidth
    'Resize litsview
    LstV.Width = (frmmain.ScaleWidth - LstV.Left)
    LstV.Height = (frmmain.ScaleHeight - sBar1.Height - LstV.Top)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmmain = Nothing
End Sub

Private Sub LstV_DblClick()
    If (mButton <> vbLeftButton) Then
        Exit Sub
    Else
        Call OpenUrl(sFolder)
    End If
End Sub

Private Sub LstV_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error Resume Next
    'Store folder location
    sFolder = LstV.ListItems(LstV.SelectedItem.Index).SubItems(3)
End Sub

Private Sub LstV_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    mButton = Button
End Sub

Private Sub mnuAbout_Click()
    MsgBox frmmain.Caption & " Ver 1.0" _
    & vbCrLf & "Make by DreamVB", vbInformation, "About"
End Sub

Private Sub mnuCpyItem_Click()
    Call CopyListItem
End Sub

Private Sub mnuExit_Click()
    'Unload program
    Call Unload(frmmain)
End Sub

Private Sub mnuFind_Click()
Dim sFind As String
Dim lItem As ListItem
    
    'Deselect all items
    Call SelectAllItems(False)
    
    sFind = Trim$(InputBox$("Enter a string to serach", "Find"))
    
    If Len(sFind) Then
        For Each lItem In LstV.ListItems
            If InStr(1, lItem.Text, sFind, vbTextCompare) Then
                'Click the index if found
                Call ClickItem(lItem.Index)
                'Move to the found index
                Call LstV.ListItems(lItem.Index).EnsureVisible
                Exit For
            End If
        Next lItem
    End If
    
End Sub

Private Sub mnuFolCpy_Click()
    Call Clipboard.Clear
    Call Clipboard.SetText(sFolder, vbCFText)
End Sub

Private Sub mnuOpenFol_Click()
Dim lItem As ListItem

    For Each lItem In LstV.ListItems
        If (lItem.Selected) Then
            'Open the special folder
            Call OpenUrl(lItem.SubItems(3))
        End If
    Next lItem
    
End Sub

Private Sub mnuProp_Click()
    'Show folder properties
    Call DisplayFileProperties(sFolder)
End Sub

Private Sub mnuReadme_Click()
    'Show readme file
    Call OpenUrl(FixPath(App.Path) & "Readme.txt")
End Sub

Private Sub mnuSelAll_Click()
    'Select all items in Listview
    Call SelectAllItems(True)
End Sub

Private Sub Mnushortcut_Click()
    Call ShortCut
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "OPEN"
            Call mnuOpenFol_Click
        Case "PROP"
            Call mnuProp_Click
        Case "FIND"
            Call mnuFind_Click
        Case "SCUT"
            If MsgBox("Do you want to create a shortcut on your desktop for:" & vbCrLf & vbCrLf & "'" _
            & LstV.SelectedItem.Text & "'", vbYesNo Or vbQuestion, frmmain.Caption) = vbYes Then
                Call Mnushortcut_Click
            End If
    End Select
End Sub
