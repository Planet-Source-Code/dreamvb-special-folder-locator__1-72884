Attribute VB_Name = "Tools"
Option Explicit

Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function ShellExecuteEx Lib "shell32.dll" (ByRef s As SHELLEXECUTEINFO) As Long

'Listview Consts
Private Const LVM_FIRST As Long = &H1000
Private Const LVM_SETCOLUMNWIDTH As Long = (LVM_FIRST + 30)
Private Const LVSCW_AUTOSIZE_USEHEADER As Long = -2

Private Const SW_SHOW = 5
Private Const SEE_MASK_INVOKEIDLIST = &HC

Private Type SHELLEXECUTEINFO
    cbSize As Long
    fMask As Long
    hwnd As Long
    lpVerb As String
    lpFile As String
    lpParameters As String
    lpDirectory As String
    nShow As Long
    hInstApp As Long
    lpIDList As Long
    lpClass As String
    hkeyClass As Long
    dwHotKey As Long
    hIcon As Long
    hProcess As Long
End Type

Public Function FixPath(ByVal lPath As String) As String
    If Right$(lPath, 1) = "\" Then
        FixPath = lPath
    Else
        FixPath = lPath & "\"
    End If
End Function

Public Sub lvSizeColumns(lv As ListView)
Dim Counter As Long
    'Resizes Listview Column Headers
    For Counter = 0 To lv.ColumnHeaders.Count - 1
        Call SendMessage(lv.hwnd, LVM_SETCOLUMNWIDTH, Counter, _
        ByVal LVSCW_AUTOSIZE_USEHEADER)
    Next Counter
End Sub

Public Sub OpenUrl(ByVal URL As String)
Dim Ret As Long
    Ret = ShellExecute(frmmain.hwnd, "open", URL, vbNullString, vbNullString, 1)
End Sub

Public Sub DisplayFileProperties(ByVal sFullFileAndPathName As String)
Dim Ret As Long
Dim shInfo As SHELLEXECUTEINFO

    With shInfo
        .cbSize = LenB(shInfo)
        .lpFile = sFullFileAndPathName
        .nShow = SW_SHOW
        .fMask = SEE_MASK_INVOKEIDLIST
        .lpVerb = "properties"
    End With

    Ret = ShellExecuteEx(shInfo)
End Sub

