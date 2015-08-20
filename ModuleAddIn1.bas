Attribute VB_Name = "ModuleAddIn1"
Option Explicit
Private Type OPENFILENAME
    lStructSize       As Long
    hwndOwner         As Long
    hInstance         As Long
    lpstrFilter       As String
    lpstrCustomFilter As String
    nMaxCustFilter    As Long
    nFilterIndex      As Long
    lpstrFile         As String
    nMaxFile          As Long
    lpstrFileTitle    As String
    nMaxFileTitle     As Long
    lpstrInitialDir   As String
    lpstrTitle        As String
    flags             As Long
    nFileOffset       As Integer
    nFileExtension    As Integer
    lpstrDefExt       As String
    lCustData         As Long
    lpfnHook          As Long
    lpTemplateName    As String
End Type
Private Type BrowseInfo
    hwndOwner         As Long
    pIDLRoot          As Long
    pszDisplayName    As Long
    lpszTitle         As Long
    ulFlags           As Long
    lpfnCallback      As Long
    lParam            As Long
    iImage            As Long
End Type

'Private Const OFN_READONLY             As Long = &H1
'Private Const OFN_OVERWRITEPROMPT      As Long = &H2
Private Const OFN_HIDEREADONLY         As Long = &H4
'Private Const OFN_NOCHANGEDIR          As Long = &H8
'Private Const OFN_SHOWHELP             As Long = &H10
'Private Const OFN_ENABLEHOOK           As Long = &H20
'Private Const OFN_ENABLETEMPLATE       As Long = &H40
'Private Const OFN_ENABLETEMPLATEHANDLE As Long = &H80
'Private Const OFN_NOVALIDATE           As Long = &H100
Private Const OFN_ALLOWMULTISELECT     As Long = &H200
'Private Const OFN_EXTENSIONDIFFERENT   As Long = &H400
'Private Const OFN_PATHMUSTEXIST        As Long = &H800
'Private Const OFN_FILEMUSTEXIST        As Long = &H1000
'Private Const OFN_CREATEPROMPT         As Long = &H2000
'Private Const OFN_SHAREAWARE           As Long = &H4000
'Private Const OFN_NOREADONLYRETURN     As Long = &H8000
'Private Const OFN_NOTESTFILECREATE     As Long = &H10000
'Private Const OFN_NONETWORKBUTTON      As Long = &H20000
'Private Const OFN_NOLONGNAMES          As Long = &H40000
Private Const OFN_EXPLORER             As Long = &H80000
'Private Const OFN_NODEREFERENCELINKS   As Long = &H100000
'Private Const OFN_LONGNAMES            As Long = &H200000

'Private Const OFN_SHAREFALLTHROUGH     As Long = 2
'Private Const OFN_SHARENOWARN          As Long = 1
'Private Const OFN_SHAREWARN            As Long = 0

Private Const BrowseForFolders         As Long = &H1
'Private Const BrowseForComputers       As Long = &H1000
'Private Const BrowseForPrinters        As Long = &H2000
'Private Const BrowseForEverything      As Long = &H4000

'Private Const CSIDL_BITBUCKET          As Long = 10
'Private Const CSIDL_CONTROLS           As Long = 3
'Private Const CSIDL_DESKTOP            As Long = 0
Private Const CSIDL_DRIVES             As Long = 17
'Private Const CSIDL_FONTS              As Long = 20
'Private Const CSIDL_NETHOOD            As Long = 18
'Private Const CSIDL_NETWORK            As Long = 19
'Private Const CSIDL_PERSONAL           As Long = 5
'Private Const CSIDL_PRINTERS           As Long = 4
'Private Const CSIDL_PROGRAMS           As Long = 2
'Private Const CSIDL_RECENT             As Long = 8
'Private Const CSIDL_SENDTO             As Long = 9
'Private Const CSIDL_STARTMENU          As Long = 11

Private Const MAX_PATH                 As Long = 260

Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Private Declare Function GetModuleFileName Lib "kernel32" Alias "GetModuleFileNameA" (ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long

Private Declare Function SHBrowseForFolder Lib "shell32" (lpBI As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Public Declare Function SHGetSpecialFolderLocation Lib "shell32" (ByVal hwndOwner As Long, ByVal nFolder As Long, ListId As Long) As Long

Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long

Private Const MaxTitle As Long = 64, MaxFileCount As Integer = 512
Private iFileName As String, iTitle As String, iFilter As String, iInitDir As String, iExtention As String, iExtentionIndex As Integer
Private iReadOnly As Boolean, iMultiple As Boolean
Private hasInit As Boolean

Property Get ShowReadOnly() As Boolean
 ShowReadOnly = iReadOnly
End Property
Property Let ShowReadOnly(value As Boolean)
 init
 iReadOnly = value
End Property
Property Get MultipleSelect() As Boolean
 ShowReadOnly = iMultiple
End Property
Property Let MultipleSelect(value As Boolean)
 init
 iMultiple = value
End Property
Property Get fileName() As String
 fileName = iFileName
End Property
Property Let fileName(value As String)
 init
 iFileName = value
End Property
Property Get Title() As String
 Title = iTitle
End Property
Property Let Title(value As String)
 init
 iTitle = value
End Property
Property Get Filter() As String
 Filter = iFilter
End Property
Property Let Filter(value As String)
 init
 iFilter = value
End Property
Property Get InitDir() As String
 InitDir = iInitDir
End Property
Property Let InitDir(value As String)
 init
 iInitDir = value
End Property
Property Get Extention() As String
 Extention = iExtention
End Property
Property Let Extention(value As String)
 init
 iExtention = value
End Property
Property Get ExtentionIndex() As Integer
 ExtentionIndex = iExtentionIndex
End Property
Property Let ExtentionIndex(value As Integer)
 init
 iExtentionIndex = ExtentionIndex
End Property

Private Sub init()
 If hasInit Then Exit Sub
 iReadOnly = False: iMultiple = False: iInitDir = AppPath: iExtention = vbNullString: iExtentionIndex = 0
 hasInit = True
End Sub
Function ShowOpen(OwnerForm As Form) As Boolean
 init
 If Not formatFilter Then Exit Function
 If Not formatTitle(False) Then Exit Function
 Dim OFN   As OPENFILENAME
 Dim r     As Long

    With OFN
        .lStructSize = Len(OFN)
        .hwndOwner = OwnerForm.hwnd
        .hInstance = App.hInstance
        .lpstrFilter = iFilter
        .nMaxFile = MAX_PATH * MaxFileCount
        .nFilterIndex = iExtentionIndex
        .lpstrFile = String(.nMaxFile, 0)
        .nMaxFileTitle = MAX_PATH
        .lpstrFileTitle = String(256, 0)
        .lpstrInitialDir = iInitDir
        .lpstrTitle = iTitle
        .flags = OFN_EXPLORER + OFN_HIDEREADONLY + OFN_ALLOWMULTISELECT * Abs(iMultiple)
        .lpstrDefExt = iExtention
    End With
    Dim l As Long: l = GetTickCount
    r = GetOpenFileName(OFN)
    If GetTickCount - l < 20 Then
     OFN.lpstrFile = ""
     r = GetOpenFileName(OFN)
    End If
    iFileName = vbNullString
    If r = 1 Then
     Dim temp As String
     temp = Left$(OFN.lpstrFile, InStr(1, OFN.lpstrFile, String(2, vbNullChar)))
     If Mid$(OFN.lpstrFile, InStr(1, OFN.lpstrFile, vbNullChar) + 1, 1) <> vbNullChar Then
      Dim path As String, p As Long, p2 As Long
      p = InStr(temp, vbNullChar)
      path = addSlash(Left(temp, p - 1))
      p2 = InStr(p + 1, temp, vbNullChar)
      Do While p2 > 0
       If p2 - p < 1 Then Exit Do
       iFileName = iFileName & Chr(34) & path & Mid$(temp, p + 1, p2 - p - 1) & Chr(34)
       p = p2
       p2 = InStr(p + 1, temp, vbNullChar)
      Loop
     Else
      If Len(temp) > 1 Then iFileName = Left(temp, Len(temp) - 1)
     End If
    End If
    iExtentionIndex = OFN.nFilterIndex
    ShowOpen = iFileName <> vbNullString
End Function
Function ShowSave(OwnerForm As Form) As Boolean
 init
 If Not formatFile Then Exit Function
 If Not formatFilter Then Exit Function
 If Not formatTitle(True) Then Exit Function
 Dim OFN   As OPENFILENAME
 Dim r     As Long

    With OFN
        .lStructSize = LenB(OFN)
        .hwndOwner = OwnerForm.hwnd
        .hInstance = App.hInstance
        .lpstrFilter = Replace(Filter, "|", vbNullChar)
        .nFilterIndex = iExtentionIndex
        .lpstrFile = iFileName
        .nMaxFile = MAX_PATH
        .lpstrFileTitle = Space$(MAX_PATH - 1)
        .nMaxFileTitle = MAX_PATH
        .lpstrInitialDir = iInitDir
        .lpstrTitle = iTitle
        .flags = 0
        .lpstrDefExt = iExtention
    End With
    Dim l As Long: l = GetTickCount
    r = GetOpenFileName(OFN)
    If GetTickCount - l < 20 Then
     OFN.lpstrFile = ""
     r = GetOpenFileName(OFN)
    End If
    iExtentionIndex = OFN.nFilterIndex
    If r = 1 Then
     fileName = Left$(OFN.lpstrFile, InStr(1, OFN.lpstrFile & vbNullChar, vbNullChar) - 1)
     ShowSave = True
    Else
     fileName = vbNullString
    End If
End Function

Public Function BrowseFolders(FormObject As Form, sMessage As String) As String
    Dim B As BrowseInfo
    Dim r As Long
    Dim l As Long
    Dim f As String

    FormObject.Enabled = False
    With B
        .hwndOwner = FormObject.hwnd
        .lpszTitle = lstrcat(sMessage, "")
        .ulFlags = BrowseForFolders
    End With

    SHGetSpecialFolderLocation FormObject.hwnd, CSIDL_DRIVES, B.pIDLRoot
    r = SHBrowseForFolder(B)

    If r <> 0 Then
        f = String(MAX_PATH, vbNullChar)
        SHGetPathFromIDList r, f
        CoTaskMemFree r
        l = InStr(1, f, vbNullChar) - 1
        If l < 0 Then l = 0
        f = Left(f, l)
        addSlash f
    End If

    BrowseFolders = f
    FormObject.Enabled = True

End Function
Public Property Get AppPath() As String
    Static m_AppPath As String 'Returns Program EXE File Name
    If Len(m_AppPath) = 0 Then
        Dim ret As Long
        Dim Length As Long
        Dim FilePath As String
        Dim FileHandle As Long
        FilePath = String(MAX_PATH, 0)
        FileHandle = GetModuleHandle(App.EXEName)
        ret = GetModuleFileName(FileHandle, FilePath, MAX_PATH)
        Length = InStr(1, FilePath, vbNullChar) - 1
        If Length > 0 Then m_AppPath = Left$(FilePath, Length)
    End If
    AppPath = m_AppPath
End Property
Private Function formatFile() As Boolean
 On Error GoTo formatError
 If Len(fileName) > MAX_PATH Or Len(InitDir) > MAX_PATH Then GoTo formatError
 fileName = fileName + String(MAX_PATH - Len(fileName), 0)
 formatFile = True
formatError:
End Function
Private Function formatTitle(isSave As Boolean) As Boolean
 If iTitle = vbNullString Then
         If isSave Then iTitle = "Save File" Else iTitle = "Open File"
         formatTitle = True
 Else
         If Len(iTitle) <= MaxTitle Then formatTitle = True
 End If
End Function
Private Function formatFilter() As Boolean
 If Filter = vbNullString Then iFilter = "All files(*.*)|*.*"
 iFilter = Replace(iFilter, "|", vbNullChar) & vbNullChar & vbNullChar
 formatFilter = True
End Function
Private Function addSlash(str As String) As String
 If Right(str, 1) <> "\" Then addSlash = str & "\" Else addSlash = str
End Function
