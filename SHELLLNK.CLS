VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cShellLink"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'---------------------------------------------------------------
'- Public enums...
'---------------------------------------------------------------
Public Enum STGM
    STGM_DIRECT = &H0&
    STGM_TRANSACTED = &H10000
    STGM_SIMPLE = &H8000000
    STGM_READ = &H0&
    STGM_WRITE = &H1&
    STGM_READWRITE = &H2&
    STGM_SHARE_DENY_NONE = &H40&
    STGM_SHARE_DENY_READ = &H30&
    STGM_SHARE_DENY_WRITE = &H20&
    STGM_SHARE_EXCLUSIVE = &H10&
    STGM_PRIORITY = &H40000
    STGM_DELETEONRELEASE = &H4000000
    STGM_CREATE = &H1000&
    STGM_CONVERT = &H20000
    STGM_FAILIFTHERE = &H0&
    STGM_NOSCRATCH = &H100000
End Enum

Public Enum SHELLFOLDERS            ' Shell Folder Path Constants...
    CSIDL_DESKTOP = &H0&            ' ..\WinNT\profiles\username\Desktop
    CSIDL_PROGRAMS = &H2&           ' ..\WinNT\profiles\username\Start Menu\Programs
    CSIDL_CONTROLS = &H3&           ' No Path
    CSIDL_PRINTERS = &H4&           ' No Path
    CSIDL_PERSONAL = &H5&           ' ..\WinNT\profiles\username\Personal
    CSIDL_FAVORITES = &H6&          ' ..\WinNT\profiles\username\Favorites
    CSIDL_STARTUP = &H7&            ' ..\WinNT\profiles\username\Start Menu\Programs\Startup
    CSIDL_RECENT = &H8&             ' ..\WinNT\profiles\username\Recent
    CSIDL_SENDTO = &H9&             ' ..\WinNT\profiles\username\SendTo
    CSIDL_BITBUCKET = &HA&          ' No Path
    CSIDL_STARTMENU = &HB&          ' ..\WinNT\profiles\username\Start Menu
    CSIDL_DESKTOPDIRECTORY = &H10&  ' ..\WinNT\profiles\username\Desktop
    CSIDL_DRIVES = &H11&            ' No Path
    CSIDL_NETWORK = &H12&           ' No Path
    CSIDL_NETHOOD = &H13&           ' ..\WinNT\profiles\username\NetHood
    CSIDL_FONTS = &H14&             ' ..\WinNT\fonts
    CSIDL_TEMPLATES = &H15&         ' ..\WinNT\ShellNew
    CSIDL_COMMON_STARTMENU = &H16&  ' ..\WinNT\profiles\All Users\Start Menu
    CSIDL_COMMON_PROGRAMS = &H17&   ' ..\WinNT\profiles\All Users\Start Menu\Programs
    CSIDL_COMMON_STARTUP = &H18&    ' ..\WinNT\profiles\All Users\Start Menu\Programs\Startup
    CSIDL_COMMON_DESKTOPDIRECTORY = &H19& '..\WinNT\profiles\All Users\Desktop
    CSIDL_APPDATA = &H1A&           ' ..\WinNT\profiles\username\Application Data
    CSIDL_PRINTHOOD = &H1B&         ' ..\WinNT\profiles\username\PrintHood
End Enum

Public Enum SHOWCMDFLAGS
    SHOWNORMAL = 5
    SHOWMAXIMIZE = 3
    SHOWMINIMIZE = 7
End Enum

'---------------------------------------------------------------
Public Function GetSystemFolderPath(ByVal hwnd As Long, ByVal Id As Integer, sfPath As String) As Long
'---------------------------------------------------------------
    Dim rc As Long                                      ' Return code
    Dim pidl As Long                                    ' ptr to Item ID List
    Dim cbPath As Long                                  ' char count of path
    Dim szPath As String                                ' String var for path
'---------------------------------------------------------------
    szPath = Space(MAX_PATH)                            ' Pre-allocate path string for api call

    rc = SHGetSpecialFolderLocation(hwnd, Id, pidl)     ' Get pidl for Id...
    If (rc = 0) Then                                    ' If success is 0
#If UNICODE Then
        rc = SHGetPathFromIDList(pidl, StrPtr(szPath))  ' Get Path from Item Id List
#Else
        rc = SHGetPathFromIDList(pidl, szPath)          ' Get Path from Item Id List
#End If
        If (rc = 1) Then                                ' If success is 1
            szPath = Trim$(szPath)                      ' Fix path string
            cbPath = Len(szPath)                        ' Get length of path
            If (Asc(Right(szPath, 1)) = 0) Then cbPath = cbPath - 1 ' Adjust path length
            If (cbPath > 0) Then sfPath = Left$(szPath, cbPath) ' Adjust path string variable
            GetSystemFolderPath = True                  ' Return success
        End If
    End If
'---------------------------------------------------------------
End Function
'---------------------------------------------------------------

'---------------------------------------------------------------
Public Function CreateShellLink(LnkFile As String, ExeFile As String, WorkDir As String, _
                                ExeArgs As String, IconFile As String, IconIdx As Long, _
                                ShowCmd As SHOWCMDFLAGS) As Long
'---------------------------------------------------------------
    Dim rc As Long
    Dim pidl As Long                                    ' Item id list
    Dim dwReserved As Long                              ' Reserved flag
    Dim cShellLink As ShellLinkA                        ' An explorer IShellLinkA(Win 95/Win NT) instance
    Dim cPersistFile As IPersistFile                    ' An explorer IPersistFile instance
'---------------------------------------------------------------
    If ((LnkFile = "") Or (ExeFile = "")) Then Exit Function    ' Validate min. input requirements.
    
    On Error GoTo ErrHandler
    Set cShellLink = New ShellLinkA                     ' Create new IShellLink interface
    Set cPersistFile = cShellLink                       ' Implement cShellLink's IPersistFile interface
    
    With cShellLink
        .SetPath ExeFile                                ' set command line exe name & path to new ShortCut.
        
        If (WorkDir <> "") Then .SetWorkingDirectory WorkDir ' Set working directory in shortcut
        
        If (ExeArgs <> "") Then .SetArguments ExeArgs   ' Add arguments to command line
        
'       if (LnkDesc <> "") then .SetDescription pszName ' Set shortcut description
'       .SetHotkey wHotKey
       
        If (IconFile <> "") Then .SetIconLocation IconFile, IconIdx ' Set shortcut icon location & index
        
        .SetDescription "ShellLink Sample" & vbNullChar
'       .SetIDList pidl
'       dwReserved = 0
'       .SetRelativePath pszPathRel, dwReserved

        .SetShowCmd ShowCmd                             ' Set shortcut's startup mode (min,max,normal)
    End With
    
    cShellLink.Resolve 0, SLR_UPDATE
    cPersistFile.Save StrConv(LnkFile, vbUnicode), 0    ' Unicode conversion hack... This must be done!
    CreateShellLink = True                              ' Return Success

'---------------------------------------------------------------
ErrHandler:
'---------------------------------------------------------------
    Set cPersistFile = Nothing                          ' Destroy Object
    Set cShellLink = Nothing                            ' Destroy Object
'---------------------------------------------------------------
End Function
'---------------------------------------------------------------

'---------------------------------------------------------------
Public Function GetShellLinkInfo(LnkFile As String, ExeFile As String, WorkDir As String, _
                                 ExeArgs As String, IconFile As String, IconIdx As Long, _
                                 ShowCmd As Long) As Long
'---------------------------------------------------------------
    Dim pidl As Long                                    ' Item id list
    Dim wHotKey As Long                                 ' Hotkey to shortcut...
    Dim fd As WIN32_FIND_DATA
    Dim Description As String
    Dim buffLen As Long
    Dim cShellLink As ShellLinkA                        ' An explorer IShellLink instance
    Dim cPersistFile As IPersistFile                    ' An explorer IPersistFile instance
'---------------------------------------------------------------
    If (LnkFile = "") Then Exit Function                ' Validate min. input requirements.
    
    Set cShellLink = New ShellLinkA                     ' Create new IShellLink interface
    Set cPersistFile = cShellLink                       ' Implement cShellLink's IPersistFile interface
    
    ' Load Shortcut file...(must do this UNICODE hack!)
    On Error GoTo ErrHandler
    cPersistFile.Load StrConv(LnkFile, vbUnicode), STGM_DIRECT
    
    With cShellLink
        ' Get command line exe name & path of shortcut
        ExeFile = Space(MAX_PATH)
        buffLen = Len(ExeFile)
        .GetPath ExeFile, buffLen, fd, SLGP_UNCPRIORITY
        Dim s As String
        s = fd.cFileName                                ' Not returned to calling function
        
        ' Get working directory of shortcut
        WorkDir = Space(MAX_PATH)
        buffLen = Len(WorkDir)
        .GetWorkingDirectory WorkDir, buffLen
        
        ' Get command line arguments of shortcut
        ExeArgs = Space(MAX_PATH)
        buffLen = Len(ExeArgs)
        .GetArguments ExeArgs, buffLen
        
        ' Get description of shortcut
        Description = Space(MAX_PATH)
        buffLen = Len(Description)
        .GetDescription Description, buffLen            ' Not returned to calling function
        
        ' Get the HotKey for shortcut
        .GetHotkey wHotKey                              ' Not returned to calling function
       
        ' Get shortcut icon location & index
        IconFile = Space(MAX_PATH)
        buffLen = Len(IconFile)
        .GetIconLocation IconFile, buffLen, IconIdx
        
        ' Get Item ID List...
        .GetIDList pidl                                 ' Not returned to calling function
                
        ' Set shortcut's startup mode (min,max,normal)
        .GetShowCmd ShowCmd
    End With

    GetShellLinkInfo = True                             ' Return Success
'---------------------------------------------------------------
ErrHandler:
'---------------------------------------------------------------
    Set cPersistFile = Nothing                          ' Destroy Object
    Set cShellLink = Nothing                            ' Destroy Object
'---------------------------------------------------------------
End Function
'---------------------------------------------------------------
