Attribute VB_Name = "modFileOperation"
Option Explicit

Public blnCopyFile As Boolean 'copy the file to target path?
Public blnDeleteAfterCopy As Boolean 'delete original after (successful) copy (=Move)
Public iWhatIfFileNewer As Integer 'holds what to do when target file is newer
' 0 = prompt user
' 1 = overwrite target file
' 2 = don't copy source file
Public iWhatIfAttrDifferent As Integer ''holds what to do when file is same but attributes differ
' 0 = prompt user
' 1 = overwrite target file
' 2 = don't copy source file

Private Const MSG_TITLE_COPY As String = "Copy File"
Private Const MAXDWORD As Long = &HFFFFFFFF
Private Const MAX_PATH As Long = 260
Private Const INVALID_HANDLE_VALUE As Long = -1

Private Const FILE_ATTRIBUTE_HIDDEN = &H2
Private Const FILE_ATTRIBUTE_SYSTEM = &H4
Private Const FILE_ATTRIBUTE_DIRECTORY As Long = &H10
Private Const FILE_ATTRIBUTE_NORMAL = &H80
Private Const FILE_ATTRIBUTE_TEMPORARY = &H100

Private Const ERROR_NO_MORE_FILES = 18&

Private Const FORMAT_MESSAGE_FROM_SYSTEM As Long = &H1000
Private Const FORMAT_MESSAGE_IGNORE_INSERTS As Long = &H200
Private Const FORMAT_MESSAGE_MAX_WIDTH_MASK As Long = &HFF

Public Enum FO_FLAGS
    FOF_MULTIDESTFILES = &H1
    FOF_CONFIRMMOUSE = &H2
    FOF_SILENT = &H4
    FOF_RENAMEONCOLLISION = &H8
    FOF_NOCONFIRMATION = &H10
    FOF_WANTMAPPINGHANDLE = &H20
    FOF_ALLOWUNDO = &H40
    FOF_FILESONLY = &H80
    FOF_SIMPLEPROGRESS = &H100
    FOF_NOCONFIRMMKDIR = &H200
    FOF_DeleteFlags = &H154
    FOF_CopyFlags = &H3DD
End Enum

Public Enum FO_FUNC
    FO_MOVE = 1
    FO_COPY = 2
    FO_DELETE = 3
    FO_RENAME = 4
End Enum

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

Private Type SECURITY_ATTRIBUTES
   nLength As Long
   lpSecurityDescriptor As Long
   bInheritHandle As Long
End Type

Private Type SHFILEOPSTRUCT
   hwnd        As Long
   wFunc       As Long
   pFrom       As String
   pTo         As String
   fFlags      As Integer
   fAborted    As Boolean
   hNameMaps   As Long
   sProgress   As String
End Type

Private Declare Function SHFileOperation Lib "shell32" _
   Alias "SHFileOperationA" _
  (lpFileOp As SHFILEOPSTRUCT) As Long
  
Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" _
    (ByVal dwFlags As Long, _
    lpSource As Any, _
    ByVal dwMessageId As Long, _
    ByVal dwLanguageId As Long, _
    ByVal lpBuffer As String, _
    ByVal nSize As Long, _
    Arguments As Long) As Long
    
Private Declare Function FindFirstFile Lib "kernel32" _
   Alias "FindFirstFileA" _
  (ByVal lpFileName As String, _
   lpFindFileData As WIN32_FIND_DATA) As Long
   
Private Declare Function FindNextFile Lib "kernel32" _
   Alias "FindNextFileA" _
  (ByVal hFindFile As Long, _
   lpFindFileData As WIN32_FIND_DATA) As Long
   
Private Declare Function FindClose Lib "kernel32" _
  (ByVal hFindFile As Long) As Long
      
Private Declare Function CompareFileTime Lib "kernel32" _
  (lpFileTime1 As FILETIME, _
   lpFileTime2 As FILETIME) As Long

'Private Declare Function GetFullPathName Lib "kernel32" _
'    Alias "GetFullPathNameA" _
'    (ByVal lpFileName As String, _
'    ByVal nBufferLength As Long, _
'    ByVal lpBuffer As String, _
'    ByVal lpFilePart As String) As Long

Private Declare Function CopyFile Lib "kernel32" _
   Alias "CopyFileA" _
  (ByVal lpExistingFileName As String, _
   ByVal lpNewFileName As String, _
   ByVal bFailIfExists As Long) As Long
   
'Private Declare Function CreateDirectory Lib "kernel32" _
'    Alias "CreateDirectoryA" _
'   (ByVal lpPathName As String, _
'    lpSecurityAttributes As SECURITY_ATTRIBUTES) As Long

'file title
Private Declare Function GetFileTitle Lib "comdlg32.dll" _
    Alias "GetFileTitleA" ( _
    ByVal lpszFile As String, ByVal lpszTitle As String, _
    ByVal cbBuf As Integer) As Integer
    
Private Declare Function MakeSureDirectoryPathExists Lib "imagehlp.dll" _
    (ByVal DirPath As String) As Long

' ---- END DECLARES ----

Private Function FileCompareFileDates(WFDSource As WIN32_FIND_DATA, _
                                      WFDTarget As WIN32_FIND_DATA) As Long
   
   Dim CTSource As FILETIME
   Dim CTTarget As FILETIME
   
  'assign the source and target file write
  'times to a FILETIME structure, and compare.
   CTSource.dwHighDateTime = WFDSource.ftLastWriteTime.dwHighDateTime
   CTSource.dwLowDateTime = WFDSource.ftLastWriteTime.dwLowDateTime
   
   CTTarget.dwHighDateTime = WFDTarget.ftLastWriteTime.dwHighDateTime
   CTTarget.dwLowDateTime = WFDTarget.ftLastWriteTime.dwLowDateTime
   
   FileCompareFileDates = CompareFileTime(CTSource, CTTarget)
   
End Function

Public Function CopySourceFile(ByVal sSourceFile As String, _
                                    ByVal sSourceFolder As String, _
                                    ByVal sTargetFolder As String) As Long

  'common local working variables
   Dim WFDSource As WIN32_FIND_DATA
   Dim hFileSource As Long
   Dim sTmp As String
   Dim sTargetMsg As String
   Dim sSourceMsg As String
   Dim diff As Long
   Dim reply As Integer
   
  'variables used for the source files and folders
   Dim dwSourceFileSize As Long

  'variables used for the target files and folders
   Dim WFDTarget As WIN32_FIND_DATA
   Dim hTargetFile As Long
   Dim dwTargetFileSize As Long

    hFileSource = FileGetFileHandle(sSourceFolder, WFDSource, sSourceFile)
   
  'last check!
   If hFileSource <> INVALID_HANDLE_VALUE Then

    'remove trailing nulls from the first retrieved object
     sTmp = TrimNull(WFDSource.cFileName)
     
    'if the object is not a folder..
     If (WFDSource.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) <> FILE_ATTRIBUTE_DIRECTORY Then
     
       'check for the corresponding file
       'in the target folder by using the API
       'to locate that specific file
        hTargetFile = FindFirstFile(sTargetFolder & sTmp, WFDTarget)
       
       'if the file is located in the target folder..
        If hTargetFile <> INVALID_HANDLE_VALUE Then
        
          'get the file size for the source and target files
           dwSourceFileSize = FileGetFileSize(WFDSource)
           dwTargetFileSize = FileGetFileSize(WFDTarget)

          'compare the dates.
          'If diff = 0 source and target are the same
          'If diff = 1 source is newer than target
          'If diff = -1 source is older than target
           diff = FileCompareFileDates(WFDSource, WFDTarget)
           
          'if the dates, attributes and file times
          'are the same...
           If (dwSourceFileSize = dwTargetFileSize) And _
              WFDSource.dwFileAttributes = WFDTarget.dwFileAttributes And _
              diff = 0 Then
                '...the files are the same...
                'use target file name as link
                sLinkFile = sTargetFolder & sTmp
           Else
           
             'files are not the same

              If diff = 1 Then
                'perform the preferred copy method ONLY if
                'diff indicated that the source was newer!
                Call APICopyFile(sSourceFolder & sTmp, sTargetFolder & sTmp, False)
                
              ElseIf diff = -1 Then 'source is older
              
                Select Case iWhatIfFileNewer
                    Case 0 'prompt user
                        reply = MsgBox("Target file is newer than the file you want to link. Overwrite it?" _
                                    & vbCrLf & vbCrLf & "Answer No to link the original file, Cancel to abort.", vbQuestion + vbYesNoCancel, MSG_TITLE_COPY)
                        If reply = vbYes Then
                            Call APICopyFile(sSourceFolder & sTmp, sTargetFolder & sTmp, False)
                        ElseIf reply = vbCancel Then
                            blnAbortLinking = True
                            Exit Function
                        Else
                            'user chose not to overwrite, so return old file link
                            sLinkFile = sSourceFolder & sSourceFile
                        End If
                    Case 1 'overwrite target
                        Call APICopyFile(sSourceFolder & sTmp, sTargetFolder & sTmp, False)
                    Case 2 'do not overwrite target
                        sLinkFile = sSourceFolder & sSourceFile
                End Select

              ElseIf diff = 0 Then
                'the dates are the same but the file attributes
                'are different.
                Select Case iWhatIfFileNewer
                    Case 0 'prompt user
                        reply = MsgBox("Target file already exists and has same date and size as source file," _
                                    & vbCrLf & "but the file attributes are different. Overwrite it?" _
                                    & vbCrLf & vbCrLf & "Answer No to link the original file, Cancel to abort.", vbQuestion + vbYesNoCancel, MSG_TITLE_COPY)
                        If reply = vbYes Then
                            Call APICopyFile(sSourceFolder & sTmp, sTargetFolder & sTmp, False)
                        ElseIf reply = vbCancel Then
                            blnAbortLinking = True
                            Exit Function
                        Else
                            sLinkFile = sSourceFolder & sTmp
                        End If
                    Case 1 'overwrite target
                        Call APICopyFile(sSourceFolder & sTmp, sTargetFolder & sTmp, False)
                    Case 2 'do not overwrite target
                        sLinkFile = sSourceFolder & sSourceFile
                End Select
                
              End If
              
           End If  'If dwSourceFileSize
           
          'since the target file was found,
          'close the handle
           Call FindClose(hTargetFile)
           
        Else:
        
             'the target file was not found so
             'copy the file to the target directory
              Call APICopyFile(sSourceFolder & sTmp, sTargetFolder & sTmp, False)
          
        End If  'If hTargetFile
     End If  'If WFDSource.dwFileAttributes

    'clear the local variables
     dwSourceFileSize = 0
     dwTargetFileSize = 0

   End If
   
End Function


Private Function FileGetFileSize(WFD As WIN32_FIND_DATA) As Long

   FileGetFileSize = (WFD.nFileSizeHigh * (MAXDWORD + 1)) + WFD.nFileSizeLow
   
End Function


Private Function FileGetFileHandle(sPathToFiles As String, _
                                   WFD As WIN32_FIND_DATA, _
                                   Optional FileTitle As String) As Long

   Dim sPath As String
   Dim sRoot As String
      
   sRoot = QualifyPath(sPathToFiles)
    If IsMissing(FileTitle) Then
        sPath = sRoot & "*.*"
    Else
        sPath = sRoot & FileTitle
    End If
    
  'obtain handle to the first match
  'in the target folder
   FileGetFileHandle = FindFirstFile(sPath, WFD)
   
End Function

Function FileOrDirExists(Optional ByVal sFile As String = "", _
        Optional ByVal sFolder As String = "") As Boolean

    Dim lpFindFileData As WIN32_FIND_DATA, lFileHdl  As Long
    Dim sTemp As String, sTemp2 As String, lRet As Long, iLastIndex  As Integer
    Dim sPath As String
    Dim sStartDir As String
    
    On Error Resume Next
    '// both params are empty
    If sFile = "" And sFolder = "" Then Exit Function
    '// both are full, empty folder param
    If sFile <> "" And sFolder <> "" Then sFolder = ""
    If sFolder <> "" Then
        '// set start directory
        sStartDir = sFolder
    Else
        '// extract start directory from file path
        sStartDir = Left$(sFile, InStrRev(sFile, "\"))
        '// just get filename
        sFile = Right$(sFile, Len(sFile) - InStrRev(sFile, "\"))
    End If
    '// add trailing to start directory if required
    If Right$(sStartDir, 1) <> "" Then sStartDir = sStartDir & ""
    
    sStartDir = sStartDir & "*.*"
    
    '// get a file handle
    lFileHdl = FindFirstFile(sStartDir, lpFindFileData)
    
    If lFileHdl <> -1 Then
        If sFolder <> "" Then
            '// folder exists
            FileOrDirExists = True
        Else
            Do Until lRet = ERROR_NO_MORE_FILES
                sPath = Left$(sStartDir, Len(sStartDir) - 4) & ""
                '// if it is a file
                If (lpFindFileData.dwFileAttributes And FILE_ATTRIBUTE_NORMAL) = vbNormal Then
                    sTemp = StrConv(StripTerminator(lpFindFileData.cFileName), vbProperCase)
                    '// remove LCase$ if you want the search to be case sensitive (unlikely!)
                    If LCase$(sTemp) = LCase$(sFile) Then
                        FileOrDirExists = True '// file found
                        Exit Do
                    End If
                End If
                '// based on the file handle iterate through all files and dirs
                lRet = FindNextFile(lFileHdl, lpFindFileData)
                If lRet = 0 Then Exit Do
            Loop
        End If
    End If
    '// close the file handle
    lRet = FindClose(lFileHdl)
End Function

Public Function APIGetFileTitle(sFile As String) As String
    Dim sFileTitle As String, cFileTitle As Integer

    cFileTitle = MAX_PATH
    sFileTitle = String$(MAX_PATH, 0)
    cFileTitle = GetFileTitle(sFile, sFileTitle, MAX_PATH)
    If cFileTitle Then
        APIGetFileTitle = ""
    Else
        APIGetFileTitle = Left$(sFileTitle, InStr(sFileTitle, vbNullChar) - 1)
    End If

End Function

Public Function UnQualifyPath(ByVal sFolder As String) As String

  'remove any trailing slash
   sFolder = Trim(sFolder)
   
   If Right(sFolder, 1) = "\" Then
      UnQualifyPath = Left(sFolder, Len(sFolder) - 1)
   Else
      UnQualifyPath = sFolder
   End If
   
End Function

Public Function QualifyPath(sPath As String) As String

  'assures that a passed path ends in a slash
   If Right(sPath, 1) <> "\" Then
      QualifyPath = sPath & "\"
   Else
      QualifyPath = sPath
   End If
      
End Function


Private Function TrimNull(startstr As String) As String

  'returns the string up to the first
  'null, if present, or the passed string
   Dim pos As Integer
   
   pos = InStr(startstr, Chr$(0))
   
   If pos Then
      TrimNull = Left$(startstr, pos - 1)
      Exit Function
   End If
  
   TrimNull = startstr
  
End Function

Private Function StripTerminator(ByVal strString As String) As String
    Dim intZeroPos As Integer
    intZeroPos = InStr(strString, Chr$(0))
    If intZeroPos > 0 Then
        StripTerminator = Left$(strString, intZeroPos - 1)
    Else
        StripTerminator = strString
    End If
End Function

Public Function APICreatePath(NewPath As String) As Boolean
    Dim errMsg As String
    Dim errCode As Long
    
    'Add a trailing slash if none
    NewPath = QualifyPath(NewPath)
    
    'Call API
    If MakeSureDirectoryPathExists(NewPath) <> 0 Then
        'No errors, return True
        APICreatePath = True
    Else 'log error message
        errCode = Err.LastDllError
        errMsg = APIErrorMessage(errCode)
        sLogMsg = "Error creating directory '" & NewPath & "'." _
                & "Error: " & errCode & " " & errMsg
        Call AppendToLog(sLogFile, sLogMsg)
    End If

End Function

' FormatMessage API wrapper function: Returns system error description.
'
Private Function APIErrorMessage(ByVal errCode As Long) As String

    Dim MsgBuffer As String * 257
    Dim MsgLength As Long
  
    MsgLength = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM Or _
        FORMAT_MESSAGE_IGNORE_INSERTS Or FORMAT_MESSAGE_MAX_WIDTH_MASK, 0&, _
        errCode, 0&, MsgBuffer, 256&, 0&)
    If (MsgLength = 0) Then
        APIErrorMessage = "Unknown error."
    Else
        APIErrorMessage = Left$(MsgBuffer, MsgLength)
    End If

End Function

Public Function IsInvalidPathSyntax(ByVal s As String) As Boolean
    'checks to see if the pathname contains invalid characters
    Dim BadChars() As Variant
    Dim buffer As String
    Dim pos As Integer
    Dim i As Integer
    BadChars = Array("\\", ":", """", "*", "?", "<", ">", "|")
    
    pos = InStr(s, ":\")
    If pos Then 'drive letter
        buffer = Right(s, Len(s) - pos - 1)
    ElseIf Left(s, 2) = "\\" Then 'UNC path
        buffer = Right(s, Len(s) - 2)
    Else
        buffer = s
    End If

    For i = 0 To UBound(BadChars)
        If InStr(buffer, BadChars(i)) > 0 Then
            IsInvalidPathSyntax = True
            Exit For
        End If
    Next i
    
End Function

Private Sub APICopyFile(ByVal sFrom As String, ByVal sTo As String, bFailIfExists As Boolean)
    'wrapper function around API call
    Dim retval As Long
    Dim errCode As Long
    
    'call the API
    retval = CopyFile(sFrom, sTo, bFailIfExists) 'returns non-zero if successful
    
    If retval <> 0 Then
        sLinkFile = sTo
        sLogMsg = "Successfully copied file '" & sFrom & "' to '" & sLinkFile & "'."
        blnCopySucceeded = True
    Else
        sLinkFile = sFrom
        errCode = Err.LastDllError
        sLogMsg = APIErrorMessage(errCode)
        sLogMsg = "Error occurred while copying file: '" & sFrom & "' to '" & sTo & "'. Error " & errCode & ": " & sLogMsg
    End If
    
    Call AppendToLog(sLogFile, sLogMsg)
    
End Sub

Public Sub APIDeleteFile(sFile As String, FOF_FLAGS As Long)
 
    'set some working variables
    Dim shf As SHFILEOPSTRUCT
    Dim retval As Long
    Dim errCode As Long
    
  'terminate the file string with
  'a pair of null chars
   sFile = sFile & vbNullChar '& vbNullChar
 
    'set up the SHFile options
    With shf
       .wFunc = FO_DELETE  'action to perform
       .pFrom = sFile      'the file to act on
       .fFlags = FOF_FLAGS 'special flags
    End With
 
    'perform the delete
    retval = SHFileOperation(shf) 'returns zero -f successful
    If retval = 0 Then
        sLogMsg = "Successfully deleted file '" & sFile & "'."
    Else
        errCode = Err.LastDllError
        sLogMsg = APIErrorMessage(errCode)
        sLogMsg = "Error occurred while deleting file: '" & sFile & "'. Error " & errCode & ": " & sLogMsg
    End If
    
    Call AppendToLog(sLogFile, sLogMsg)
    
End Sub

Function GetLogText() As String

    Dim filenumber As Integer
    
    filenumber = FreeFile           'Freefile returns the first unused file number
    Open sLogFile For Input Access Read As filenumber    'Opens the file for input
    GetLogText = Input(LOF(filenumber), filenumber)
    Close filenumber            'Close when done
    
End Function

Sub PurgeLog()

    Dim filenumber As Integer
    
    filenumber = FreeFile           'Freefile returns the first unused file number
    Open sLogFile For Output As filenumber    'Opens the file for input
    Print #filenumber, CStr(Now) & ": Purged logfile"
    Close filenumber            'Close when done
    MsgBox "Log file was purged.", vbInformation, app.EXEName
    
End Sub
