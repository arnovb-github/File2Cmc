Attribute VB_Name = "modMain"
Option Explicit

Public Const COPYRIGHT As String = " © 2004 - 2007 Arno van Boven"
Public Const SWP_NOMOVE = 2
Public Const SWP_NOSIZE = 1
Public Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const SW_SHOWNORMAL As Long = 1  'ShowWindow mode

'section Commence and keys
Public Const INI_SECTION_CMC As String = "Commence"
Public Const INI_KEY_CMC_CAT As String = "Category"
Public Const INI_KEY_CMC_NAMEFIELD As String = "NameField"
Public Const INI_KEY_CMC_DATAFILE As String = "DataFile"
Public Const INI_KEY_CMC_IDF As String = "DetailForm"

'section OtherOptions and keys
Public Const INI_SECTION_OTHER As String = "OtherOptions"
Public Const INI_KEY_OTHER_SHARE As String = "ShareDefault"
Public Const INI_KEY_OTHER_SHOWIDF As String = "ShowForm"
Public Const INI_KEY_OTHER_SHOWCMC As String = "ShowCommence"
Public Const INI_KEY_OTHER_AUTOCON As String = "AutoConnect"
Public Const INI_KEY_OTHER_DDE As String = "DisableDDETRigger"

'section Display and keys
Public Const INI_SECTION_DISPLAY As String = "Display"
Public Const INI_KEY_DISPLAY_CLOSE As String = "CloseWindow"
Public Const INI_KEY_DISPLAY_ONTOP As String = "AlwaysOnTop"

'section File Options and keys
Public Const INI_SECTION_FILE As String = "FileOptions"
Public Const INI_KEY_FILE_COPY As String = "CopyFile"
Public Const INI_KEY_FILE_DELETEORIGINAL As String = "DeleteAfterCopy"
Public Const INI_KEY_FILE_CAT As String = "PathCategory"
Public Const INI_KEY_FILE_FIELD As String = "PathField"
Public Const INI_KEY_FILE_CUSTOM As String = "PathCustom"
Public Const INI_KEY_FILE_ORDER As String = "PathPrecedence"
'Public Const INI_KEY_FILE_SHCUT As String = "CreateShortcut"
Public Const INI_KEY_FILE_NEWER As String = "HandleNewer"
Public Const INI_KEY_FILE_ATTR As String = "HandleAttrDiff"

'variables
Public sIniFile As String 'ini file containing settings
Public sLogCat As String 'category to log items in
Public sNameField As String 'namefield of logcat
Public sDataFileField As String 'datafile field of logcat
Public sDetailForm As String 'form to show after a link
Public sCmcCategory As String 'Cmc category info read from ini
Public sFields() As String 'array of fieldnames used to better control what is passed to add item routine
Public sFieldValues() As String 'array or fieldname values used to better control what is passed to add item routine
Public sFileNames() As String 'array to hold filenames
Public bNoAutoConnect As Boolean 'automatically create connections or not
Public sLinkFile As String 'file to be linked; might be a shortcut
Public sCopyPath As String 'string holding path to copy to
Public sCopyCategory As String 'category containing pathname to copy to
Public sCopyFieldName As String 'Fieldname containing pathname to copy to
Public iPrecedence As Integer 'integer specifying what to do when copying a file
'0 = use path from database
'1 = use path from ini file (=userdefined)
'2 = try database path first, then fromn ini file
'if no path specified, do not copy the file
Public sTextFields() As String  'fields that can hold string values, used for Copy field names aray
Public blnAbortLinking As Boolean 'abort linking files, continue as normal
Public blnCopySucceeded As Boolean 'only true if APICopyFile succeeded

' put focus on Commence
Public Declare Function FindWindow Lib "user32" _
    Alias "FindWindowA" ( _
    ByVal lpClassName As String, ByVal lpWindowName As String) As Long

Public Declare Function ShowWindow Lib "user32" ( _
    ByVal hwnd As Long, ByVal nCmdShow As Long) As Long

Public Declare Function SetForegroundWindow Lib "user32" ( _
    ByVal hwnd As Long) As Long
    
Private Declare Function SetWindowPos Lib "user32" _
      (ByVal hwnd As Long, _
      ByVal hWndInsertAfter As Long, _
      ByVal X As Long, _
      ByVal Y As Long, _
      ByVal cx As Long, _
      ByVal cy As Long, _
      ByVal wFlags As Long) As Long

' read ini
Private Declare Function GetPrivateProfilestring Lib "kernel32" _
        Alias "GetPrivateProfileStringA" ( _
        ByVal lpApplicationName As String, ByVal lpKeyName As Any, _
        ByVal lpDefault As String, ByVal lpReturnedString As String, _
        ByVal nSize As Long, ByVal lpFileName As String) As Long

'write ini
Private Declare Function WritePrivateProfileString Lib "kernel32" _
        Alias "WritePrivateProfileStringA" ( _
        ByVal lpApplicationName As String, ByVal lpKeyName As String, _
        ByVal lpString As String, ByVal lpFileName As String) As Long
        
' get long name
Private Declare Function GetLongPathName Lib "kernel32" _
    Alias "GetLongPathNameA" ( _
    ByVal lpszShortPath As String, ByVal lpszLongPath As String, _
    ByVal cchBuffer As Long) As Long
    
Private Declare Function ShellExecute Lib "shell32.dll" _
   Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation _
   As String, ByVal lpFile As String, ByVal lpParameters As String, _
   ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Const INI_EXT As String = ".ini"
Private Const MSG_TITLE_COPY As String = "Copy File"

Sub Cleanup()

    On Error Resume Next
    Set oConv = Nothing
    Set oCmc = Nothing

End Sub

Function Init() As Boolean

    Const INIT_ERROR As String = "Initialization error"
    Dim sTitle As String
    Dim bCategory As Byte
    Dim iReply As Integer
       
    sLogFile = app.Path & "\" & app.EXEName & ".log"
    sLogMsg = app.EXEName & " started. Initializing..."
    Call AppendToLog(sLogFile, sLogMsg)
    'check commence meermalen open
    Dim prc As New clsProcesses
    Select Case prc.Count(CMC_NAME_MODULE)
    
    Case Is > 1
        sLogMsg = "Failed to initialize: more than 1 instance of Commence is active."
        Call AppendToLog(sLogFile, sLogMsg)
        MsgBox "There are multiple instances of Commence are running" _
            & vbCrLf & "Please close one or more before you continue", _
            vbCritical, INIT_ERROR
        GoTo err_handler
    Case 0
        sLogMsg = "Failed to initialize: Commence is not running."
        Call AppendToLog(sLogFile, sLogMsg)
        MsgBox "Please start Commence", vbExclamation, INIT_ERROR
        GoTo err_handler
        
    End Select
      
    'initialiseer Commence
    If Not InstantiateCmc Then
        sLogMsg = "Failed to initialize: Unable to get or create ICommence interface. Make sure Commence is installed properly."
        Call AppendToLog(sLogFile, sLogMsg)
        MsgBox "Application Commence is not available", vbCritical, CMC_ERR_API
        GoTo err_handler
    End If
    
    sTitle = app.Title
    sIniFile = oCmc.Path & "\" & sTitle & INI_EXT
   
    'check if we got an ini file to read settings from
    If Not FileOrDirExists(sIniFile, "") Then
        iReply = MsgBox(sIniFile & " not found in " & oCmc.Path _
            & vbCrLf & "Do you want to configure " & sTitle & " now?", vbQuestion + vbYesNo, app.EXEName)
        If iReply = vbYes Then
            frmOptions.Show vbModal, frmMain
        Else
            sLogMsg = "Failed to initialize: creation of options file '" & sIniFile & "' was aborted."
            Call AppendToLog(sLogFile, sLogMsg)
            GoTo err_handler
        End If
        
        If Not FileOrDirExists(sIniFile, "") Then
            sLogMsg = "Failed to initialize: options file '" & sIniFile & "' was not found."
            Call AppendToLog(sLogFile, sLogMsg)
            GoTo err_handler
        End If
    End If
    
    Init = True
    sLogMsg = "Initialized: success!"
    Call AppendToLog(sLogFile, sLogMsg)
    Exit Function
    
err_handler:
    On Error Resume Next
    sLogMsg = "Failed to initialize " & app.EXEName
    Call AppendToLog(sLogFile, sLogMsg)
    Set prc = Nothing
    Init = False
    
End Function

Function sReadIni(ByVal sIni As String, ByVal sSection As String, ByVal sKey As String) As String
    Dim sAppName As String
    Dim sKeyName As String
    Dim sReturn As String
    Dim sFilename As String
    Dim lValid As Long

   '* variabelen voor het aanduiden van onderdelen van het ini-bestand
   sAppName = sSection
   sKeyName = sKey
   sReturn = Space$(255)

   '* This is het pad naar de naam van het ini-bestand bestand
   sFilename = sIni

   'Naam lezen uit het ini-bestand
   lValid = GetPrivateProfilestring(sAppName, sKeyName, "", sReturn, 255, sFilename)

   '* Discard the trailing spaces and null character.
   sReadIni = Left$(sReturn, lValid)
   
End Function

Function sWriteIni(ByVal sSection As String, ByVal sKey As String, ByVal sValue, ByVal sIni As String) As Long
    Dim retval As Long

    retval = WritePrivateProfileString(sSection, sKey, sValue, sIni)
    sWriteIni = retval

End Function

Function ValidateFile(s As String) As Boolean

    'if argument was passed by Windows SendTo,
    'it will be in 8.3 format. Dont want that
    'also make sure that we are dealing with a valid filename
    If FileOrDirExists(s, "") Then
        s = GetLongFilename(s)
    Else
        MsgBox s & "does not exist or is not a valid filename", _
            vbCritical, app.EXEName
        GoTo err_handler
    End If
    
    If Len(s) > CMC_MAX_DATAFILE_LENGTH Then
        MsgBox "Maximum field length exceeded, cannot add file to Commence.", vbCritical, app.EXEName
        GoTo err_handler
    End If
    
err_handler:

End Function

Public Function GetLongFilename(ByVal sShortFilename As String) As String
    'Returns the Long Filename associated wi
    '     th sShortFilename
    Dim lRet As Long
    Dim sLongFilename As String
    'First attempt using 1024 character buff
    '     er.
    sLongFilename = String$(1024, " ")
    lRet = GetLongPathName(sShortFilename, sLongFilename, Len(sLongFilename))
    
    'If buffer is too small lRet contains buffer size needed.

    If lRet > Len(sLongFilename) Then
        'Increase buffer size...
        sLongFilename = String$(lRet + 1, " ")
        'and try again.
        lRet = GetLongPathName(sShortFilename, sLongFilename, Len(sLongFilename))
    End If
    
    'lRet contains the number of characters returned.

    If lRet > 0 Then
        GetLongFilename = Left$(sLongFilename, lRet)
    End If
    
End Function

Public Function IsDimensioned(pArray() As String) As Boolean
    
    On Error GoTo err_handler
    Dim Temp As Integer
    Temp = UBound(pArray)
    IsDimensioned = True
    Exit Function
err_handler:
    IsDimensioned = False
    
End Function

Public Function dq(ByVal s As String) As String

    dq = Chr(34) & s & Chr(34)
    
End Function

Public Function ParseOutSpaces(ByVal s As String)
    'strips all spaces until only single spaces are left
    Dim sStrip As String
    
    sStrip = Space(2)
    Do While InStr(s, sStrip) > 0
        s = Replace(s, sStrip, Space(1))
    Loop
    
    ParseOutSpaces = s

End Function

Public Function SetTopMostWindow(hwnd As Long, Topmost As Boolean) As Long

   If Topmost = True Then 'Make the window topmost
      SetTopMostWindow = SetWindowPos(hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
   Else
      SetTopMostWindow = SetWindowPos(hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS)
      SetTopMostWindow = False
   End If
End Function

Public Function CreateLinkFile(ByVal sSourceFile As String) As String
    
    Dim sPath As String
    Dim sSourcePath As String
    Dim sTargetPath As String
    Dim reply As Integer
    Dim sTmp As String
    
    'default return value
    CreateLinkFile = sSourceFile
    'store file title in sTmp
    sTmp = APIGetFileTitle(sSourceFile)
    'store source path
    sSourcePath = GetPathFromFileName(sSourceFile)
    'default path
    sPath = ""
    
    'process the file according to what was set in settings file
    Select Case iPrecedence
        
        Case 0 'use field value from database only

            Call GetPathFromCommence(sPath, True)
            
            If Not sPath = "" Then
                'check for valid syntax
                If IsInvalidPathSyntax(sPath) Then
                    reply = MsgBox("Syntax error in pathname '" & sPath & "'" _
                                 & vbCrLf & vbCrLf & app.EXEName & " will not copy the file link but to the original instead.", vbExclamation + vbOKCancel, MSG_TITLE_COPY)
                    If Not reply = vbOK Then blnAbortLinking = True
                    Exit Function
                End If
                
                sTargetPath = IIf(FileOrDirExists("", sPath), QualifyPath(sPath), CreatePath(sPath))
                Call CopySourceFile(sTmp, sSourcePath, sTargetPath)
                
                If Not FileOrDirExists(sTargetPath & sTmp, "") Then
                    reply = MsgBox("An error occurred while creating file:" _
                                & vbCrLf & vbCrLf & sSourceFile _
                                & vbCrLf & " to " _
                                & vbCrLf & sTargetPath & sTmp _
                                & vbCrLf & vbCrLf & app.EXEName & " will not copy the file link but to the original instead.", vbExclamation + vbOKCancel, MSG_TITLE_COPY)
                    If Not reply = vbOK Then blnAbortLinking = True
                    Exit Function
                End If
            End If
        
        Case 1 'use path as specified in INI file
            sPath = sReadIni(sIniFile, INI_SECTION_FILE, INI_KEY_FILE_CUSTOM)
            If sPath = "" Then
                reply = MsgBox("No pathname specified." _
                            & vbCrLf & vbCrLf & app.EXEName & " will not copy the file link but to the original instead.", vbExclamation + vbOKCancel, MSG_TITLE_COPY)
                If Not reply = vbOK Then blnAbortLinking = True
                Exit Function
            ElseIf IsInvalidPathSyntax(sPath) Then
                reply = MsgBox("Syntax error in pathname '" & sPath & "'" _
                             & vbCrLf & vbCrLf & app.EXEName & " will not copy the file link but to the original instead.", vbExclamation + vbOKCancel, MSG_TITLE_COPY)
                If Not reply = vbOK Then blnAbortLinking = True
                Exit Function
            Else
                sTargetPath = IIf(FileOrDirExists("", sPath), QualifyPath(sPath), CreatePath(sPath))
                Call CopySourceFile(sTmp, sSourcePath, sTargetPath)
                
                If Not FileOrDirExists(sTargetPath & sTmp, "") Then
                    reply = MsgBox("An error occurred while creating file:" _
                            & vbCrLf & vbCrLf & sSourceFile _
                            & vbCrLf & " to " _
                            & vbCrLf & sTargetPath & sTmp _
                            & vbCrLf & vbCrLf & app.EXEName & " will not copy the file link but to the original instead.", vbExclamation + vbOKCancel, MSG_TITLE_COPY)
                    If Not reply = vbOK Then blnAbortLinking = True
                    Exit Function
                End If
            End If
                
        Case 2 'database first, then INI file
            Call GetPathFromCommence(sPath, False)
            If sPath = "" Then
                sPath = sReadIni(sIniFile, INI_SECTION_FILE, INI_KEY_FILE_CUSTOM)
                If sPath = "" Then
                    reply = MsgBox("No pathname specified." _
                                & vbCrLf & vbCrLf & app.EXEName & " will not copy the file link but to the original instead.", vbExclamation + vbOKCancel, MSG_TITLE_COPY)
                    If Not reply = vbOK Then blnAbortLinking = True
                    Exit Function
                ElseIf IsInvalidPathSyntax(sPath) Then
                    reply = MsgBox("Syntax error in pathname '" & sPath & "'" _
                                 & vbCrLf & vbCrLf & app.EXEName & " will not copy the file link but to the original instead.", vbExclamation + vbOKCancel, MSG_TITLE_COPY)
                    If Not reply = vbOK Then blnAbortLinking = True
                    Exit Function
                Else
                    sTargetPath = IIf(FileOrDirExists("", sPath), QualifyPath(sPath), CreatePath(sPath))
                    Call CopySourceFile(sTmp, sSourcePath, sTargetPath)
                    
                    If Not FileOrDirExists(sTargetPath & sTmp, "") Then
                        reply = MsgBox("An error occurred while creating file:" _
                                & vbCrLf & vbCrLf & sSourceFile _
                                & vbCrLf & " to " _
                                & vbCrLf & sTargetPath & sTmp _
                                & vbCrLf & vbCrLf & app.EXEName & " will not copy the file link but to the original instead.", vbExclamation + vbOKCancel, MSG_TITLE_COPY)
                        If Not reply = vbOK Then blnAbortLinking = True
                        Exit Function
                    End If
                End If
            End If
            
        Case Else
            sLogMsg = "An error ocurred while reading Copy file options."
            Call AppendToLog(sLogFile, sLogMsg)
            MsgBox sLogMsg, vbCritical, MSG_TITLE_COPY
            blnAbortLinking = True
            Exit Function
        
    End Select

    CreateLinkFile = sLinkFile 'sLinkFile is set in other routine
    Exit Function
    
err_cannot_process:
    'return reference to original file
    CreateLinkFile = sSourceFile
    
End Function

Public Function GetPathFromFileName(ByVal s As String) As String

    Dim pos As Integer
   
    pos = InStrRev(s, "\")
    
    If pos Then
        GetPathFromFileName = Mid(s, 1, pos)
    Else
        GetPathFromFileName = s
    End If
        
End Function

Private Sub GetPathFromCommence(sPath As String, ByVal PromptIfNotFound As Boolean)
    Dim reply As Integer
    
    'default value
    sPath = ""
    'get active info from Commence
    Call GetActiveItem
    'if the active category corresponds with the category to read path from
    If sActiveItemDetails(0) = sCopyCategory Then
        If Not PromptIfNotFound Then Exit Sub 'ignore
       sPath = GetFieldValue("", "", sCopyFieldName) 'note that we don't pass the category name, we do this because then Commence will use the active item
       If sPath = "" Then
           reply = MsgBox("No path to copy file to specified in the active Commence item" _
                        & vbCrLf & vbCrLf & app.EXEName & " will not copy the file link but to the original instead.", vbExclamation + vbOKCancel, MSG_TITLE_COPY)
           If Not reply = vbOK Then blnAbortLinking = True
           sLogMsg = "No target path to copy file to specified in the active item."
           Call AppendToLog(sLogFile, sLogMsg)
           Exit Sub
       End If
    Else 'unable to read path info in Commence
        If Not PromptIfNotFound Then Exit Sub 'ignore
        reply = MsgBox("Cannot get a path to copy to." _
             & vbCrLf & "The active item in Commence is not in category '" & sCopyCategory & "'." _
             & vbCrLf & vbCrLf & app.EXEName & " will not copy the file but link to the original instead.", vbExclamation + vbOKCancel, MSG_TITLE_COPY)
        If Not reply = vbOK Then blnAbortLinking = True
        Exit Sub
    End If

End Sub

Private Function CreatePath(ByVal s As String) As String
    Dim reply As Integer
    
    CreatePath = QualifyPath(s)  'default return value
    
    If Not APICreatePath(s) Then  'create path
        reply = MsgBox("An error occurred while creating the path '" _
            & vbCrLf & vbCrLf & "'" & s & "'" _
            & vbCrLf & vbCrLf & "that is specified in the active item." _
            & vbCrLf & vbCrLf & app.EXEName & " will not copy the file link but to the original instead.", vbExclamation + vbOKCancel, MSG_TITLE_COPY)
        If Not reply = vbOK Then blnAbortLinking = True
        CreatePath = QualifyPath(s)
        Exit Function
    End If
                    
End Function

Public Sub APIShellExecute(ByVal Handle As Long, app As String)
    Call ShellExecute(Handle, vbNullString, app, vbNullString, vbNullString, vbNormalFocus)
End Sub
