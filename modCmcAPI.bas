Attribute VB_Name = "modCmcAPI"
Option Explicit
' the actual Add Item in Commence using the Cmc API

'Subset of Commence.DB constants
Public Const CMC_MAX_DATAFILE_LENGTH As Integer = 279 'Cmc datafile field only takes this many characters!
Public Const CMC_MAX_NAME_LENGTH As Integer = 50
Public Const CMC_ERR_API  As String = "Commence Automation error"
Public Const CMC_NAME_MODULE As String = "COMMENCE.EXE"
Public Const CMC_CLS_NAME = "Commence.DB"

'DDE Stuff
Public Const CMC As String = "Commence" 'dde appl name
Public Const CMC_DELIM As String = "||||||||" 'delimiter
Public Const CMC_FIELDTYPE_NAME As Long = 11 'name field type identifier
Public Const CMC_FIELDTYPE_DATA As Long = 12 'data file field type identifier
Public Const CMC_FIELDTYPE_URL As Long = 24 'internet address field type identifier
Public Const CMC_FIELDTYPE_TEXT As Long = 0 'text field type identifier

'string arrays etc.
Public sActiveItemDetails(1) As String
Public sFormNames() As String
Public sCategories() As String

'Commence variables
Public oCmc As CommenceDB
Public oConv As ICommenceConversation
Public strDDE As String 'dde command

'misc. variables
Public blnMultipleFiles As Boolean

Function InstantiateCmc() As Boolean
    'get a reference to Cmc
    InstantiateCmc = True
    On Error Resume Next
    
    Set oCmc = GetObject(, CMC_CLS_NAME)
    If oCmc Is Nothing Then
        Set oCmc = CreateObject(CMC_CLS_NAME)
    End If
    
    If oCmc Is Nothing Then InstantiateCmc = False

End Function

Sub GetActiveItem(Optional s As String)
    'opens a connection to Commence to see what item is currently active
    'returns s as fieldvalue for field s

    Dim buffer
    Dim dde_error As Integer
    On Error Resume Next
    
    'record that we are trying to read from Commence
    sLogMsg = "Requested refresh of Commence data:"
    Call AppendToLog(sLogFile, sLogMsg)
    
    Set oConv = oCmc.GetConversation("Commence", "GetData") 'open dde connection
    'use clarify field value if possible
    strDDE = "[ClarifyItemNames(True)]"
    buffer = oConv.Request(strDDE)
    'mark the active item
    strDDE = "[MarkActiveItem]"
    buffer = oConv.Request(strDDE)
    strDDE = "[GetLastError]"
    dde_error = oConv.Request(strDDE)
    If Not dde_error = 0 Then
        MsgBox "Unable to determine active item in " & CMC, vbExclamation, app.EXEName
        frmMain.cmdOK.Enabled = False
        sLogMsg = "Unable to get active item." & CMC & " returned DDE error: " & dde_error & "."
        Call AppendToLog(sLogFile, sLogMsg)
        GoTo err_handler
    End If
    strDDE = "[GetActiveViewInfo(" & CMC_DELIM & " )]"
    'returns {ViewName}Delim {ViewType}Delim {CategoryName}Delim {ItemName}Delim {FieldName}
    buffer = Split(oConv.Request(strDDE), CMC_DELIM)
    sActiveItemDetails(0) = buffer(2) 'category
    sActiveItemDetails(1) = buffer(3) 'itemname
    On Error Resume Next
    strDDE = "[GetField(,," & dq(s) & ")]"
    buffer = oConv.Request(strDDE)
    s = buffer
    
err_handler:
    Set oConv = Nothing
    
End Sub

Function AddFileToCommence(ByVal s As String) As Boolean
    'adds file with filename s to log category in commence
    Const CON_CAT_DELIM As String = "!!!!"
    Dim i As Integer
    Dim sCat As String
    Dim sConCat As String
    Dim sConName As String
    Dim sConItem As String
    Dim sCons() As String
    Dim buffer() As String
    Dim buffer2() As String
    Dim bInitializedDimension As Boolean
    Dim sItemName As String
    Dim iLastError As Long
    
    sItemName = APIGetFileTitle(s)
    If sItemName = "" Then GoTo err_filetitle
    sItemName = Left(sItemName, CMC_MAX_NAME_LENGTH)
    
    On Error GoTo err_handler
    'determine what category to use
    sCat = sReadIni(sIniFile, INI_SECTION_CMC, INI_KEY_CMC_CAT)
    'obtain fieldnames to use
    ReDim sFields(1)
    sFields(0) = sReadIni(sIniFile, INI_SECTION_CMC, INI_KEY_CMC_NAMEFIELD)
    sFields(1) = sReadIni(sIniFile, INI_SECTION_CMC, INI_KEY_CMC_DATAFILE)    'determine category of active item
    sConCat = frmMain.lblActiveCategory.Caption
    'determine active item
    sConItem = frmMain.lblActiveItem.Caption
    'get connection names for that category
    strDDE = "[GetConnectionNames(" & dq(sCat) & "," & CMC_DELIM & "," & CON_CAT_DELIM & " )]"
    Set oConv = oCmc.GetConversation(CMC, "GetData")
    buffer = Split(oConv.Request(strDDE), CMC_DELIM)
    For i = 0 To UBound(buffer)

        buffer2 = Split(buffer(i), CON_CAT_DELIM)
        If buffer2(1) = sConCat Then 'see if connection is to the log category
            'there may be more than 1 connection to the log category, acommodate for that
            If Not bInitializedDimension Then
                ReDim sCons(0) As String
                sCons(0) = buffer2(0) 'connection
                bInitializedDimension = True
            Else
                ReDim Preserve sCons(UBound(sCons) + 1) As String 'can only resize last element when using Preserve!
                sCons(UBound(sCons)) = buffer2(0)
            End If
        End If
    Next i
    
     'do the add item
    strDDE = "[AddItem(" & dq(sCat) & "," & dq(sItemName) & ")]"
    oConv.Execute strDDE
    strDDE = "[GetLastError()]"
    iLastError = oConv.Request(strDDE)
    If iLastError <> 0 Then Call AppendToLog(sLogFile, "Commence returned DDE Error: " & iLastError)
    strDDE = "[EditItem(,," & dq(sReadIni(sIniFile, INI_SECTION_CMC, INI_KEY_CMC_DATAFILE)) & "," & dq(s) & ")]"
    oConv.Execute strDDE
    strDDE = "[GetLastError()]"
    iLastError = oConv.Request(strDDE)
    If iLastError <> 0 Then Call AppendToLog(sLogFile, "Commence returned DDE Error: " & iLastError)

    'create connections
    If Not bNoAutoConnect Then
        If IsDimensioned(sCons) Then
            For i = 0 To UBound(sCons)
                strDDE = "[AssignConnection(,," & dq(sCons(i)) & "," & dq(sConCat) & "," & dq(sConItem) & ")]"
                oConv.Execute strDDE
            Next i
            strDDE = "[GetLastError()]"
            iLastError = oConv.Request(strDDE)
            If iLastError <> 0 Then Call AppendToLog(sLogFile, "Commence returned DDE Error: " & iLastError)
        End If
    End If
    'make item shared
    If sReadIni(sIniFile, INI_SECTION_OTHER, INI_KEY_OTHER_SHARE) = 1 Then
        strDDE = "[PromoteItemToShared(,)]"
        oConv.Execute strDDE
        strDDE = "[GetLastError()]"
        iLastError = oConv.Request(strDDE)
        If iLastError <> 0 Then Call AppendToLog(sLogFile, "Commence returned DDE Error: " & iLastError)
    End If
    'show detail form
    If frmMain.chkShowForm.Enabled And frmMain.chkShowForm And Not blnMultipleFiles Then
        strDDE = "[ShowItem(,," & dq(sReadIni(sIniFile, INI_SECTION_CMC, INI_KEY_CMC_IDF)) & ")]"
        oConv.Execute strDDE
        strDDE = "[GetLastError()]"
        iLastError = oConv.Request(strDDE)
        If iLastError <> 0 Then Call AppendToLog(sLogFile, "Commence returned DDE Error: " & iLastError)
    End If
    
    'throw a DDE event end-users can use in agents to deal with the newly added item
    On Error Resume Next
    If sReadIni(sIniFile, INI_SECTION_OTHER, INI_KEY_OTHER_DDE) = 0 Then
        strDDE = "[FireTrigger(" & dq("File2CmcFinished") & "," & dq(sCat) & "," & dq(sItemName) & ")]"
        oConv.Execute strDDE
    End If
    On Error GoTo err_handler
    
    sLogMsg = "Item '" & sItemName & "' successfully logged to '" & sCat & "'."
    Call AppendToLog(sLogFile, sLogMsg)
    AddFileToCommence = True
    Set oConv = Nothing
    Exit Function
    
err_filetitle:
    sLogMsg = "An error occurred obtaining the file title for: '" & s & "'. Logging was aborted."
    Call AppendToLog(sLogFile, sLogMsg)
    MsgBox "There was a problem obtaining the file title for: " _
        & vbCrLf & "'" & s & "'", vbExclamation, app.EXEName
    
err_handler:
    If iLastError > 0 Then
        sLogMsg = "Error &" & iLastError & " occurred while processing file '" & s & "'."
        Call AppendToLog(sLogFile, sLogMsg)
    End If
    Set oConv = Nothing
    
End Function

Function ValidateINISettings() As Boolean
    Dim buffer() As String
    Dim i As Integer
    Dim X As String, Y As String, Z As String
    
    Set oConv = oCmc.GetConversation(CMC, "GetData")
    X = sReadIni(sIniFile, INI_SECTION_CMC, INI_KEY_CMC_CAT)
    strDDE = "[GetCategoryNames(" & CMC_DELIM & ")]"
    buffer = Split(oConv.Request(strDDE), CMC_DELIM)
    For i = LBound(buffer) To UBound(buffer)
        If buffer(i) = X Then Exit For
    Next i
    
    If i - 1 = UBound(buffer) Then Exit Function
    
    strDDE = "[GetFieldNames(" & dq(X) & "," & CMC_DELIM & ")]"
    buffer = Split(oConv.Request(strDDE), CMC_DELIM)
    Y = sReadIni(sIniFile, INI_SECTION_CMC, INI_KEY_CMC_NAMEFIELD)
    
    For i = LBound(buffer) To UBound(buffer)
        If buffer(i) = Y Then Exit For
    Next i
    
    If i - 1 = UBound(buffer) Then Exit Function
    
    Y = sReadIni(sIniFile, INI_SECTION_CMC, INI_KEY_CMC_DATAFILE)
    
    For i = LBound(buffer) To UBound(buffer)
        If buffer(i) = Y Then Exit For
    Next i
    
    If i - 1 = UBound(buffer) Then Exit Function
    
    Z = sReadIni(sIniFile, INI_SECTION_CMC, INI_KEY_CMC_IDF)
    strDDE = "[GetFormNames(" & dq(X) & "," & CMC_DELIM & ")]"
    buffer = Split(oConv.Request(strDDE), CMC_DELIM)
    
    For i = LBound(buffer) To UBound(buffer)
        If buffer(i) = Z Then Exit For
    Next i
    
    If i - 1 = UBound(buffer) Then Exit Function
    
    ValidateINISettings = True
    Set oConv = Nothing
    
End Function

Public Function GetFieldValue(ByVal sCat As String, _
                    ByVal sItem As String, _
                    ByVal sField As String) _
                    As String
                    
    On Error Resume Next
    Set oConv = oCmc.GetConversation(CMC, "GetData")
    strDDE = "[MarkActiveItem]"
    oConv.Request strDDE
    strDDE = "[GetField(" & dq(sCat) & "," & dq(sItem) & "," & dq(sField) & ")]"
    GetFieldValue = oConv.Request(strDDE)
    Set oConv = Nothing

End Function

Public Function GetConnections(ByVal cat As String) As Boolean
    'retrieve a list of connections and the categories from Commence
    'compose a 2-dimensional array to hold them
    Const delim As String = "~!@##@!~" 'delimiter to separate connection from category. Note that it must not be the same as the default delimiter CMC_DELIM
    Dim buffer() As String
    Dim item() As String
    Dim i As Integer
    
    Set oConv = oCmc.GetConversation(CMC, "GetData") 'open dde connection
    
    'first check if there any connections to this category
    strDDE = "[GetConnectionCount(" & dq(cat) & "," & CMC_DELIM & ")]"
    i = oConv.Request(strDDE)
    If i = 0 Then GoTo error_handler 'exit if no connections are present

    strDDE = "[GetConnectionNames(" & dq(cat) & "," & CMC_DELIM & "," & delim & ")]"
    buffer = Split(oConv.Request(strDDE), CMC_DELIM)
    'buffer now holds "connectionname1<delim>categoryname1, connectionname2<delim>categoryname2, connectionnameN<delim>categorynameN"
    ReDim Connections(0 To UBound(buffer), 0 To 1) As Variant
    For i = 0 To UBound(buffer)
        item = Split(buffer(i), delim)
        Connections(i, 0) = item(0)
        Connections(i, 1) = item(1)
    Next i
    GetConnections = True
    Set oConv = Nothing
    Exit Function
    
error_handler:
    GetConnections = False
    
End Function

Public Sub GetCategories()
    'returns array of category names
    Dim i As Integer

    Set oConv = oCmc.GetConversation(CMC, "GetData") 'open dde connection
    strDDE = "[GetCategoryNames(" & CMC_DELIM & ")]" 'construct strDDE command
    sCategories = Split(oConv.Request(strDDE), CMC_DELIM) 'get category names from commence and split them into array
    Set oConv = Nothing

End Sub

Public Sub GetDataFields(ByVal cat As String)
    'parse all fields for category
    'build array of relevant fields
    'sFields(0) will hold name field
    'sFields(1) will hold data file field or url field(!). If more than 1, Fields will be expanded accordingly
    
    Dim buffer() As String
    Dim parseFields() As String
    Dim i As Integer
    
    Set oConv = oCmc.GetConversation(CMC, "GetData")
    
    ReDim sFields(0) As String
    
    strDDE = "[GetFieldNames(" & dq(cat) & "," & CMC_DELIM & ")]"
    parseFields() = Split(oConv.Request(strDDE), CMC_DELIM)
    sFields(0) = parseFields(0) 'name field
    
    For i = LBound(parseFields) To UBound(parseFields)
        strDDE = "[GetFieldDefinition(" & dq(cat) & "," & dq(parseFields(i)) & "," & CMC_DELIM & ")]"
        'GetFieldDefinition returns {FieldType}Delim 000000{C}{S}{M}{R}Delim {MaxChars}Delim {DefaultString}
        buffer = Split(oConv.Request(strDDE), CMC_DELIM)
        If (buffer(0) = CMC_FIELDTYPE_DATA Or buffer(0) = CMC_FIELDTYPE_URL) Then
                ReDim Preserve sFields(UBound(sFields) + 1) As String
                sFields(UBound(sFields)) = parseFields(i)
        End If
    Next i
    
    Set oConv = Nothing
        
End Sub

Public Sub GetTextEnabledFields(ByVal cat As String)
    'parse all fields for category
    'build array of text enabled fields
    
    Dim buffer() As String
    Dim parseFields() As String
    Dim i As Integer
    
    Set oConv = oCmc.GetConversation(CMC, "GetData")
    
    ReDim sTextFields(0) As String
    
    strDDE = "[GetFieldNames(" & dq(cat) & "," & CMC_DELIM & ")]"
    parseFields() = Split(oConv.Request(strDDE), CMC_DELIM)
    'sTextFields(0) = parseFields(0) 'name field
    
    For i = LBound(parseFields) To UBound(parseFields)
        strDDE = "[GetFieldDefinition(" & dq(cat) & "," & dq(parseFields(i)) & "," & CMC_DELIM & ")]"
        'GetFieldDefinition returns {FieldType}Delim 000000{C}{S}{M}{R}Delim {MaxChars}Delim {DefaultString}
        buffer = Split(oConv.Request(strDDE), CMC_DELIM)
        If (buffer(0) = CMC_FIELDTYPE_TEXT Or buffer(0) = CMC_FIELDTYPE_URL) Then
                ReDim Preserve sTextFields(UBound(sTextFields) + 1) As String
                sTextFields(UBound(sTextFields)) = parseFields(i)
        End If
    Next i
    
    Set oConv = Nothing
        
End Sub

Public Function GetDetailForms(ByVal strCat As String) As String
    Dim i As Integer
    Dim buffer() As String 'temporary array
    Dim s As String
    
    Set oConv = oCmc.GetConversation(CMC, "GetData") 'open dde connection
    
    strDDE = "[GetFormNames(" & dq(strCat) & "," & CMC_DELIM & ")]"
    buffer = Split(oConv.Request(strDDE), CMC_DELIM)
    sFormNames = buffer
    
    Set oConv = Nothing
    
End Function
