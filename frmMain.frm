VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "File to Commence"
   ClientHeight    =   5190
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   8025
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   5190
   ScaleWidth      =   8025
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtCopyTo 
      Enabled         =   0   'False
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   26
      Top             =   2520
      Width           =   6735
   End
   Begin VB.CommandButton cmdDeselectAll 
      Caption         =   "Deselect all"
      Enabled         =   0   'False
      Height          =   315
      Left            =   6960
      TabIndex        =   2
      Top             =   720
      Width           =   975
   End
   Begin VB.CommandButton cmdSelectAll 
      Caption         =   "Select all"
      Enabled         =   0   'False
      Height          =   315
      Left            =   6960
      TabIndex        =   1
      Top             =   360
      Width           =   975
   End
   Begin VB.Frame frmActiveitem 
      Caption         =   "Active item in Commence:"
      Height          =   1815
      Left            =   120
      TabIndex        =   19
      ToolTipText     =   "A connection will be made with this item, if possible"
      Top             =   2880
      Width           =   3855
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "&Refresh"
         Height          =   315
         Left            =   2760
         TabIndex        =   8
         ToolTipText     =   "Get current item in Commence"
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label lblActiveItem 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   960
         TabIndex        =   25
         Top             =   720
         Width           =   2775
      End
      Begin VB.Label lblActiveCategory 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   960
         TabIndex        =   24
         Top             =   480
         Width           =   2775
      End
      Begin VB.Label Label1 
         Caption         =   "Category:"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         ToolTipText     =   "Active category in Commence"
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Item:"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         ToolTipText     =   "Active item in Commence"
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label6 
         Caption         =   "Database:"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         ToolTipText     =   "Currently opened Commence database"
         Top             =   240
         Width           =   735
      End
      Begin VB.Label lblDatabase 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   960
         TabIndex        =   20
         Top             =   240
         Width           =   2775
      End
   End
   Begin VB.Frame frmSettings 
      Caption         =   "Link files to:"
      Height          =   1815
      Left            =   4080
      TabIndex        =   13
      Top             =   2880
      Width           =   3855
      Begin VB.CommandButton cmdOptions 
         Caption         =   "Con&figure"
         Height          =   315
         Left            =   1680
         TabIndex        =   14
         ToolTipText     =   "Configure file link settings"
         Top             =   120
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CheckBox chkShowForm 
         Caption         =   "Show &detail form"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   1320
         Value           =   1  'Checked
         Width           =   2655
      End
      Begin VB.CheckBox chkFocusCommence 
         Caption         =   "Sho&w Commence when finished"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   1560
         Value           =   1  'Checked
         Width           =   2655
      End
      Begin VB.CheckBox chkCopyFile 
         Caption         =   "Copy the file(s)"
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   840
         Width           =   2655
      End
      Begin VB.CheckBox chkDeleteAfterCopy 
         Caption         =   "Delete original after copying"
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   1080
         Width           =   2655
      End
      Begin VB.Label Label4 
         Caption         =   "Field:"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         ToolTipText     =   "Field to store link in"
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Category:"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         ToolTipText     =   "Category to link files to"
         Top             =   240
         Width           =   855
      End
      Begin VB.Label lblToCategory 
         Height          =   255
         Left            =   960
         TabIndex        =   16
         Top             =   240
         Width           =   2655
      End
      Begin VB.Label lblToField 
         Height          =   255
         Left            =   960
         TabIndex        =   15
         Top             =   480
         Width           =   2655
      End
   End
   Begin VB.CheckBox chkClose 
      Caption         =   "Close this &window when done"
      Height          =   195
      Left            =   120
      TabIndex        =   9
      Top             =   4800
      Value           =   1  'Checked
      Width           =   2415
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear list"
      Height          =   315
      Left            =   6960
      TabIndex        =   3
      ToolTipText     =   "Clear files in list"
      Top             =   1680
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   315
      Left            =   6840
      TabIndex        =   11
      Top             =   4800
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   315
      Left            =   5640
      TabIndex        =   10
      ToolTipText     =   "Add selected files to Commence"
      Top             =   4800
      Width           =   975
   End
   Begin VB.ListBox lstFiles 
      Height          =   1620
      Left            =   120
      MultiSelect     =   2  'Extended
      OLEDropMode     =   1  'Manual
      Sorted          =   -1  'True
      TabIndex        =   0
      ToolTipText     =   "You can drag/drop files onto this window"
      Top             =   360
      Width           =   6735
   End
   Begin VB.Label lblSelected 
      Alignment       =   1  'Right Justify
      Caption         =   "0 files selected"
      Height          =   255
      Left            =   5160
      TabIndex        =   28
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Label lblPathCopyTo 
      Caption         =   "Output path"
      Height          =   255
      Left            =   120
      TabIndex        =   27
      ToolTipText     =   "Only available with Copy File option"
      Top             =   2295
      Width           =   1095
   End
   Begin VB.Label Label5 
      Caption         =   "Select the files you want to link to Commence"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   3375
   End
   Begin VB.Menu File 
      Caption         =   "&File"
      Begin VB.Menu Logitem 
         Caption         =   "&Log item"
      End
      Begin VB.Menu Close 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu Tools 
      Caption         =   "&Tools"
      Begin VB.Menu AlwaysOnTop 
         Caption         =   "Always on &Top"
      End
      Begin VB.Menu Refresh 
         Caption         =   "Refresh [F5]"
      End
      Begin VB.Menu ClearList 
         Caption         =   "Clear file list [F8]"
      End
      Begin VB.Menu sep1 
         Caption         =   "-"
      End
      Begin VB.Menu Configure 
         Caption         =   "&Options [F6]"
      End
   End
   Begin VB.Menu Help 
      Caption         =   "&Help"
      Begin VB.Menu ViewLogFile 
         Caption         =   "&View log file"
      End
      Begin VB.Menu sep2 
         Caption         =   "-"
      End
      Begin VB.Menu About 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const MAX_LINK_FILES As Byte = 255
Private sPathCmc As String
Private sPathUser As String

'-- Begin menu options --

Private Sub AlwaysOnTop_Click()
    
    AlwaysOnTop.Checked = Not AlwaysOnTop.Checked
    Call ToggleWindowOnTopState
    
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Select Case KeyCode
    
        Case vbKeyF5
            If Shift = 0 Then Call cmdRefresh_Click
        Case vbKeyF6
            If Shift = 0 Then Call cmdOptions_Click
        Case vbKeyF8
            If Shift = 0 Then Call cmdClear_Click
    
    End Select

End Sub

Private Sub Logitem_Click()
    Call cmdOK_Click
End Sub

Private Sub Close_Click()
    Call cmdCancel_Click
End Sub

Private Sub ClearList_Click()
    Call cmdClear_Click
End Sub

Private Sub lstFiles_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    
    Select Case KeyCode
    
        Case vbKeyDelete
            For i = lstFiles.ListCount - 1 To 0 Step -1
                If lstFiles.Selected(i) Then
                    lstFiles.RemoveItem i
                End If
            Next i
            
    End Select
    
End Sub

Private Sub Refresh_Click()
    Call cmdRefresh_Click
End Sub

Private Sub Options_Click()
    Call cmdOptions_Click
End Sub

Private Sub About_Click()

    Load frmAbout
    If Me.AlwaysOnTop.Checked Then Call SetTopMostWindow(frmAbout.hwnd, True)
    frmAbout.Show vbModal, Me
    
End Sub

' -- End menu options --

Private Sub cmdSelectAll_Click()
    Dim i As Integer
    
    With lstFiles
        For i = 0 To .ListCount - 1
            .Selected(i) = True
        Next i
    End With
    
End Sub

Private Sub cmdDeselectAll_Click()
    Dim i As Integer
    
    With lstFiles
        For i = 0 To .ListCount - 1
            .Selected(i) = False
        Next i
    End With
    
End Sub

Private Sub cmdCancel_Click()

    Unload Me 'fires Unload event which takes care of releasing stuff

End Sub

Private Sub cmdClear_Click()
    
    With Me
        .lstFiles.Clear
        .cmdOK.Enabled = False
        .cmdDeselectAll.Enabled = False
        .cmdSelectAll.Enabled = False
        .lblSelected = ""
    End With

End Sub

Private Sub chkCopyFile_Click()
    
    blnCopyFile = chkCopyFile
    chkDeleteAfterCopy.Enabled = chkCopyFile
    lblPathCopyTo.Enabled = chkCopyFile
    txtCopyTo.Enabled = chkCopyFile
    'set txtCopyTo value, note that this may not be correct if underlying Commence item has changed
    sPathUser = IIf(sCopyPath = "", "(Not specified)", sCopyPath)
    Select Case iPrecedence
        Case 0
            'display path from database
            txtCopyTo.Text = "Field '" & sCopyFieldName & "' in category '" & sCopyCategory & "': " & sPathCmc
        Case 1
            'display user defined path
            txtCopyTo.Text = sPathUser
        Case 2
            'display both path from database and user-defined path
            txtCopyTo.Text = "Field '" & sCopyFieldName & "' in category '" & sCopyCategory & "': " & sPathCmc & " OR " & sPathUser
        Case Else
            txtCopyTo.Text = "Error reading settings"
    End Select
   
End Sub

Private Sub chkDeleteAfterCopy_Click()

    blnDeleteAfterCopy = chkDeleteAfterCopy
    
End Sub

Private Sub cmdOK_Click()
    Dim i As Integer
    Dim blnFileSelected As Boolean
    Dim iSelected As Integer
    
    blnAbortLinking = False
    blnCopySucceeded = False
    
    'check if we didnt outnumber the max allowed no. of files to add
    For i = 0 To lstFiles.ListCount - 1
        If lstFiles.Selected(i) Then iSelected = iSelected + 1
    Next i

    'set a flag if we are processing more than 1 file
    'in the click event of lstFile the checkbox chkShowForm is toggled
    'if we are processing the last file, this checbox is toggled on
    'using a flag allows us to not do anything to its value,
    'yet prevent the display of the item detail form in commence
    'after successfully linking the last file
    blnMultipleFiles = (iSelected > 1)
    
    'warn user if he is about to link to the same category as the active item is in
    'this can be an indication that a prior linking has not yet been processed
    'TO BE IMPLEMENTED - IS IT USEFUL OR OBNOXIOUS?
    
    'warn user when he is about to link too many files
    'the number is arbitrary, really
    'it depends on the receiving database and esp. the number of agents an Add item triggers
    'for the selected link category
    If iSelected > MAX_LINK_FILES Then
        MsgBox "Maximum number of files you can add to " & CMC & " at one time exceeded." _
                & vbCrLf & "Please de-select files until you have at most " & MAX_LINK_FILES & " files selected", _
                vbExclamation + vbOKOnly, app.EXEName
        Exit Sub
    End If
    
    If lstFiles.ListCount = 0 Then
        MsgBox "No files to link", vbExclamation, app.EXEName
        Exit Sub
    End If
    
    For i = 0 To lstFiles.ListCount - 1
        If lstFiles.Selected(i) = True Then
            blnFileSelected = True
            Exit For
        End If
    Next i
    
    If Not blnFileSelected Then
        MsgBox "No file(s) selected for linking", vbExclamation, app.EXEName
        Exit Sub
    End If
    
    'main loop
    For i = 0 To lstFiles.ListCount - 1
        'we will remove successfully processed items from the list
        'so we need to check if we do not run out of array bounds
        If lstFiles.ListCount = 0 Or i > lstFiles.ListCount - 1 Then Exit For
        'only process items that are selected
        If lstFiles.Selected(i) = True Then
            'als Copy is aangevinkt moeten we bestand eerst kopieren alvorens te linken
            'de link moet dan het gekopieerde bestand zijn!
            'deze linken aan Commence
            'dit doen we allemaal op andere plek
            If blnCopyFile Then
                sLinkFile = CreateLinkFile(lstFiles.List(i))
                If blnAbortLinking Then Exit For
            Else
                sLinkFile = lstFiles.List(i)
            End If
            
            'delete file after copying
            If sLinkFile <> lstFiles.List(i) And blnDeleteAfterCopy And blnCopySucceeded Then
                Call APIDeleteFile(lstFiles.List(i), FOF_ALLOWUNDO + FOF_NOCONFIRMATION)
            End If
            
            If Not AddFileToCommence(sLinkFile) Then
                GoTo err_handler
            Else
                'display (new) path and filename of successfully processed file
                'lstFiles.List(i) = sLinkFile
                lstFiles.Selected(i) = False
                'how shall we deal with items from the list that were processed?
                lstFiles.RemoveItem i
                i = i - 1
                'update the main form so as to show that a new item is now the active item
                Me.RefreshCommenceItem
            End If
        End If
        DoEvents
    Next i

    'put focus on Commence
    If Me.chkFocusCommence Then
        Dim hwnd As Long
        On Error Resume Next 'dont care about errors
        hwnd = FindWindow(CMC, vbNullString)
        Call ShowWindow(hwnd, SW_SHOWNORMAL)  'maximize window
        Call SetForegroundWindow(hwnd) 'move to top
        On Error GoTo -1
    End If
    
    If Me.chkClose Then Unload Me
    Exit Sub
    
err_handler:
    MsgBox "An error occured while adding files to Commence", vbExclamation, app.EXEName
    
End Sub

Private Sub cmdOptions_Click()

    Load frmOptions
    If Me.AlwaysOnTop.Checked Then Call SetTopMostWindow(frmOptions.hwnd, True)
    frmOptions.Show vbModal, Me
    
End Sub

Private Sub cmdRefresh_Click()

  Call RefreshCommenceItem

End Sub

Private Sub Configure_Click()

    Load frmOptions
    If Me.AlwaysOnTop.Checked Then Call SetTopMostWindow(frmOptions.hwnd, True)
    frmOptions.Show vbModal, Me

End Sub

Private Sub lstFiles_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Integer
    
    If Data.GetFormat(15) Then
        For i = 1 To Data.Files.Count
            Me.lstFiles.AddItem Data.Files(i)
            Me.lstFiles.Selected(lstFiles.NewIndex) = True
        Next i
    End If
    
    Me.cmdOK.Enabled = True

End Sub

Private Sub lstFiles_Click()
    'if > 1 files selected, disable option to show detail form
    
    Dim i As Integer
    Dim iSelected As Integer
    iSelected = 0

    For i = 0 To Me.lstFiles.ListCount - 1
        If lstFiles.Selected(i) = True Then
            iSelected = iSelected + 1
        End If

    Next i
        
    With Me
        chkShowForm.Enabled = IIf(iSelected > 1, False, True)
        cmdSelectAll.Enabled = IIf(lstFiles.ListCount > 0, True, False)
        cmdDeselectAll.Enabled = IIf(lstFiles.ListCount > 0, True, False)
    End With
    
    lblSelected.Caption = iSelected & " of " & Me.lstFiles.ListCount & " selected"
    
End Sub

Public Sub RefreshCommenceItem()
    'Dim sTmp As String 'temp string
    'Dim s As String 'temp string
    
    sPathCmc = sCopyFieldName 'put fieldname in string, string returns as field value
    Call GetActiveItem(sPathCmc)
    
    If bNoAutoConnect Then
        lblActiveCategory = "connecting to item is disabled"
        lblActiveItem = "connecting to item is disabled"
    Else
        lblActiveCategory = sActiveItemDetails(0)
        lblActiveItem = sActiveItemDetails(1)
    End If
    
    lblActiveCategory.ToolTipText = lblActiveCategory
    lblActiveItem.ToolTipText = ParseOutSpaces(lblActiveItem)
    lblDatabase = oCmc.Name
    'if there is a copy attempt, display where to
    If blnCopyFile Then
        sPathUser = IIf(sCopyPath = "", "(Not specified)", sCopyPath)
        Select Case iPrecedence
            Case 0
                'display path from database
                txtCopyTo.Text = "Field '" & sCopyFieldName & "' in category '" & sCopyCategory & "': " & sPathCmc
            Case 1
                'display user defined path
                txtCopyTo.Text = sPathUser
            Case 2
                'display both path from database and user-defined path
                txtCopyTo.Text = "Field '" & sCopyFieldName & "' in category '" & sCopyCategory & "': " & sPathCmc & " OR " & sPathUser
            Case Else
                txtCopyTo.Text = "Error reading settings"
        End Select
    End If
    
    sLogMsg = "Active Commence database is: '" & lblDatabase & "'."
    Call AppendToLog(sLogFile, sLogMsg)
    sLogMsg = "Active item is: '" & lblActiveItem.ToolTipText & "'."
    Call AppendToLog(sLogFile, sLogMsg)
    
End Sub
Private Sub Form_Load()

    'initialization routines
    If Not Init Then
        Unload Me
        Exit Sub
    End If
    
    'get a list of what files are to be added and display it
    If Not PopulateFileList Then
        sLogMsg = "An error occurred during the processing of the file(s) list."
        Call AppendToLog(sLogFile, sLogMsg)
        MsgBox "Error processing filename(s)", vbCritical, app.EXEName
        Unload Me
        Exit Sub
    End If
    
    'see if INI settings correspond to current database
    If Not ValidateINISettings Then
        sLogMsg = "Invalid option detected in " & sIniFile & "."
        Call AppendToLog(sLogFile, sLogMsg)
        MsgBox "Your current settings do not correspond with the active database." _
                & vbCrLf & "Select a category and/or field to log files to before you continue.", vbExclamation, app.EXEName
        Unload Me
        Load frmOptions
        frmOptions.Show
        frmOptions.SetFocus
        Exit Sub
    End If
    
    Call ReadOptions
    Call RefreshCommenceItem
    If bNoAutoConnect Then Me.cmdOK.Enabled = True 'if user has selected to not automatically create a connection, enable OK button anyway
    Call ToggleWindowOnTopState
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Call Cleanup
    Unload frmOptions
    Unload frmAbout
    Unload frmConflictingFile

End Sub

Private Sub Form_Terminate()

    sLogMsg = app.EXEName & " terminated."
    Call AppendToLog(sLogFile, sLogMsg)
    
End Sub

Public Sub ReadOptions()

    On Error GoTo err_handler
    lblToCategory = sReadIni(sIniFile, INI_SECTION_CMC, INI_KEY_CMC_CAT)
    lblToCategory.ToolTipText = lblToCategory
    lblToField = sReadIni(sIniFile, INI_SECTION_CMC, INI_KEY_CMC_DATAFILE)
    lblToField.ToolTipText = lblToField
    chkCopyFile.Value = sReadIni(sIniFile, INI_SECTION_FILE, INI_KEY_FILE_COPY) 'causes Click event in control!
    blnCopyFile = chkCopyFile.Value
    iPrecedence = sReadIni(sIniFile, INI_SECTION_FILE, INI_KEY_FILE_ORDER)
    sCopyCategory = sReadIni(sIniFile, INI_SECTION_FILE, INI_KEY_FILE_CAT)
    sCopyFieldName = sReadIni(sIniFile, INI_SECTION_FILE, INI_KEY_FILE_FIELD)
    sCopyPath = sReadIni(sIniFile, INI_SECTION_FILE, INI_KEY_FILE_CUSTOM)
    chkDeleteAfterCopy.Value = sReadIni(sIniFile, INI_SECTION_FILE, INI_KEY_FILE_DELETEORIGINAL) 'causes Click event in control!
    blnDeleteAfterCopy = chkDeleteAfterCopy.Value
    chkShowForm = sReadIni(sIniFile, INI_SECTION_OTHER, INI_KEY_OTHER_SHOWIDF) 'causes Click event in control!
    chkFocusCommence = sReadIni(sIniFile, INI_SECTION_OTHER, INI_KEY_OTHER_SHOWCMC) 'causes Click event in control!
    chkClose = sReadIni(sIniFile, INI_SECTION_DISPLAY, INI_KEY_DISPLAY_CLOSE) 'causes Click event in control!
    AlwaysOnTop.Checked = CBool(sReadIni(sIniFile, INI_SECTION_DISPLAY, INI_KEY_DISPLAY_ONTOP))
    bNoAutoConnect = CBool(sReadIni(sIniFile, INI_SECTION_OTHER, INI_KEY_OTHER_AUTOCON))
    cmdRefresh.Enabled = Not bNoAutoConnect 'note that it is set to false if enabled is true and vice versa
    iWhatIfFileNewer = sReadIni(sIniFile, INI_SECTION_FILE, INI_KEY_FILE_NEWER)
    iWhatIfAttrDifferent = sReadIni(sIniFile, INI_SECTION_FILE, INI_KEY_FILE_ATTR)
    txtCopyTo.BackColor = Me.BackColor
    Exit Sub

err_handler:
    sLogMsg = "An error occurred while reading settings from '" & sIniFile & "'."
    Call AppendToLog(sLogFile, sLogMsg)
    MsgBox "Error reading settings from '" & sIniFile & "'." _
        & vbCrLf & vbCrLf & "Probable cause: you have upgraded to a newer version of '" & app.EXEName & "'." _
        & vbCrLf & "Simply select Tools | Options from the menu to re-configure." _
        & vbCrLf & vbCrLf & "If the problem persists then manually delete this file." _
        & vbCrLf & "This will not in any way harm your system.", vbExclamation, app.EXEName
    
End Sub

Private Function PopulateFileList() As Boolean
    Dim sArg As String
    Dim sTmp As String
    Dim i As Integer
    Dim iSelected As Integer
    
    'the Command exists of filenames separated by space.
    'if the filename itself contains a space, it is enclosed in double quotes
    
    sArg = Command
    For i = 1 To Len(sArg)
        If Mid(sArg, i, 1) = Chr(34) Then
            'skip and loop until next "
            i = i + 1
            Do While Not Mid(sArg, i, 1) = Chr(34)
                sTmp = sTmp & Mid(sArg, i, 1)
                i = i + 1
            Loop
            lstFiles.AddItem Trim(sTmp)
            sTmp = ""
            i = i + 1
        Else
            Do While Not Mid(sArg, i, 1) = " " And i <= Len(sArg)
                sTmp = sTmp & Mid(sArg, i, 1)
                i = i + 1
            Loop
            lstFiles.AddItem Trim(sTmp)
            sTmp = ""
        End If
    Next i
    
    'make list items selected
    For i = 0 To lstFiles.ListCount - 1
        lstFiles.Selected(i) = True
        iSelected = i
    Next i

    'if more than 1 file selected for linking, disable option to show detail form
    Me.chkShowForm.Enabled = IIf(iSelected > 1, False, True)
    PopulateFileList = True

End Function

Public Sub ToggleWindowOnTopState()
    'deze lijkt raar om public te hebben, keertje naar kijken?
    Call SetTopMostWindow(Me.hwnd, Me.AlwaysOnTop.Checked)
End Sub

Private Sub ViewLogFile_Click()

    'Call APIShellExecute(Me.hwnd, sLogFile)
    Load frmOptions
    frmOptions.tvwOptions.Nodes("history").Selected = True
    frmOptions.Frame(3).Visible = True
    frmOptions.rtfLog.Text = GetLogText
    frmOptions.Show vbModal, Me
    
End Sub

