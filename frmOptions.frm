VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmOptions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Options"
   ClientHeight    =   9270
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12360
   Icon            =   "frmOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9270
   ScaleWidth      =   12360
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   5040
      TabIndex        =   41
      Top             =   4800
      Width           =   855
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   6000
      TabIndex        =   40
      Top             =   4800
      Width           =   855
   End
   Begin VB.Frame Frame 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   4695
      Index           =   3
      Left            =   7440
      TabIndex        =   35
      Top             =   3720
      Width           =   4575
      Begin VB.Frame frmLog 
         Caption         =   "Log"
         Height          =   4575
         Left            =   0
         TabIndex        =   36
         Top             =   0
         Width           =   4575
         Begin VB.CommandButton cmdPurgeLog 
            Caption         =   "Purge..."
            Height          =   375
            Left            =   1440
            TabIndex        =   38
            Top             =   4080
            Width           =   1255
         End
         Begin VB.CommandButton cmdViewLog 
            Caption         =   "Open &Logfile"
            Height          =   375
            Left            =   120
            TabIndex        =   37
            ToolTipText     =   "Open log file in text editor"
            Top             =   4080
            Width           =   1255
         End
         Begin RichTextLib.RichTextBox rtfLog 
            Height          =   3735
            Left            =   120
            TabIndex        =   39
            Top             =   240
            Width           =   4335
            _ExtentX        =   7646
            _ExtentY        =   6588
            _Version        =   393217
            BorderStyle     =   0
            Enabled         =   -1  'True
            ReadOnly        =   -1  'True
            ScrollBars      =   3
            RightMargin     =   8192
            TextRTF         =   $"frmOptions.frx":030A
         End
      End
   End
   Begin VB.Frame Frame 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   3495
      Index           =   0
      Left            =   2520
      TabIndex        =   24
      Top             =   5520
      Width           =   4575
      Begin VB.Frame Frame1 
         Caption         =   "Other options"
         Height          =   1215
         Index           =   1
         Left            =   0
         TabIndex        =   34
         Top             =   1680
         Width           =   4575
         Begin VB.CheckBox chkNoDDETrigger 
            Caption         =   "Do not send DDE triggers to Commence"
            Height          =   255
            Left            =   120
            TabIndex        =   43
            Top             =   800
            Width           =   4215
         End
         Begin VB.CheckBox chkShared 
            Caption         =   "&Share item(s) by default"
            Height          =   255
            Left            =   120
            TabIndex        =   3
            Top             =   240
            Value           =   1  'Checked
            Width           =   3855
         End
         Begin VB.CheckBox chkAutoConnect 
            Caption         =   "Do not create co&nnections automatically"
            Height          =   315
            Left            =   120
            TabIndex        =   4
            ToolTipText     =   "Do not connect the linked file to the active Commence item"
            Top             =   480
            Width           =   3855
         End
      End
      Begin VB.Frame frmCmcOptions 
         Caption         =   "Link files to:"
         Height          =   1575
         Left            =   0
         TabIndex        =   25
         Top             =   0
         Width           =   4575
         Begin VB.ComboBox cboDetailForm 
            Height          =   315
            Left            =   1080
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   1080
            Width           =   2895
         End
         Begin VB.ComboBox cboCategory 
            Height          =   315
            Left            =   1080
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   360
            Width           =   2895
         End
         Begin VB.ComboBox cboDataFile 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1080
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   720
            Width           =   2895
         End
         Begin VB.Label Label1 
            Caption         =   "C&ategory"
            Height          =   255
            Left            =   120
            TabIndex        =   32
            ToolTipText     =   "Category to store the file info in"
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label2 
            Caption         =   "&Field"
            Height          =   255
            Left            =   120
            TabIndex        =   30
            ToolTipText     =   "Field that will hold the link to the file"
            Top             =   720
            Width           =   855
         End
         Begin VB.Label Label3 
            Caption         =   "&Detail Form"
            Height          =   375
            Left            =   120
            TabIndex        =   26
            ToolTipText     =   "Detail form to open in Commence after item was written"
            Top             =   1080
            Width           =   975
         End
      End
   End
   Begin VB.Frame Frame 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   3495
      Index           =   2
      Left            =   5520
      TabIndex        =   22
      Top             =   480
      Width           =   4575
      Begin VB.CommandButton cmdCreateShortcut 
         Caption         =   "Add to SendTo"
         Height          =   375
         Left            =   0
         TabIndex        =   33
         ToolTipText     =   "Create shortcut to File2Cmc in Sendto Menu"
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Frame Frame3 
         Caption         =   "Display"
         Height          =   1575
         Left            =   0
         TabIndex        =   23
         Top             =   0
         Width           =   4575
         Begin VB.CheckBox chkCommence 
            Caption         =   "Show Co&mmence when done"
            Height          =   195
            Left            =   120
            TabIndex        =   29
            Top             =   820
            Value           =   1  'Checked
            Width           =   3855
         End
         Begin VB.CheckBox chkShowDetailForm 
            Caption         =   "Sho&w item detail form after linking"
            Height          =   315
            Left            =   120
            TabIndex        =   28
            Top             =   480
            Value           =   1  'Checked
            Width           =   3855
         End
         Begin VB.CheckBox chkClose 
            Caption         =   "Close &window when done"
            Height          =   255
            Left            =   120
            TabIndex        =   27
            Top             =   240
            Value           =   1  'Checked
            Width           =   3855
         End
         Begin VB.CheckBox chkAlwaysOnTop 
            Caption         =   "Always on &Top"
            Height          =   255
            Left            =   120
            TabIndex        =   31
            Top             =   1080
            Width           =   2175
         End
      End
   End
   Begin VB.Frame Frame 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   4695
      Index           =   1
      Left            =   2400
      TabIndex        =   12
      Top             =   60
      Width           =   4575
      Begin VB.Frame frmFileOperations 
         Caption         =   "Copy options"
         Height          =   4575
         Left            =   0
         TabIndex        =   13
         Top             =   0
         Width           =   4575
         Begin VB.Frame frmCopyOptions 
            Caption         =   "Order of precedence"
            Enabled         =   0   'False
            Height          =   975
            Left            =   120
            TabIndex        =   17
            ToolTipText     =   "If no path settings are found, the file is not copied"
            Top             =   3240
            Width           =   4215
            Begin VB.OptionButton optPathOrder 
               Caption         =   "Use path in database field"
               Enabled         =   0   'False
               Height          =   195
               Index           =   0
               Left            =   120
               TabIndex        =   18
               Top             =   240
               Width           =   3735
            End
            Begin VB.OptionButton optPathOrder 
               Caption         =   "Use user specified path"
               Enabled         =   0   'False
               Height          =   195
               Index           =   1
               Left            =   120
               TabIndex        =   19
               Top             =   480
               Value           =   -1  'True
               Width           =   3735
            End
            Begin VB.OptionButton optPathOrder 
               Caption         =   "Try database field first, then user specified path"
               Enabled         =   0   'False
               Height          =   195
               Index           =   2
               Left            =   120
               TabIndex        =   20
               Top             =   720
               Width           =   3735
            End
         End
         Begin VB.CommandButton cmdBrowseForFolder 
            Caption         =   "Select..."
            Enabled         =   0   'False
            Height          =   285
            Left            =   3600
            TabIndex        =   11
            Top             =   2760
            Width           =   735
         End
         Begin VB.TextBox txtPath 
            Enabled         =   0   'False
            Height          =   285
            Left            =   240
            TabIndex        =   10
            Top             =   2760
            Width           =   3015
         End
         Begin VB.CheckBox chkCopyFile 
            Caption         =   "Copy the file to specified location"
            Height          =   255
            Left            =   240
            TabIndex        =   5
            Top             =   240
            Width           =   3015
         End
         Begin VB.CommandButton cmdOptions 
            Caption         =   "Options"
            Height          =   285
            Left            =   3600
            TabIndex        =   6
            Top             =   240
            Width           =   735
         End
         Begin VB.Frame frmDatabaseCopyFilePath 
            Caption         =   "To path stored in database item"
            Height          =   1575
            Left            =   120
            TabIndex        =   14
            Top             =   840
            Width           =   4215
            Begin VB.ComboBox cboLinkPathCategory 
               Enabled         =   0   'False
               Height          =   315
               Left            =   120
               Style           =   2  'Dropdown List
               TabIndex        =   8
               Top             =   480
               Width           =   3135
            End
            Begin VB.ComboBox cboLinkPathFieldName 
               Enabled         =   0   'False
               Height          =   315
               Left            =   120
               Style           =   2  'Dropdown List
               TabIndex        =   9
               Top             =   1080
               Width           =   3135
            End
            Begin VB.Label lblCategory 
               Caption         =   "category"
               Enabled         =   0   'False
               Height          =   255
               Left            =   120
               TabIndex        =   16
               Top             =   240
               Width           =   735
            End
            Begin VB.Label lblField 
               Caption         =   "field"
               Enabled         =   0   'False
               Height          =   255
               Left            =   120
               TabIndex        =   15
               Top             =   840
               Width           =   3135
            End
         End
         Begin VB.CheckBox chkDeleteAfterCopy 
            Caption         =   "Delete original after copying"
            Enabled         =   0   'False
            Height          =   255
            Left            =   240
            TabIndex        =   7
            Top             =   480
            Width           =   3015
         End
         Begin VB.Label lblCopyPath 
            Caption         =   "To this path"
            Enabled         =   0   'False
            Height          =   255
            Left            =   240
            TabIndex        =   21
            Top             =   2520
            Width           =   3255
         End
      End
   End
   Begin MSComctlLib.TreeView tvwOptions 
      Height          =   5115
      Left            =   60
      TabIndex        =   42
      Top             =   60
      Width           =   2235
      _ExtentX        =   3942
      _ExtentY        =   9022
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   353
      LabelEdit       =   1
      Style           =   2
      Appearance      =   1
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private LastNode As Node

' --- events ---

Private Sub Form_Load()
    Dim i As Integer
    
    'size form
    Me.Height = 5700
    Me.Width = 7100
    
    Dim oNode As Node
    
    'populate treeview
    Set oNode = tvwOptions.Nodes.add(, tvwFirst, "root", "Options")
    tvwOptions.Nodes.add "root", tvwChild, "general", "General"
    tvwOptions.Nodes.add "root", tvwChild, "file", "File Operations"
    tvwOptions.Nodes.add "root", tvwChild, "display", "Display"
    tvwOptions.Nodes.add "root", tvwChild, "history", "History"
    oNode.Expanded = True 'show expanded
    'select first child
    oNode.Child.Selected = True
    Set oNode = Nothing
    
    Call HideAndSetFrames
    'tvwOptions.Nodes("general").Selected = True
    Call tvwOptions_NodeClick(tvwOptions.Nodes("general")) 'raise click event for first node
    
    'clear all comboboxes
    cboCategory.Clear
    cboDataFile.Clear
    cboDetailForm.Clear
    
    Call SetInitialFormOptions
    
    'get a list of categories
    Call GetCategories
    For i = 0 To UBound(sCategories)
        cboCategory.AddItem sCategories(i)
        cboLinkPathCategory.AddItem sCategories(i)
    Next i
    
    'put focus on right entry if found (general categories)
    With cboCategory
        If sLogCat <> "" Then
            For i = 0 To .ListCount - 1
                If .List(i) = sLogCat Then
                    .ListIndex = i
                    Exit For
                End If
            Next i
        Else
            .ListIndex = 0
        End If
    End With
    
    'put focus on right entry if found
    With cboCategory
        If sDataFileField <> "" Then
            For i = 0 To cboCategory.ListCount - 1
                If .List(i) = sDataFileField Then
                    .ListIndex = i
                    Exit For
                End If
            Next i
        End If
    End With

    'put focus on right entry if found
    With cboDetailForm
        If sDetailForm <> "" Then
            For i = 0 To .ListCount - 1
                If .List(i) = sDetailForm Then
                    .ListIndex = i
                    Exit For
                End If
            Next i
        End If
    End With
    
    'put focus on right entry if found (path category)
    With cboLinkPathCategory
        If sCopyCategory <> "" Then
            For i = 0 To .ListCount - 1
                If .List(i) = sCopyCategory Then
                    .ListIndex = i
                    Exit For
                End If
            Next i
        Else
            .ListIndex = 0
        End If
    End With
    
    'put focus on right entry if found
    With cboLinkPathFieldName
        If sCopyFieldName <> "" Then
            For i = 0 To .ListCount - 1
                If .List(i) = sCopyFieldName Then
                    .ListIndex = i
                    Exit For
                End If
            Next
        End If
    End With


End Sub



Private Sub tvwOptions_NodeClick(ByVal Node As MSComctlLib.Node)
    
    'set all frames to initial state
    Call HideAndSetFrames
    
    Select Case Node.Key
            
        Case "root"
            'prevent user from clicking node, return him to last node
            tvwOptions.Nodes(LastNode.Key).Selected = True
            Call tvwOptions_NodeClick(LastNode)
           
        Case "general"
            Set LastNode = Node
            Call ShowFrame(0)

        Case "file"
            Set LastNode = Node
            Call ShowFrame(1)
        
        Case "display"
            Set LastNode = Node
            Call ShowFrame(2)
        
        Case "history"
            Set LastNode = Node
            'show log in form
            rtfLog.Text = GetLogText
            Call ShowFrame(3)
            
    End Select

End Sub

Private Sub cmdCancel_Click()

    Unload Me
    
End Sub

Private Sub chkCopyFile_Click()
    Dim ctl As OptionButton
    
    If chkCopyFile.Value Then
        blnCopyFile = True
        chkDeleteAfterCopy.Enabled = True
        frmDatabaseCopyFilePath.Enabled = True
        cboLinkPathCategory.Enabled = True
        cboLinkPathFieldName.Enabled = True
        lblCopyPath.Enabled = True
        lblCategory.Enabled = True
        lblField.Enabled = True
        txtPath.Enabled = True
        cmdBrowseForFolder.Enabled = True
        frmCopyOptions.Enabled = True
        For Each ctl In optPathOrder
            ctl.Enabled = True
        Next ctl
    Else
        chkDeleteAfterCopy.Value = False
        blnCopyFile = False
        chkDeleteAfterCopy.Enabled = False
        frmDatabaseCopyFilePath.Enabled = False
        cboLinkPathCategory.Enabled = False
        cboLinkPathFieldName.Enabled = False
        lblCopyPath.Enabled = False
        lblCategory.Enabled = False
        lblField.Enabled = False
        txtPath.Enabled = False
        cmdBrowseForFolder.Enabled = False
        frmCopyOptions.Enabled = False
        For Each ctl In optPathOrder
            ctl.Enabled = False
        Next ctl
    End If
    
End Sub

Private Sub cmdBrowseForFolder_Click()
    Dim sPath As String
    
    sPath = BrowseForFolder(Me, "Select a directory to copy file(s) to:", CurDir())
    
    If sPath <> "" Then
        txtPath.Text = sPath
    End If
    
End Sub

Private Sub cmdOptions_Click()
    
    Load frmConflictingFile
    If frmOptions.chkAlwaysOnTop.Value Then Call SetTopMostWindow(frmConflictingFile.hwnd, True)
    frmConflictingFile.Show vbModal, Me
    
End Sub

Private Sub cmdOK_Click()
    Dim ctl As OptionButton
    Dim reply  As Integer
    
    'check if target category allows duplicates
    If Not DupsAllowed(cboCategory.Text) Then
        MsgBox "Category '" & cboCategory.Text & "' does not allow duplicates." _
            & vbCrLf & "Select another category to log to.", vbExclamation + vbOKOnly, app.EXEName
        Call ShowFrame(0)
        Exit Sub
    End If
    
    'see if pathname is valid
    If IsInvalidPathSyntax(txtPath.Text) Then
        reply = MsgBox("Pathname '" & txtPath.Text & "' contains invalid characters." _
                        & vbCrLf & "Are you sure you want to save the settings?", vbYesNo, "Syntax error")
        If reply = vbNo Then
            Call ShowFrame(1)
            Exit Sub
        End If
    End If
    
    'store values in ini file and close
    sWriteIni INI_SECTION_CMC, INI_KEY_CMC_CAT, Me.cboCategory.Text, sIniFile
    sWriteIni INI_SECTION_CMC, INI_KEY_CMC_NAMEFIELD, sFields(0), sIniFile
    sWriteIni INI_SECTION_CMC, INI_KEY_CMC_DATAFILE, Me.cboDataFile.Text, sIniFile
    sWriteIni INI_SECTION_CMC, INI_KEY_CMC_IDF, Me.cboDetailForm.Text, sIniFile
    sWriteIni INI_SECTION_OTHER, INI_KEY_OTHER_SHARE, Me.chkShared.Value, sIniFile
    sWriteIni INI_SECTION_OTHER, INI_KEY_OTHER_SHOWIDF, Me.chkShowDetailForm, sIniFile
    sWriteIni INI_SECTION_OTHER, INI_KEY_OTHER_SHOWCMC, Me.chkCommence, sIniFile
    sWriteIni INI_SECTION_OTHER, INI_KEY_OTHER_AUTOCON, Me.chkAutoConnect, sIniFile
    sWriteIni INI_SECTION_OTHER, INI_KEY_OTHER_DDE, Me.chkNoDDETrigger, sIniFile
    sWriteIni INI_SECTION_DISPLAY, INI_KEY_DISPLAY_CLOSE, Me.chkClose, sIniFile
    sWriteIni INI_SECTION_DISPLAY, INI_KEY_DISPLAY_ONTOP, Me.chkAlwaysOnTop, sIniFile
    'sWriteIni INI_SECTION_FILE, INI_KEY_FILE_SHCUT, Me.chkCreateShortcut, sIniFile
    sWriteIni INI_SECTION_FILE, INI_KEY_FILE_COPY, Me.chkCopyFile, sIniFile
    sWriteIni INI_SECTION_FILE, INI_KEY_FILE_DELETEORIGINAL, Me.chkDeleteAfterCopy, sIniFile
    sWriteIni INI_SECTION_FILE, INI_KEY_FILE_CAT, Me.cboLinkPathCategory.Text, sIniFile
    sWriteIni INI_SECTION_FILE, INI_KEY_FILE_FIELD, Me.cboLinkPathFieldName, sIniFile
    sWriteIni INI_SECTION_FILE, INI_KEY_FILE_CUSTOM, Me.txtPath.Text, sIniFile
    For Each ctl In optPathOrder
        If ctl.Value Then
            sWriteIni INI_SECTION_FILE, INI_KEY_FILE_ORDER, ctl.Index, sIniFile
            Exit For
        End If
    Next ctl
    
    'handling of file conflicts
    'this is done in form frmConflictingFile
    'if that was never loaded, set initial values here
    On Error Resume Next
    iWhatIfFileNewer = sReadIni(sIniFile, INI_SECTION_FILE, INI_KEY_FILE_NEWER)
    If Err.Number > 0 Then
        iWhatIfFileNewer = 0
        sWriteIni INI_SECTION_FILE, INI_KEY_FILE_NEWER, iWhatIfFileNewer, sIniFile
        Err.Clear
    End If
    iWhatIfAttrDifferent = sReadIni(sIniFile, INI_SECTION_FILE, INI_KEY_FILE_ATTR)
    If Err.Number > 0 Then
        iWhatIfAttrDifferent = 0
        sWriteIni INI_SECTION_FILE, INI_KEY_FILE_ATTR, iWhatIfAttrDifferent, sIniFile
        Err.Clear
    End If
    On Error GoTo 0
    
    Unload Me
    With frmMain
        .ReadOptions
        .RefreshCommenceItem
        If bNoAutoConnect Then .cmdOK.Enabled = True
        .ToggleWindowOnTopState
        .Show
        
    End With
    sLogMsg = "Option values written to " & sIniFile & "."
    Call AppendToLog(sLogFile, sLogMsg)
    
End Sub

Private Sub cmdViewLog_Click()
    Call APIShellExecute(Me.hwnd, sLogFile)
End Sub

Private Sub cboCategory_Click()
    'refresh stuff when another category was selected
    Dim i As Integer
    
    'clear comboboxes
    cboDataFile.Clear
    cboDetailForm.Clear
    sLogCat = cboCategory.Text  'get what category to log document to
    Call GetDataFields(sLogCat) 'retrieve fields for selected category
    For i = 1 To UBound(sFields)
        cboDataFile.AddItem sFields(i) 'add all fields of type data file or url to combobox
    Next i
    
    'retrieve detail form names
    Call GetDetailForms(sLogCat)
    For i = 0 To UBound(sFormNames)
        cboDetailForm.AddItem sFormNames(i)
    Next i
    
    'there is always a detail form
    cboDetailForm.ListIndex = 0
    
    'enable controls if URL or datafile field found
    If cboDataFile.ListCount > 0 Then
        cboDataFile.ListIndex = 0 'select first value in combobox
        cboDataFile.Enabled = True
        cmdOK.Enabled = True
    Else
        cboDataFile.Enabled = False
        cmdOK.Enabled = False
    End If
    
End Sub

Private Sub cboLinkPathCategory_Click()
    'refresh stuff when another category was selected
    Dim i As Integer

    'clear comboboxes
    cboLinkPathFieldName.Clear

    sCopyCategory = cboLinkPathCategory.Text
    Call GetTextEnabledFields(sCopyCategory)  'retrieve fields for selected category
    For i = 1 To UBound(sTextFields)
        cboLinkPathFieldName.AddItem sTextFields(i)  'add all fields of type data file or url to combobox
    Next i
    
    'enable controls if URL or datafile field found
    With cboLinkPathFieldName
        If .ListCount > 0 Then
            .ListIndex = 0 'select first value in combobox
            If blnCopyFile Then 'note additional check for copy file option
                .Enabled = True
                cmdOK.Enabled = True
            End If
        Else
            .Enabled = False
            cmdOK.Enabled = False
        End If
    End With
    
End Sub

Private Sub cmdCreateShortcut_Click()
    Dim strLnkPath As String 'special folder containing the .lnk file
    Dim LnkFile As String                               ' Link file name
    Dim ExeFile As String                               ' Link - Exe file name
    Dim WorkDir As String                               '      - Working directory
    Dim ExeArgs As String                               '      - Command line arguments
    Dim IconFile As String                              '      - Icon File name
    Dim IconIdx As Long                                 '      - Icon Index
    Dim ShowCmd As Long                                 '      - Program start state...
    Dim oLnk As New cShellLink

    LnkFile = ""
    ExeFile = ""
    WorkDir = ""
    ExeArgs = ""
    IconFile = ""
    IconIdx = 0
    ShowCmd = 0
    
    If Not oLnk.GetSystemFolderPath(Me.hwnd, CSIDL_SENDTO, strLnkPath) Then
        MsgBox "Unable to retrieve SendTo path", vbCritical, app.EXEName
        GoTo err_handler
    End If
    LnkFile = strLnkPath & "\" & CMC & ".lnk"
    oLnk.GetShellLinkInfo LnkFile, ExeFile, WorkDir, ExeArgs, IconFile, IconIdx, ShowCmd
    If Not ExeFile = "" Then
        MsgBox "Shortcut already exists.", vbInformation, app.EXEName
    Else
        ExeFile = app.Path & "\" & app.EXEName & ".exe"
        WorkDir = app.Path
        ExeArgs = ""
        IconFile = ""
        IconIdx = 0
        ShowCmd = 0
        If Not oLnk.CreateShellLink(LnkFile, ExeFile, WorkDir, ExeArgs, IconFile, IconIdx, ShowCmd) Then
            MsgBox "Unable to create shortcut", vbCritical, app.EXEName
            GoTo err_handler
        Else
            MsgBox "Shortcut was successfully added to the SendTo menu.", vbInformation + vbOKOnly, app.EXEName
        End If
    End If

err_handler:
    Set oLnk = Nothing

End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    Set LastNode = Nothing

End Sub

Private Sub cmdPurgeLog_Click()
    Dim reply As Integer
    reply = MsgBox("Are you sure you want to purge the log file?", vbYesNo + vbQuestion, app.EXEName)
    If reply = vbYes Then
        Call PurgeLog
        rtfLog.Text = GetLogText
    End If
End Sub

' --- end events ---

' --- Helper routines ---

Private Sub HideAndSetFrames()
    Dim f As Frame
    'set position and hide frames
    For Each f In Frame
        f.Top = 0
        f.Left = 2350
        f.Visible = False
    Next f

End Sub

Private Sub ShowFrame(ByVal i As Integer)

    Dim f As Frame
    'hide frames
    For Each f In Frame
        f.Visible = False
    Next f
    'show frame
    Frame(i).Visible = True
    
End Sub
Private Function DupsAllowed(ByVal strCat As String) As Boolean

    Dim i As Integer
    Dim buffer() As String
    
    Set oConv = oCmc.GetConversation(CMC, "GetData")
    'check if duplicates are allowed, they should be
    strDDE = "[GetCategoryDefinition(" & dq(strCat) & "," & CMC_DELIM & ")]"
    buffer = Split(oConv.Request(strDDE), CMC_DELIM)
    If CStr(Mid(buffer(1), 9, 1)) = "1" Then DupsAllowed = True
    Set oConv = Nothing
    
End Function

Private Sub SetInitialFormOptions()

    sLogCat = sReadIni(sIniFile, INI_SECTION_CMC, INI_KEY_CMC_CAT)
    sNameField = sReadIni(sIniFile, INI_SECTION_CMC, INI_KEY_CMC_NAMEFIELD)
    sDataFileField = sReadIni(sIniFile, INI_SECTION_CMC, INI_KEY_CMC_DATAFILE)
    sDetailForm = sReadIni(sIniFile, INI_SECTION_CMC, INI_KEY_CMC_IDF)

    On Error Resume Next
    chkShared.Value = sReadIni(sIniFile, INI_SECTION_OTHER, INI_KEY_OTHER_SHARE)
    chkShowDetailForm.Value = sReadIni(sIniFile, INI_SECTION_OTHER, INI_KEY_OTHER_SHOWIDF)
    chkCommence.Value = sReadIni(sIniFile, INI_SECTION_OTHER, INI_KEY_OTHER_SHOWCMC)
    chkAutoConnect.Value = sReadIni(sIniFile, INI_SECTION_OTHER, INI_KEY_OTHER_AUTOCON)
    chkNoDDETrigger.Value = sReadIni(sIniFile, INI_SECTION_OTHER, INI_KEY_OTHER_DDE)
    chkClose.Value = sReadIni(sIniFile, INI_SECTION_DISPLAY, INI_KEY_DISPLAY_CLOSE)
    chkAlwaysOnTop.Value = sReadIni(sIniFile, INI_SECTION_DISPLAY, INI_KEY_DISPLAY_ONTOP)
    'chkCreateShortcut.Value = sReadIni(sIniFile, INI_SECTION_FILE, INI_KEY_FILE_SHCUT)
    chkCopyFile.Value = sReadIni(sIniFile, INI_SECTION_FILE, INI_KEY_FILE_COPY)
    chkCopyFile.Value = blnCopyFile
    chkDeleteAfterCopy = sReadIni(sIniFile, INI_SECTION_FILE, INI_KEY_FILE_DELETEORIGINAL)
    chkDeleteAfterCopy = blnDeleteAfterCopy
    sCopyCategory = sReadIni(sIniFile, INI_SECTION_FILE, INI_KEY_FILE_CAT)
    sCopyFieldName = sReadIni(sIniFile, INI_SECTION_FILE, INI_KEY_FILE_FIELD)
    sCopyPath = sReadIni(sIniFile, INI_SECTION_FILE, INI_KEY_FILE_CUSTOM)
    txtPath.Text = sCopyPath
    iPrecedence = sReadIni(sIniFile, INI_SECTION_FILE, INI_KEY_FILE_ORDER)
    optPathOrder(iPrecedence).Value = True 'set option button value
    
End Sub
