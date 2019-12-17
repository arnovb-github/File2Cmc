VERSION 5.00
Begin VB.Form frmConflictingFile 
   Caption         =   "Conflicting file options"
   ClientHeight    =   3840
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4215
   Icon            =   "frmConflictingFile.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3840
   ScaleWidth      =   4215
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3240
      TabIndex        =   9
      Top             =   3360
      Width           =   855
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2280
      TabIndex        =   8
      Top             =   3360
      Width           =   855
   End
   Begin VB.Frame Frame2 
      Caption         =   "Files are the same but have different file attributes"
      Height          =   1215
      Left            =   120
      TabIndex        =   1
      Top             =   2040
      Width           =   3975
      Begin VB.OptionButton optAttributesDifferent 
         Caption         =   "Overwrite target file (Recommended)"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   6
         Top             =   600
         Value           =   -1  'True
         Width           =   3495
      End
      Begin VB.OptionButton optAttributesDifferent 
         Caption         =   "Do not overwrite target file"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   7
         Top             =   840
         Width           =   3495
      End
      Begin VB.OptionButton optAttributesDifferent 
         Caption         =   "Prompt me"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   3495
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Target file already exists and is newer"
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   3975
      Begin VB.OptionButton optTargetNewer 
         Caption         =   "Do not overwrite target file"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   4
         Top             =   840
         Width           =   3495
      End
      Begin VB.OptionButton optTargetNewer 
         Caption         =   "Overwrite target file"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   3
         Top             =   600
         Width           =   3495
      End
      Begin VB.OptionButton optTargetNewer 
         Caption         =   "Prompt me (Recommended)"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Value           =   -1  'True
         Width           =   3495
      End
   End
   Begin VB.Label Label1 
      Caption         =   "In case file to copy already exists in target location"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   4575
   End
End
Attribute VB_Name = "frmConflictingFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()

    Unload Me
    
End Sub

Private Sub cmdOK_Click()

    Dim ctl As OptionButton
    
    For Each ctl In optTargetNewer
        If ctl.Value Then
            Call sWriteIni(INI_SECTION_FILE, INI_KEY_FILE_NEWER, ctl.Index, sIniFile)
            iWhatIfFileNewer = ctl.Index
            Exit For
        End If
    Next ctl
    
    For Each ctl In optAttributesDifferent
        If ctl.Value Then
            Call sWriteIni(INI_SECTION_FILE, INI_KEY_FILE_ATTR, ctl.Index, sIniFile)
            iWhatIfAttrDifferent = ctl.Index
            Exit For
        End If
    Next ctl
    
    Unload Me
    
End Sub

Private Sub Form_Load()
    Dim i As Integer
    
    On Error Resume Next
    i = sReadIni(sIniFile, INI_SECTION_FILE, INI_KEY_FILE_NEWER)
    optTargetNewer.item(i).Value = True
    i = sReadIni(sIniFile, INI_SECTION_FILE, INI_KEY_FILE_ATTR)
    optAttributesDifferent.item(i).Value = True
    
End Sub
