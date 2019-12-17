Attribute VB_Name = "mShellLink"
Option Explicit

'---------------------------------------------------------------
'- Public API Declares...
'---------------------------------------------------------------
#If UNICODE Then
    Public Declare Function SHGetPathFromIDList Lib "shell32" Alias "SHGetPathFromIDListW" (ByVal pidl As Long, ByVal szPath As Long) As Long
#Else
    Public Declare Function SHGetPathFromIDList Lib "shell32" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal szPath As String) As Long
#End If

Public Declare Function SHGetSpecialFolderLocation Lib "shell32" (ByVal hwndOwner As Long, ByVal nFolder As Integer, ppidl As Long) As Long

'---------------------------------------------------------------
'- Public constants...
'---------------------------------------------------------------
Public Const MAX_PATH = 255
Public Const MAX_NAME = 40
