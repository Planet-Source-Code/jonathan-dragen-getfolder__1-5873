<div align="center">

## GetFolder


</div>

### Description

It displays a folder chooser using API.

This API call can't be found with the API Viewer.
 
### More Info
 
Title, hWnd

Paste this into a module.

The folder path


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Jonathan Dragen](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/jonathan-dragen.md)
**Level**          |Intermediate
**User Rating**    |4.5 (18 globes from 4 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Windows API Call/ Explanation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-api-call-explanation__1-39.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/jonathan-dragen-getfolder__1-5873/archive/master.zip)

### API Declarations

```
Option Explicit
Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long
Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
'These constants are to be set to the ulFlags property in the BROWSEINFO type depending of what result you want
Const BIF_RETURNONLYFSDIRS = &H1  'Allows you to browse for system folders only.
Const BIF_DONTGOBELOWDOMAIN = &H2  'Using this value forces the _
                   user to stay within the domain level of the _
                   Network Neighborhhood
Const BIF_STATUSTEXT = &H4     'Displays a statusbar on the selection dialog
Const BIF_RETURNFSANCESTORS = &H8  'Returns file system ancestor only
Const BIF_BROWSEFORCOMPUTER = &H1000 'Allows you to browse for a computer
Const BIF_BROWSEFORPRINTER = &H2000 'Allows you to browse the Printers folder
Type BROWSEINFO
  hOwner As Long
  pidlRoot As Long
  pszDisplayName As String
  lpszTitle As String
  ulFlags As Long
  lpfn As Long
  lParam As Long
  iImage As Long
End Type
```


### Source Code

```
Sub Main()
Dim Folder As String
Folder = GetFolder
If Not Folder = "" Then
  MsgBox Folder
Else
  MsgBox "Couldn't find folder."
End If
End Sub
Function GetFolder(Optional Title As String, Optional hWnd) As String
Dim bi As BROWSEINFO
Dim pidl As Long
Dim Folder As String
Folder = String$(255, Chr$(0))
With bi
  If IsNumeric(hWnd) Then .hOwner = hWnd
  .ulFlags = BIF_RETURNONLYFSDIRS
  .pidlRoot = 0
  If Not IsMissing(Title) Then
    .lpszTitle = Title
  Else
    .lpszTitle = "Select a Folder" & Chr$(0)
  End If
End With
pidl = SHBrowseForFolder(bi)
If SHGetPathFromIDList(ByVal pidl, ByVal Folder) Then
  GetFolder = Left(Folder, InStr(Folder, Chr$(0)) - 1)
Else
  GetFolder = ""
End If
End Function
```

