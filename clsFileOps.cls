VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFileOps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Module      : modShellFileDisk
' Description : Routines for working with the Windows 95/NT 4.0 shell
' Source      : Total VB SourceBook 6
'

Option Explicit


Private Type BrowseInfo
    hWndOwner      As Long
    pIDLRoot       As Long
    pszDisplayName As Long
    lpszTitle      As Long
    ulFlags        As Long
    lpfnCallback   As Long
    lParam         As Long
    iImage         As Long
End Type

Private Const BIF_RETURNONLYFSDIRS = 1
Private Const MAX_PATH = 260

Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long




Private Const CSIDL_DESKTOP = &H0
Private Const CSIDL_PROGRAMS = &H2
Private Const CSIDL_CONTROLS = &H3
Private Const CSIDL_PRINTERS = &H4
Private Const CSIDL_PERSONAL = &H5
Private Const CSIDL_FAVORITES = &H6
Private Const CSIDL_STARTUP = &H7
Private Const CSIDL_RECENT = &H8
Private Const CSIDL_SENDTO = &H9
Private Const CSIDL_BITBUCKET = &HA
Private Const CSIDL_STARTMENU = &HB
Private Const CSIDL_DESKTOPDIRECTORY = &H10
Private Const CSIDL_DRIVES = &H11
Private Const CSIDL_NETWORK = &H12
Private Const CSIDL_NETHOOD = &H13
Private Const CSIDL_FONTS = &H14
Private Const CSIDL_TEMPLATES = &H15
Private Const CSIDL_COMMON_STARTMENU = &H16
Private Const CSIDL_COMMON_PROGRAMS = &H17
Private Const CSIDL_COMMON_STARTUP = &H18
Private Const CSIDL_COMMON_DESKTOPDIRECTORY = &H19
Private Const CSIDL_APPDATA = &H1A
Private Const CSIDL_PRINTHOOD = &H1B

'Private Const BIF_RETURNONLYFSDIRS = &H1


Private Type SHITEMID
  cb As Long
  abID As Byte
End Type

Private Type ITEMIDLIST
  mkid As SHITEMID
End Type

Private Type SHFILEOPSTRUCT
  hWnd As Long
  wFunc As Long
  pFrom As String
  pTo As String
  fFlags As Integer
  fAnyOperationsAborted As Long
  hNameMappings As Long
  lpszProgressTitle As String
End Type

'Private Declare Function SHGetPathFromIDList '  Lib "shell32.dll" '  Alias "SHGetPathFromIDListA" '  (ByVal pidl As Long, '   ByVal pszPath As String) '  As Long

Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hWndOwner As Long, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long

Private Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long


Private Declare Sub SHAddToRecentDocs Lib "shell32.dll" (ByVal uFlags As Long, ByVal pszPath As String)

Private Declare Function SHFormatDrive Lib "shell32" (ByVal hWnd As Long, ByVal Drive As Long, ByVal fmtID As Long, ByVal options As Long) As Long
   
Private Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long

'Private Declare Sub CoTaskMemFree '  Lib "ole32.dll" '  (ByVal pv As Long)

Private Const FO_MOVE = &H1
Private Const FO_COPY = &H2
Private Const FO_DELETE = &H3
Private Const FO_RENAME = &H4

Private Const FOF_MULTIDESTFILES = &H1
Private Const FOF_CONFIRMMOUSE = &H2
Private Const FOF_SILENT = &H4
Private Const FOF_RENAMEONCOLLISION = &H8
Private Const FOF_NOCONFIRMATION = &H10
Private Const FOF_WANTMAPPINGHANDLE = &H20

Private Const FOF_ALLOWUNDO = &H40
Private Const FOF_FILESONLY = &H80
Private Const FOF_SIMPLEPROGRESS = &H100
Private Const FOF_NOCONFIRMMKDIR = &H200
Private Const FOF_NOERRORUI = &H400
Private Const SHARD_PATH = &H2&

' GetDriveType return values
Private Const DRIVE_NO_ROOT_DIR = 1
Private Const DRIVE_REMOVABLE = 2
Private Const DRIVE_FIXED = 3
Private Const DRIVE_REMOTE = 4
Private Const DRIVE_CDROM = 5
Private Const DRIVE_RAMDISK = 6

Private Const SHFMT_OPT_FULL = &H1
Private Const SHFMT_OPT_SYSONLY = &H2

Public Enum eSplitPath

    eSplitPathGetDriveLetter = 0
    eSplitPathGetFileName = 1
    eSplitPathGetExtension = 2
    eSplitPathGetFullPath = 3
    
End Enum



'__________________________________________________
' Scope  : Public
' Type   : Sub
' Name   : ShellCopyFile
' NOtes  : Copies a file or files to a single destination
' Params : lnghWnd - handle to window to serve as
'             the parent for the dialog. Use a form's
'             hWnd property for example
'             strSource - file spec for files to copy
'             strDestination - destination file name or directory
'             fSilent - if true, no warnings are displayed
'             strTitle - title of the progress dialog
' Returns: Nothing
' Desc   : The Sub uses parameters lnghWnd As Long, ByVal strSource As String, ByVal strDestination As String, Optional ByVal fSilent As Boolean = False and Optional strTitle As String = "Copying..." for ShellCopyFile and returns Nothing.
'__________________________________________________
' History
' CDK: 20010102: Added Error Trapping & Comments
'__________________________________________________
Public Sub ShellCopyFile(lnghWnd As Long, ByVal strSource As String, ByVal strDestination As String, Optional ByVal fSilent As Boolean = False, Optional strTitle As String = "Copying...")
  
  Dim foCopy As SHFILEOPSTRUCT
  Dim lngFlags As Long
  Dim lngResult As Long
    
  On Error GoTo Proc_Err
  
  ' check to be sure file exists
  If Dir$(strSource) <> "" Then
    
    ' set flags for no prompting
    If fSilent Then
      lngFlags = FOF_NOCONFIRMMKDIR Or FOF_NOCONFIRMATION Or FOF_SILENT
    End If
    
    ' set shell file operations settings
    With foCopy
      .hWnd = lnghWnd
      .pFrom = strSource
      .pTo = strDestination
      .fFlags = lngFlags
      .lpszProgressTitle = strTitle
      .wFunc = FO_COPY
      
      lngResult = SHFileOperation(foCopy)
    
    End With
    
  End If

Proc_Exit:
  Exit Sub

Proc_Err:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "ShellCopyFile"
  Resume Proc_Exit

End Sub


'__________________________________________________
' Scope  : Public
' Type   : Sub
' Name   : ShellRenameFile
' Params : lnghWnd - handle to window to serve as
'             the parent for the dialog. Use a form's
'             hWnd property for example
'             sOldName - old file name
'             sNewName - new file name
'             fSilent - if true, no warnings are displayed
'             strTitle - title of the progress dialog
' Returns: Nothing
' Desc   : The Sub uses parameters  for ShellRenameFile and returns _.
'__________________________________________________
' History
' CDK: 20010102: Added Error Trapping & Comments
'__________________________________________________
Public Sub ShellRenameFile( _
  lnghWnd As Long, _
  ByVal sOldName As String, _
  ByVal sNewName As String, _
  Optional ByVal fSilent As Boolean = False, _
  Optional strTitle As String = "Copying...")
  ' Comments  : Copies a file or files to a single destination
  ' Returns   : Nothing
  ' Source    : Total VB SourceBook 6
  '
  Dim foRename As SHFILEOPSTRUCT
  Dim lngFlags As Long
  Dim lngResult As Long
    
  On Error GoTo Proc_Err
  
  ' check to be sure file exists
  If Dir$(sOldName) <> "" Then
    
    ' set flags for no prompting
    If fSilent Then
      lngFlags = FOF_NOCONFIRMMKDIR Or FOF_NOCONFIRMATION Or FOF_SILENT
    End If
    
    ' set shell file operations settings
    With foRename
      .hWnd = lnghWnd
      .pFrom = sOldName
      .pTo = sNewName
      .fFlags = lngFlags
      .lpszProgressTitle = strTitle
      .wFunc = FO_RENAME
      
      lngResult = SHFileOperation(foRename)
    
    End With
    
  End If

Proc_Exit:
  Exit Sub

Proc_Err:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "ShellCopyFile"
  Resume Proc_Exit

End Sub





'__________________________________________________
' Scope  : Public
' Type   : Function
' Name   : BrowseForFolderPf
' Params :
'          hWndOwner As Long
'          sPrompt As String
' Returns: String
' Desc   : The Function uses parameters hWndOwner As Long and sPrompt As String for BrowseForFolderPf and returns String.
'__________________________________________________
' History
' CDK: 20010102: Added Error Trapping & Comments
'__________________________________________________
Public Function BrowseForFolderPf(hWndOwner As Long, sPrompt As String) As String
    On Error GoTo Proc_Err
    Const csProcName As String = "BrowseForFolderPf"


    Dim iNull As Integer
    Dim lpIDList As Long
    Dim lResult As Long
    Dim sPath As String
    Dim udtBI As BrowseInfo

    With udtBI
        .hWndOwner = hWndOwner
        .lpszTitle = lstrcat(sPrompt, "")
        .ulFlags = BIF_RETURNONLYFSDIRS
    End With

    lpIDList = SHBrowseForFolder(udtBI)
    If lpIDList Then
        sPath = String$(MAX_PATH, 0)
        lResult = SHGetPathFromIDList(lpIDList, sPath)
        Call CoTaskMemFree(lpIDList)
        iNull = InStr(sPath, vbNullChar)
        If iNull Then
            sPath = Left$(sPath, iNull - 1)
        End If
    End If

    BrowseForFolderPf = sPath


Proc_Exit:
    GoSub Proc_Cleanup
    Exit Function

Proc_Cleanup:
    On Error Resume Next
    'Place any cleanup of instantiated objects here
    
    On Error GoTo 0
    Return

Proc_Err:
    Dim lErrNum As String, sErrSource As String, sErrDesc As String
    lErrNum = VBA.Err.Number
    sErrSource = VBA.Err.Source & vbCrLf & "clsFileOps->" & csProcName
    sErrDesc = VBA.Err.Description
    Resume Proc_Err_Continue
    
Proc_Err_Continue:
    GoSub Proc_Cleanup
    Err.Raise Number:=lErrNum, Source:=sErrSource, Description:=sErrDesc
    Exit Function
 
End Function








'__________________________________________________
' Scope  : Public
' Type   : Function
' Name   : FilesToArray
' Desc   : Populates the passed array with a list of files in the
'             specified directory
' Params : sDirectory - Directory to read
'             fIncludeHidden - True to include hidden files
'             fIncludeSystem - True to include system files
'             asArray() - Array to hold file names (0-based)
' Returns: Number of files in the directory' Desc   : The Function uses parameters  for FilesToArray and returns _.
'__________________________________________________
' History
' CDK: 20010102: Added Error Trapping & Comments
'__________________________________________________
Public Function FilesToArray( _
  sDirectory As String, _
  fIncludeHidden As Boolean, _
  fIncludeSystem As Boolean, _
  asArray() As String) _
  As Integer
  
  
  Dim intCounter As Integer
  Dim strTmp As String
  Dim lngAttr As Long
  
  On Error GoTo Proc_Err

  ' Build up the options flag
  lngAttr = VBA.vbNormal
  
  If fIncludeHidden Then
    lngAttr = lngAttr + VBA.vbHidden
  End If
  
  If fIncludeSystem Then
    lngAttr = lngAttr + VBA.vbSystem
  End If
  
  ' Make sure there is a trailing slash
  If Right$(sDirectory, 1) <> "\" Then
    sDirectory = sDirectory & "\"
  End If
  
  ' Get the first entry
  strTmp = Dir(sDirectory, lngAttr)
  
  ' Loop through each entry
  Do Until strTmp = ""

    ' Are there entries left?
    If strTmp <> "" Then
      ' Grow the array and add the item
      ReDim Preserve asArray(intCounter)
      asArray(intCounter) = strTmp
      intCounter = intCounter + 1
    End If
    strTmp = Dir
  Loop

  ' Return the count
  FilesToArray = intCounter

Proc_Exit:
  Exit Function
  
Proc_Err:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "FilesToArray"
  Resume Proc_Exit

End Function






'__________________________________________________
' Scope  : Public
' Type   : Function
' Name   : PathFromFullPathPs
' Params :
'          ByVal strPath As String - path minus the file name if there is one
' Returns: String
' Desc   : The Function uses parameters ByVal strPath As String for PathFromFullPathPs and returns String.
'__________________________________________________
' History
' CDK: 20010102: Added Error Trapping & Comments
'__________________________________________________
Public Function PathFromFullPathPs(ByVal strPath As String) As String
  
  Dim intPos As Integer
  Dim strTmp As String
  
  On Error GoTo Proc_Err
  
  ' Initialize return value
  strTmp = ""
  
    'remove all after the last "\"
    intPos = InStrRev(strPath, "\")
    'strPath = Right(strPath, Len(strPath) - intPos)
    strTmp = Left(strPath, intPos - 1)
        
Proc_Exit:
  PathFromFullPathPs = strTmp
  Exit Function
  
Proc_Err:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "PathFromFullPathPs"
  Resume Proc_Exit
  
End Function



                        



'__________________________________________________
' Scope  : Public
' Type   : Function
' Name   : SplitPathPs
' Params :
'          ByVal sFileNameWithPath As String
'          eReturnType As eSplitPath
' Returns: String
' Desc   : The Function uses parameters ByVal sFileNameWithPath As String and eReturnType As eSplitPath for SplitPathPs and returns String.
'__________________________________________________
' History
' CDK: 20010102: Added Error Trapping & Comments
'__________________________________________________
Public Function SplitPathPs(ByVal sFileNameWithPath As String, eReturnType As eSplitPath) As String
    On Error GoTo Proc_Err
    Const csProcName As String = "SplitPathPs"

    
    Dim sDriveLetter As String
    Dim sDirPath As String
    Dim sFileName As String
    Dim sExtension As String
    Dim lPathLength As Long
    Dim lOffset As Long
    Dim lThisLength As Long
    Dim fFileNameFound As Boolean
    
    sDriveLetter = ""
    sDirPath = ""
    sFileName = ""
    sExtension = ""


    If Mid(sFileNameWithPath, 2, 1) = ":" Then ' Find the drive letter.
        sDriveLetter = Left(sFileNameWithPath, 2)
        sFileNameWithPath = Mid(sFileNameWithPath, 3)
    End If
    lPathLength = Len(sFileNameWithPath)


    For lOffset = lPathLength To 1 Step -1 ' Find the Next delimiter.


        Select Case Mid(sFileNameWithPath, lOffset, 1)
            
            Case "."
                ' This indicates either an sExtension or a . or a ..
                lThisLength = Len(sFileNameWithPath) - lOffset
                If lThisLength >= 1 And lThisLength <= 3 Then ' sExtension
                    sExtension = Mid(sFileNameWithPath, lOffset, lThisLength + 1)
                End If
                sFileNameWithPath = Left(sFileNameWithPath, lOffset - 1)
            
            Case "\"
                ' This indicates a path delimiter.
                lThisLength = Len(sFileNameWithPath) - lOffset
                If lThisLength >= 1 And lThisLength <= 40 Then ' sFileName
                    sFileName = Mid(sFileNameWithPath, lOffset + 1, lThisLength)
                    sDirPath = Left(sFileNameWithPath, lOffset)
                    fFileNameFound = True
                    Exit For
                End If
            
        End Select
        
    Next lOffset

    
    Select Case eReturnType
        Case eSplitPath.eSplitPathGetFileName
            SplitPathPs = sFileName & sExtension
        
        Case eSplitPath.eSplitPathGetFullPath
            SplitPathPs = sDriveLetter & sDirPath
        
        Case eReturnType = eSplitPath.eSplitPathGetExtension
            SplitPathPs = sExtension
        
        Case eReturnType = eSplitPath.eSplitPathGetDriveLetter
            SplitPathPs = sDriveLetter
    
    End Select


Proc_Exit:
    GoSub Proc_Cleanup
    Exit Function

Proc_Cleanup:
    On Error Resume Next
    'Place any cleanup of instantiated objects here
    
    On Error GoTo 0
    Return

Proc_Err:
    Dim lErrNum As String, sErrSource As String, sErrDesc As String
    lErrNum = VBA.Err.Number
    sErrSource = VBA.Err.Source & vbCrLf & "clsFileOps->" & csProcName
    sErrDesc = VBA.Err.Description
    Resume Proc_Err_Continue
    
Proc_Err_Continue:
    GoSub Proc_Cleanup
    Err.Raise Number:=lErrNum, Source:=sErrSource, Description:=sErrDesc
    Exit Function
 
End Function


