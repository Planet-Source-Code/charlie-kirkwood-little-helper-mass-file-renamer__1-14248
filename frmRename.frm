VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRenamer 
   Caption         =   "Little Helper: Mass File Renaming Tool"
   ClientHeight    =   7590
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6180
   LinkTopic       =   "Form1"
   ScaleHeight     =   7590
   ScaleWidth      =   6180
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkMatchCase 
      Alignment       =   1  'Right Justify
      Caption         =   "Matc&h Case"
      Height          =   255
      Left            =   45
      TabIndex        =   16
      Top             =   1305
      Width           =   2460
   End
   Begin VB.CheckBox chkMarkAllAs 
      Alignment       =   1  'Right Justify
      Caption         =   "Mark all as:"
      Height          =   270
      Left            =   75
      TabIndex        =   15
      Top             =   6030
      Value           =   1  'Checked
      Width           =   1155
   End
   Begin VB.ListBox lstFilesToRename 
      Enabled         =   0   'False
      Height          =   4335
      Left            =   75
      Style           =   1  'Checkbox
      TabIndex        =   14
      Top             =   1620
      Width           =   6015
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   330
      Left            =   5235
      TabIndex        =   12
      Top             =   7155
      Width           =   885
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "&Clear All"
      Height          =   330
      Left            =   75
      TabIndex        =   11
      Top             =   7155
      Width           =   885
   End
   Begin VB.CommandButton cmdBrowseForString 
      Caption         =   "Br&owse"
      Height          =   330
      Left            =   5190
      TabIndex        =   2
      Top             =   150
      Width           =   885
   End
   Begin VB.CommandButton cmdBrowseForFolder 
      Caption         =   "&Browse"
      Height          =   330
      Left            =   5190
      TabIndex        =   7
      Top             =   915
      Width           =   885
   End
   Begin VB.TextBox txtFolder 
      Height          =   330
      Left            =   2310
      TabIndex        =   6
      Top             =   915
      Width           =   2835
   End
   Begin VB.TextBox txtReplace 
      Height          =   330
      Left            =   2310
      TabIndex        =   4
      Top             =   540
      Width           =   2835
   End
   Begin VB.TextBox txtSearch 
      Height          =   330
      Left            =   2310
      TabIndex        =   1
      Top             =   150
      Width           =   2835
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   600
      Top             =   7020
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame fraProgress 
      Caption         =   "Files Processed: "
      Height          =   720
      Left            =   75
      TabIndex        =   9
      Top             =   6360
      Visible         =   0   'False
      Width           =   6030
      Begin MSComctlLib.ProgressBar prgProgress 
         Height          =   300
         Left            =   165
         TabIndex        =   10
         Top             =   270
         Width           =   5760
         _ExtentX        =   10160
         _ExtentY        =   529
         _Version        =   393216
         Appearance      =   1
      End
   End
   Begin VB.CommandButton cmdRename 
      Caption         =   "&Replace"
      Height          =   330
      Left            =   4290
      TabIndex        =   8
      Top             =   7155
      Width           =   885
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Ca&ncel"
      Height          =   330
      Left            =   4290
      TabIndex        =   13
      Top             =   7155
      Width           =   885
   End
   Begin VB.Label Label3 
      Caption         =   "&In Folder:"
      Height          =   225
      Left            =   75
      TabIndex        =   5
      Top             =   960
      Width           =   2235
   End
   Begin VB.Label Label2 
      Caption         =   "Replace Search String &With:"
      Height          =   210
      Left            =   75
      TabIndex        =   3
      Top             =   615
      Width           =   2205
   End
   Begin VB.Label Label1 
      Caption         =   "&Search String:"
      Height          =   210
      Left            =   75
      TabIndex        =   0
      Top             =   210
      Width           =   1215
   End
End
Attribute VB_Name = "frmRenamer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'*************************************************************************
' Author    : Charlie Kirkwood
' File      : frmCodeCommenter.frm
' NOTE:     : This program is will find files that match a search string. and
'               allow users to change file names in bulk. This is very useful
'               for fixing files names that are consistently wrong or bothersom.
'               I used it when downloading MP3's and want to fix the file names -
'               for example to remove all underscores.
'
'*************************************************************************
' History   : 200011xx - CDK - Created
'
'*************************************************************************


Private moFileOps As clsFileOps
Private mfStopProcessing As Boolean




'__________________________________________________
' Scope  : Private
' Type   : Sub
' Name   : chkMarkAllAs_Click
' Params :
' Returns: Nothing
' Desc   : The Sub uses parameters  for chkMarkAllAs_Click and returns Nothing.
'__________________________________________________
' History
' CDK: 20010102: Added Error Trapping & Comments
'__________________________________________________
Private Sub chkMarkAllAs_Click()
    On Error GoTo Proc_Err
    Const csProcName As String = "chkMarkAllAs_Click"


    Dim lCount As Long
    Dim lTotal As Long
    
    lTotal = Me.lstFilesToRename.ListCount - 1
    
    'check or uncheck all
    For lCount = 0 To lTotal
        Me.lstFilesToRename.Selected(lCount) = Me.chkMarkAllAs.Value
    Next



Proc_Exit:
    GoSub Proc_Cleanup
    Exit Sub

Proc_Cleanup:
    On Error Resume Next
    'Place any cleanup of instantiated objects here
        

    On Error GoTo 0
    Return

Proc_Err:
    Dim lErrNum As String, sErrSource As String, sErrDesc As String
    lErrNum = VBA.Err.Number
    sErrSource = VBA.Err.Source & vbCrLf & "frmRenamer->" & csProcName
    sErrDesc = VBA.Err.Description
    Resume Proc_Err_Continue
    Resume
    
Proc_Err_Continue:
    GoSub Proc_Cleanup
    MsgBox "The following Error Occurred" & vbCrLf & vbCrLf & _
                 "Source:" & vbCrLf & sErrSource & vbCrLf & vbCrLf & _
                 "Number:" & vbCrLf & lErrNum & vbCrLf & vbCrLf & _
                 "Description:" & vbCrLf & sErrDesc, vbExclamation + vbOKOnly

End Sub


'__________________________________________________
' Scope  : Private
' Type   : Sub
' Name   : chkMatchCase_Click
' Params :
' Returns: Nothing
' Desc   : The Sub uses parameters  for chkMatchCase_Click and returns Nothing.
'__________________________________________________
' History
' CDK: 20010102: Added Error Trapping & Comments
'__________________________________________________
Private Sub chkMatchCase_Click()
    On Error GoTo Proc_Err
    Const csProcName As String = "chkMatchCase_Click"

      SelectFilesForRenameM

Proc_Exit:
    GoSub Proc_Cleanup
    Exit Sub

Proc_Cleanup:
    On Error Resume Next
    'Place any cleanup of instantiated objects here
    
    On Error GoTo 0
    Return

Proc_Err:
    Dim lErrNum As String, sErrSource As String, sErrDesc As String
    lErrNum = VBA.Err.Number
    sErrSource = VBA.Err.Source & vbCrLf & "frmRenamer->" & csProcName
    sErrDesc = VBA.Err.Description
    Resume Proc_Err_Continue
    Resume
    
Proc_Err_Continue:
    GoSub Proc_Cleanup
    MsgBox "The following Error Occurred" & vbCrLf & vbCrLf & _
                 "Source:" & vbCrLf & sErrSource & vbCrLf & vbCrLf & _
                 "Number:" & vbCrLf & lErrNum & vbCrLf & vbCrLf & _
                 "Description:" & vbCrLf & sErrDesc, vbExclamation + vbOKOnly

End Sub


'__________________________________________________
' Scope  : Private
' Type   : Sub
' Name   : cmdBrowseForFolder_Click
' Params :
' Returns: Nothing
' Desc   : The Sub uses parameters  for cmdBrowseForFolder_Click and returns Nothing.
'__________________________________________________
' History
' CDK: 20010102: Added Error Trapping & Comments
'__________________________________________________
Private Sub cmdBrowseForFolder_Click()
    On Error GoTo Proc_Err
    Const csProcName As String = "cmdBrowseForFolder_Click"


    Me.txtFolder = moFileOps.BrowseForFolderPf(Me.hWnd, "select a folder")
    'fill the list box with all files in the selected dir
    SelectFilesForRenameM



Proc_Exit:
    GoSub Proc_Cleanup
    Exit Sub

Proc_Cleanup:
    On Error Resume Next
    'Place any cleanup of instantiated objects here
        
    

    On Error GoTo 0
    Return

Proc_Err:
    Dim lErrNum As String, sErrSource As String, sErrDesc As String
    lErrNum = VBA.Err.Number
    sErrSource = VBA.Err.Source & vbCrLf & "frmRenamer->" & csProcName
    sErrDesc = VBA.Err.Description
    Resume Proc_Err_Continue
    Resume
    
Proc_Err_Continue:
    GoSub Proc_Cleanup
    MsgBox "The following Error Occurred" & vbCrLf & vbCrLf & _
                 "Source:" & vbCrLf & sErrSource & vbCrLf & vbCrLf & _
                 "Number:" & vbCrLf & lErrNum & vbCrLf & vbCrLf & _
                 "Description:" & vbCrLf & sErrDesc, vbExclamation + vbOKOnly

End Sub


'__________________________________________________
' Scope  : Private
' Type   : Sub
' Name   : cmdBrowseForString_Click
' Params :
' Returns: Nothing
' Desc   : The Sub uses parameters  for cmdBrowseForString_Click and returns Nothing.
'__________________________________________________
' History
' CDK: 20010102: Added Error Trapping & Comments
'__________________________________________________
Private Sub cmdBrowseForString_Click()

    On Error GoTo Proc_Exit
    
    Dim sFileName As String
    Dim sPath As String
    
    On Error GoTo Proc_Error
    ' Set filters.
    CommonDialog1.Filter = "All Files (*.*)"
    ' Specify default filter.
    CommonDialog1.FilterIndex = 2
    
    ' Display the Open dialog box.
    CommonDialog1.ShowOpen
    
    sFileName = CommonDialog1.FileName
    If sFileName & "" <> "" Then
        'get the path for txtFolder and get the search string by
        '   removing the path
        'sPath = moFileOps.PathFromFullPathPs(sFileName)
        sPath = moFileOps.SplitPathPs(sFileName, eSplitPathGetFullPath)
        Me.txtSearch = Replace(sFileName, sPath, "")
        Me.txtFolder = sPath
        
        'fill the list box with all files in the selected dir
        SelectFilesForRenameM
        
    End If

Proc_Exit:
    On Error Resume Next
    Exit Sub
    
Proc_Error:
    MsgBox "an error occurred: " & Err.Number & " " & Err.Description
    Resume Proc_Exit

End Sub


'__________________________________________________
' Scope  : Private
' Type   : Sub
' Name   : cmdCancel_Click
' Params :
' Returns: Nothing
' Desc   : The Sub uses parameters  for cmdCancel_Click and returns Nothing.
'__________________________________________________
' History
' CDK: 20010102: Added Error Trapping & Comments
'__________________________________________________
Private Sub cmdCancel_Click()
    On Error GoTo Proc_Err
    Const csProcName As String = "cmdCancel_Click"


    mfStopProcessing = True


Proc_Exit:
    GoSub Proc_Cleanup
    Exit Sub

Proc_Cleanup:
    On Error Resume Next
    'Place any cleanup of instantiated objects here
    
    On Error GoTo 0
    Return

Proc_Err:
    Dim lErrNum As String, sErrSource As String, sErrDesc As String
    lErrNum = VBA.Err.Number
    sErrSource = VBA.Err.Source & vbCrLf & "frmRenamer->" & csProcName
    sErrDesc = VBA.Err.Description
    Resume Proc_Err_Continue
    Resume
    
Proc_Err_Continue:
    GoSub Proc_Cleanup
    MsgBox "The following Error Occurred" & vbCrLf & vbCrLf & _
                 "Source:" & vbCrLf & sErrSource & vbCrLf & vbCrLf & _
                 "Number:" & vbCrLf & lErrNum & vbCrLf & vbCrLf & _
                 "Description:" & vbCrLf & sErrDesc, vbExclamation + vbOKOnly

End Sub


'__________________________________________________
' Scope  : Private
' Type   : Sub
' Name   : cmdClear_Click
' Params :
' Returns: Nothing
' Desc   : The Sub uses parameters  for cmdClear_Click and returns Nothing.
'__________________________________________________
' History
' CDK: 20010102: Added Error Trapping & Comments
'__________________________________________________
Private Sub cmdClear_Click()
    On Error GoTo Proc_Err
    Const csProcName As String = "cmdClear_Click"


    Dim oCtl As Control

    For Each oCtl In Me.Controls

        If TypeOf oCtl Is TextBox Then
            oCtl.Text = ""
        ElseIf TypeOf oCtl Is ListBox Then
            oCtl.Clear
        ElseIf TypeOf oCtl Is CheckBox Then
            oCtl.Value = vbUnchecked
        End If

    Next


Proc_Exit:
    GoSub Proc_Cleanup
    Exit Sub

Proc_Cleanup:
    On Error Resume Next
    'Place any cleanup of instantiated objects here
    
    On Error GoTo 0
    Return

Proc_Err:
    Dim lErrNum As String, sErrSource As String, sErrDesc As String
    lErrNum = VBA.Err.Number
    sErrSource = VBA.Err.Source & vbCrLf & "frmRenamer->" & csProcName
    sErrDesc = VBA.Err.Description
    Resume Proc_Err_Continue
    Resume
    
Proc_Err_Continue:
    GoSub Proc_Cleanup
    MsgBox "The following Error Occurred" & vbCrLf & vbCrLf & _
                 "Source:" & vbCrLf & sErrSource & vbCrLf & vbCrLf & _
                 "Number:" & vbCrLf & lErrNum & vbCrLf & vbCrLf & _
                 "Description:" & vbCrLf & sErrDesc, vbExclamation + vbOKOnly

End Sub


'__________________________________________________
' Scope  : Private
' Type   : Sub
' Name   : cmdExit_Click
' Params :
' Returns: Nothing
' Desc   : The Sub uses parameters  for cmdExit_Click and returns Nothing.
'__________________________________________________
' History
' CDK: 20010102: Added Error Trapping & Comments
'__________________________________________________
Private Sub cmdExit_Click()
    On Error GoTo Proc_Err
    Const csProcName As String = "cmdExit_Click"


    Unload Me


Proc_Exit:
    GoSub Proc_Cleanup
    Exit Sub

Proc_Cleanup:
    On Error Resume Next
    'Place any cleanup of instantiated objects here
    
    On Error GoTo 0
    Return

Proc_Err:
    Dim lErrNum As String, sErrSource As String, sErrDesc As String
    lErrNum = VBA.Err.Number
    sErrSource = VBA.Err.Source & vbCrLf & "frmRenamer->" & csProcName
    sErrDesc = VBA.Err.Description
    Resume Proc_Err_Continue
    Resume
    
Proc_Err_Continue:
    GoSub Proc_Cleanup
    MsgBox "The following Error Occurred" & vbCrLf & vbCrLf & _
                 "Source:" & vbCrLf & sErrSource & vbCrLf & vbCrLf & _
                 "Number:" & vbCrLf & lErrNum & vbCrLf & vbCrLf & _
                 "Description:" & vbCrLf & sErrDesc, vbExclamation + vbOKOnly

End Sub


'__________________________________________________
' Scope  : Private
' Type   : Sub
' Name   : cmdRename_Click
' Params :
' Returns: Nothing
' Desc   : The Sub uses parameters  for cmdRename_Click and returns Nothing.
'__________________________________________________
' History
' CDK: 20010102: Added Error Trapping & Comments
'__________________________________________________
Private Sub cmdRename_Click()
    On Error GoTo Proc_Err
    Const csProcName As String = "cmdRename_Click"


    'rename all selected files
    RenameFilesM


Proc_Exit:
    GoSub Proc_Cleanup
    Exit Sub

Proc_Cleanup:
    On Error Resume Next
    'Place any cleanup of instantiated objects here
    
    On Error GoTo 0
    Return

Proc_Err:
    Dim lErrNum As String, sErrSource As String, sErrDesc As String
    lErrNum = VBA.Err.Number
    sErrSource = VBA.Err.Source & vbCrLf & "frmRenamer->" & csProcName
    sErrDesc = VBA.Err.Description
    Resume Proc_Err_Continue
    Resume
    
Proc_Err_Continue:
    GoSub Proc_Cleanup
    MsgBox "The following Error Occurred" & vbCrLf & vbCrLf & _
                 "Source:" & vbCrLf & sErrSource & vbCrLf & vbCrLf & _
                 "Number:" & vbCrLf & lErrNum & vbCrLf & vbCrLf & _
                 "Description:" & vbCrLf & sErrDesc, vbExclamation + vbOKOnly

End Sub


'__________________________________________________
' Scope  : Private
' Type   : Sub
' Name   : Form_Load
' Params :
' Returns: Nothing
' Desc   : The Sub uses parameters  for Form_Load and returns Nothing.
'__________________________________________________
' History
' CDK: 20010102: Added Error Trapping & Comments
'__________________________________________________
Private Sub Form_Load()
    On Error GoTo Proc_Err
    Const csProcName As String = "Form_Load"

    Call CenterFormP(Me)
    
    Set moFileOps = New clsFileOps
    SetControlsWhileProcessing (False)


Proc_Exit:
    GoSub Proc_Cleanup
    Exit Sub

Proc_Cleanup:
    On Error Resume Next
    'Place any cleanup of instantiated objects here
        

    On Error GoTo 0
    Return

Proc_Err:
    Dim lErrNum As String, sErrSource As String, sErrDesc As String
    lErrNum = VBA.Err.Number
    sErrSource = VBA.Err.Source & vbCrLf & "frmRenamer->" & csProcName
    sErrDesc = VBA.Err.Description
    Resume Proc_Err_Continue
    Resume
    
Proc_Err_Continue:
    GoSub Proc_Cleanup
    MsgBox "The following Error Occurred" & vbCrLf & vbCrLf & _
                 "Source:" & vbCrLf & sErrSource & vbCrLf & vbCrLf & _
                 "Number:" & vbCrLf & lErrNum & vbCrLf & vbCrLf & _
                 "Description:" & vbCrLf & sErrDesc, vbExclamation + vbOKOnly

End Sub


'__________________________________________________
' Scope  : Private
' Type   : Sub
' Name   : Form_Terminate
' Params :
' Returns: Nothing
' Desc   : The Sub uses parameters  for Form_Terminate and returns Nothing.
'__________________________________________________
' History
' CDK: 20010102: Added Error Trapping & Comments
'__________________________________________________
Private Sub Form_Terminate()

    On Error Resume Next
    Set moFileOps = Nothing

End Sub


'__________________________________________________
' Scope  : Private
' Type   : Sub
' Name   : SetControlsWhileProcessing
' Params :
'          fProcessing As Boolean
' Returns: Nothing
' Desc   : The Sub uses parameters fProcessing As Boolean for SetControlsWhileProcessing and returns Nothing.
'__________________________________________________
' History
' CDK: 20010102: Added Error Trapping & Comments
'__________________________________________________
Private Sub SetControlsWhileProcessing(fProcessing As Boolean)
    On Error GoTo Proc_Err
    Const csProcName As String = "SetControlsWhileProcessing"


    'comment: this sets control state while app is doing rename process
    
    Dim ctl As Control
    
    For Each ctl In Me.Controls
        If TypeOf ctl Is CommandButton Or _
            TypeOf ctl Is TextBox Or _
            TypeOf ctl Is Label Or TypeOf ctl Is ListBox Then
            ctl.Enabled = Not fProcessing
        End If
    Next
       
    Me.cmdCancel.Visible = fProcessing
    Me.cmdCancel.Enabled = fProcessing
    
    Me.cmdRename.Visible = Not fProcessing
    Me.cmdRename.Enabled = Not fProcessing
    
    Me.fraProgress.Visible = fProcessing



Proc_Exit:
    GoSub Proc_Cleanup
    Exit Sub

Proc_Cleanup:
    On Error Resume Next
    'Place any cleanup of instantiated objects here
        

    On Error GoTo 0
    Return

Proc_Err:
    Dim lErrNum As String, sErrSource As String, sErrDesc As String
    lErrNum = VBA.Err.Number
    sErrSource = VBA.Err.Source & vbCrLf & "frmRenamer->" & csProcName
    sErrDesc = VBA.Err.Description
    Resume Proc_Err_Continue
    Resume
    
Proc_Err_Continue:
    GoSub Proc_Cleanup
    MsgBox "The following Error Occurred" & vbCrLf & vbCrLf & _
                 "Source:" & vbCrLf & sErrSource & vbCrLf & vbCrLf & _
                 "Number:" & vbCrLf & lErrNum & vbCrLf & vbCrLf & _
                 "Description:" & vbCrLf & sErrDesc, vbExclamation + vbOKOnly

End Sub



'__________________________________________________
' Scope  : Private
' Type   : Sub
' Name   : SelectFilesForRenameM
' Params :
' Returns: Nothing
' Desc   : The Sub uses parameters  for SelectFilesForRenameM and returns Nothing.
'__________________________________________________
' History
' CDK: 20010102: Added Error Trapping & Comments
'__________________________________________________
Private Sub SelectFilesForRenameM()
    On Error GoTo Proc_Err
    Const csProcName As String = "SelectFilesForRenameM"


    Dim sOldName As String
    Dim sNewName As String
    Dim lCounter As Long
    Dim lReplaced As Long
    Dim sMsg As String
    Dim aArrFilesInSelectedFolder() As String
    Dim lComparisonMethod As Long
    
    Call SetControlsWhileProcessing(True)
    
    'get the files into an array from the specified dir, if no files returned, just exit
    If moFileOps.FilesToArray(Me.txtFolder, False, False, aArrFilesInSelectedFolder) = 0 Then
        GoTo Proc_Exit
    End If
    
    'clear old entries in listbox
    Me.lstFilesToRename.Clear

    'roll through the files, if the search string matches, add to the list of files to operate on
    For lCounter = LBound(aArrFilesInSelectedFolder()) To UBound(aArrFilesInSelectedFolder())
    
        If Me.chkMatchCase = vbChecked Then
            lComparisonMethod = vbBinaryCompare
        Else
            lComparisonMethod = vbTextCompare
        End If
    
        If InStr(1, aArrFilesInSelectedFolder(lCounter), Me.txtSearch, lComparisonMethod) Then
            'add the matching item to the list
            lstFilesToRename.AddItem aArrFilesInSelectedFolder(lCounter)
            'check the item to be included in the rename
            lstFilesToRename.Selected(lstFilesToRename.ListCount - 1) = True
            
        End If
        
        'allow processing to break for user who wants to stop processing of files
        DoEvents
        Me.Refresh
        
    Next
    

Proc_Exit:
    GoSub Proc_Cleanup
    Exit Sub

Proc_Cleanup:
    On Error Resume Next
    Call SetControlsWhileProcessing(False)
    'Place any cleanup of instantiated objects here
    On Error GoTo 0
    Return

Proc_Err:
    Dim lErrNum As String, sErrSource As String, sErrDesc As String
    lErrNum = VBA.Err.Number
    sErrSource = VBA.Err.Source & vbCrLf & "frmRenamer->" & csProcName
    sErrDesc = VBA.Err.Description
    Resume Proc_Err_Continue
    Resume
    
Proc_Err_Continue:
    GoSub Proc_Cleanup
    MsgBox "The following Error Occurred" & vbCrLf & vbCrLf & _
                 "Source:" & vbCrLf & sErrSource & vbCrLf & vbCrLf & _
                 "Number:" & vbCrLf & lErrNum & vbCrLf & vbCrLf & _
                 "Description:" & vbCrLf & sErrDesc, vbExclamation + vbOKOnly

End Sub

'Private Sub cmdRename_Click()
'
'    Dim sOldName As String
'    Dim sNewName As String
'    Dim lCounter As Long
'    Dim lReplaced As Long
'    Dim sMsg As String
'
'
'    Const csProgressCaption As String = "Records Processed: "
'
'    sMsg = "Are you sure you want to replace all occurances of """ & Me.txtSearch & """ with """ _
'        & IIf(Me.txtReplace & "" = "", "<empty string>", Me.txtReplace) & """ in folder " & Me.txtFolder & "?"
'
'    If MsgBox(sMsg, vbYesNoCancel + vbQuestion) <> vbYes Then
'
'        GoTo PROC_EXIT
'
'    End If
'
'    Call SetControlsWhileProcessing(True)
'
'
'    'go to the specified directory and get all files in the dir, rename all those appropriately
'    Dim aArrFilesInSelectedFolder() As String
'
'    'get the files into an array from the specified dir
'    Call moFileOps.FilesToArray(Me.txtFolder, False, False, aArrFilesInSelectedFolder)
'
'    lReplaced = 0
'    Me.prgProgress.Min = LBound(aArrFilesInSelectedFolder)
'    Me.prgProgress.Max = UBound(aArrFilesInSelectedFolder)
'
'
'    'roll through the files, if the search string matches, replace with new name
'    For lCounter = LBound(aArrFilesInSelectedFolder()) To UBound(aArrFilesInSelectedFolder())
'
'
'        If InStr(1, aArrFilesInSelectedFolder(lCounter), Me.txtSearch, vbTextCompare) Then
'
'            lstFilesToRename.AddItem aArrFilesInSelectedFolder(lCounter)
'
'        End If
'
'        'allow processing to break for user who wants to stop processing of files
'        DoEvents
'        Me.Refresh
'        Me.fraProgress.Caption = csProgressCaption & lCounter
'        Me.prgProgress.Value = lCounter
'
'        'deal w/cancel button if pressed
'        If mfStopProcessing Then
'            lCounter = UBound(aArrFilesInSelectedFolder())
'            mfStopProcessing = False
'        End If
'
'
'    Next
'
'    MsgBox "Done replacing files" & vbCrLf & vbCrLf & "Renamed " & lReplaced & " files"
'
'    Call SetControlsWhileProcessing(False)
'
'PROC_EXIT:
'
'
'End Sub
'
'



'__________________________________________________
' Scope  : Private
' Type   : Sub
' Name   : RenameFilesM
' Params :
' Returns: Nothing
' Desc   : The Sub uses parameters  for RenameFilesM and returns Nothing.
'__________________________________________________
' History
' CDK: 20010102: Added Error Trapping & Comments
'__________________________________________________
Private Sub RenameFilesM()
    On Error GoTo Proc_Err
    Const csProcName As String = "RenameFilesM"


    Dim sOldName As String
    Dim sNewName As String
    Dim lCounter As Long
    Dim lReplaced As Long
    Dim sMsg As String
    Dim lTotal As Long
    
    Const csProgressCaption As String = "Records Processed: "
    
    If Me.lstFilesToRename.SelCount = 0 Then
        sMsg = "You have not selected any files to replace, you must select at least one"
        Call MsgBox(sMsg, vbExclamation + vbOKOnly)
        GoTo Proc_Exit
    End If
    
    sMsg = "Are you sure you want to replace all occurances of """ & Me.txtSearch & """ with """ _
        & IIf(Me.txtReplace & "" = "", "<empty string>", Me.txtReplace) & """ in folder " & Me.txtFolder & _
        " for the " & Me.lstFilesToRename.SelCount & " selected files?"
    
    If MsgBox(sMsg, vbYesNoCancel + vbQuestion) <> vbYes Then
    
        GoTo Proc_Exit
    
    End If
    
    Call SetControlsWhileProcessing(True)
    
    
    
    lReplaced = 0
    Me.prgProgress.Min = 0
    Me.prgProgress.Max = Me.lstFilesToRename.SelCount

    lTotal = Me.lstFilesToRename.ListCount - 1

    'roll through the files, if the search string matches, replace with new name
    For lCounter = 0 To lTotal
    
        If Me.lstFilesToRename.Selected(lCounter) = True Then
                
            sOldName = Me.txtFolder & "\" & Me.lstFilesToRename.List(lCounter)
            sNewName = Replace(sOldName, Me.txtSearch, Me.txtReplace, , , vbTextCompare)
            moFileOps.ShellRenameFile Me.hWnd, sOldName, sNewName, False, "Rename " & Me.lstFilesToRename.List(lCounter) & " to " & sNewName
            
            lReplaced = lReplaced + 1
        
        End If
        
        'allow processing to break for user who wants to stop processing of files
        DoEvents
        Me.Refresh
        Me.fraProgress.Caption = csProgressCaption & lReplaced
        Me.prgProgress.Value = lReplaced
        
        'deal w/cancel button if pressed
        If mfStopProcessing Then
            lCounter = lTotal
            mfStopProcessing = False
        End If
        
        
        
    Next
    
    MsgBox "Done replacing files" & vbCrLf & vbCrLf & "Renamed " & lReplaced & " files"
    Call cmdClear_Click
        


Proc_Exit:
    GoSub Proc_Cleanup
    Exit Sub

Proc_Cleanup:
    On Error Resume Next
    Call SetControlsWhileProcessing(False)

    'Place any cleanup of instantiated objects here
    
    On Error GoTo 0
    Return

Proc_Err:
    Dim lErrNum As String, sErrSource As String, sErrDesc As String
    lErrNum = VBA.Err.Number
    sErrSource = VBA.Err.Source & vbCrLf & "frmRenamer->" & csProcName
    sErrDesc = VBA.Err.Description
    Resume Proc_Err_Continue
    Resume
    
Proc_Err_Continue:
    GoSub Proc_Cleanup
    MsgBox "The following Error Occurred" & vbCrLf & vbCrLf & _
                 "Source:" & vbCrLf & sErrSource & vbCrLf & vbCrLf & _
                 "Number:" & vbCrLf & lErrNum & vbCrLf & vbCrLf & _
                 "Description:" & vbCrLf & sErrDesc, vbExclamation + vbOKOnly

End Sub




'__________________________________________________
' Scope  : Private
' Type   : Sub
' Name   : txtFolder_Validate
' Params :
'          Cancel As Boolean
' Returns: Nothing
' Desc   : The Sub uses parameters Cancel As Boolean for txtFolder_Validate and returns Nothing.
'__________________________________________________
' History
' CDK: 20010102: Added Error Trapping & Comments
'__________________________________________________
Private Sub txtFolder_Validate(Cancel As Boolean)
    On Error GoTo Proc_Err
    Const csProcName As String = "txtFolder_Validate"


    'fill the list box with all files in the selected dir
    SelectFilesForRenameM



Proc_Exit:
    GoSub Proc_Cleanup
    Exit Sub

Proc_Cleanup:
    On Error Resume Next
    'Place any cleanup of instantiated objects here
        

    On Error GoTo 0
    Return

Proc_Err:
    Dim lErrNum As String, sErrSource As String, sErrDesc As String
    lErrNum = VBA.Err.Number
    sErrSource = VBA.Err.Source & vbCrLf & "frmRenamer->" & csProcName
    sErrDesc = VBA.Err.Description
    Resume Proc_Err_Continue
    Resume
    
Proc_Err_Continue:
    GoSub Proc_Cleanup
    MsgBox "The following Error Occurred" & vbCrLf & vbCrLf & _
                 "Source:" & vbCrLf & sErrSource & vbCrLf & vbCrLf & _
                 "Number:" & vbCrLf & lErrNum & vbCrLf & vbCrLf & _
                 "Description:" & vbCrLf & sErrDesc, vbExclamation + vbOKOnly

End Sub


'__________________________________________________
' Scope  : Private
' Type   : Sub
' Name   : txtSearch_Validate
' Params :
'          Cancel As Boolean
' Returns: Nothing
' Desc   : The Sub uses parameters Cancel As Boolean for txtSearch_Validate and returns Nothing.
'__________________________________________________
' History
' CDK: 20010102: Added Error Trapping & Comments
'__________________________________________________
Private Sub txtSearch_Validate(Cancel As Boolean)
    On Error GoTo Proc_Err
    Const csProcName As String = "txtSearch_Validate"


    'fill the list box with all files in the selected dir
    SelectFilesForRenameM



Proc_Exit:
    GoSub Proc_Cleanup
    Exit Sub

Proc_Cleanup:
    On Error Resume Next
    'Place any cleanup of instantiated objects here
    
    On Error GoTo 0
    Return

Proc_Err:
    Dim lErrNum As String, sErrSource As String, sErrDesc As String
    lErrNum = VBA.Err.Number
    sErrSource = VBA.Err.Source & vbCrLf & "frmRenamer->" & csProcName
    sErrDesc = VBA.Err.Description
    Resume Proc_Err_Continue
    Resume
    
Proc_Err_Continue:
    GoSub Proc_Cleanup
    MsgBox "The following Error Occurred" & vbCrLf & vbCrLf & _
                 "Source:" & vbCrLf & sErrSource & vbCrLf & vbCrLf & _
                 "Number:" & vbCrLf & lErrNum & vbCrLf & vbCrLf & _
                 "Description:" & vbCrLf & sErrDesc, vbExclamation + vbOKOnly

End Sub

