Attribute VB_Name = "modPublic"
Option Explicit

Private Const mcsModuleName As String = "modPublic"


'__________________________________________________
' Scope  : Public
' Type   : Sub
' Name   : CenterFormP
' Params :
'          oform As Form
' Returns: Nothing
' Desc   : The Sub uses parameters oform As Form for CenterFormP and returns Nothing.
'__________________________________________________
' History
' CDK: 20010102: Added Error Trapping & Comments
'__________________________________________________
Public Sub CenterFormP(oform As Form)
  ' Comments  : Centers the form on the screen
  ' Parameters: none
  ' Returns   : Nothing
  ' Source    : Total VB SourceBook 6
    
    Const csProcName As String = "CenterForm_p"
    On Error GoTo Proc_Error
    
    oform.Move (Screen.Width - oform.Width) / 2, _
    (Screen.Height - oform.Height) / 2

Proc_Exit:
    Exit Sub

Proc_Error:
    Err.Raise vbObjectError Or Err.Number, mcsModuleName & "." & csProcName, Err.Description
    Resume Proc_Exit

End Sub





