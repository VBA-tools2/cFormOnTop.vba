Attribute VB_Name = "modShowForm"

Option Explicit

Private frmWorkbooks_Normal As ufWorkbooks_Normal
Private frmWorkbooks_OnTop As ufWorkbooks_OnTop

Public Sub ShowForm_Normal()
    If Workbooks.Count = 1 Then Workbooks.Add
    Set frmWorkbooks_Normal = New ufWorkbooks_Normal
    frmWorkbooks_Normal.Show vbModeless
End Sub

Public Sub ShowForm_OnTop()
    If Workbooks.Count = 1 Then Workbooks.Add
    Set frmWorkbooks_OnTop = New ufWorkbooks_OnTop
    frmWorkbooks_OnTop.Show vbModeless
End Sub

