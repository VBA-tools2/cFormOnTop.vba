VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cFormOnTop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Keep a UserForm on top (from <https://jkp-ads.com/Articles/keepuserformontop02.asp>)."

'@ModuleDescription("Keep a UserForm on top (from <https://jkp-ads.com/Articles/keepuserformontop02.asp>).")

Option Explicit

'Object variable to trigger application events
Private WithEvents XLApp As Excel.Application
Attribute XLApp.VB_VarHelpID = -1

#If VBA7 Then
    Dim mXLHwnd As LongPtr    'Excel's window handle
    Dim mhwndForm As LongPtr  'The userform's window handle
    Private Declare PtrSafe Function FindWindowA Lib "user32" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
    #If Win64 Then
        Private Declare PtrSafe Function SetWindowLongA Lib "user32" Alias "SetWindowLongPtrA" (ByVal hwnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
    #Else
        Private Declare PtrSafe Function SetWindowLongA Lib "user32" (ByVal hwnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
    #End If
    Private Declare PtrSafe Function SetForegroundWindow Lib "user32" (ByVal hwnd As LongPtr) As Long
#Else
    Dim mXLHwnd As Long    'Excel's window handle
    Dim mhwndForm As Long  'The userform's window handle
    Dim FirstWindow As Long
    Private Declare Function FindWindowA Lib "user32" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Private Declare Function SetWindowLongA Lib "user32" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
#End If

Const GWL_HWNDPARENT As Long = -8

Private moTheUserform As Object

Public Sub InitializeMe()
    If Val(Application.Version) >= 15 Then        'Only makes sense on Excel 2013 and up
        Set XLApp = Application
    End If
End Sub

Private Sub Class_Terminate()
    Set XLApp = Nothing
    Set moTheUserform = Nothing
End Sub

Private Sub XLApp_WindowActivate(ByVal Wb As Workbook, ByVal Wn As Window)
    If Val(Application.Version) >= 15 And mhwndForm <> 0 Then  'Basear o form na janela ativa do Excel.
        mXLHwnd = Application.hwnd    'Always get because in Excel 15 SDI each wb has its window with different handle.
        SetWindowLongA mhwndForm, GWL_HWNDPARENT, mXLHwnd
        SetForegroundWindow mhwndForm
    End If
End Sub

Private Sub XLApp_WindowResize(ByVal Wb As Workbook, ByVal Wn As Window)
    If Not moTheUserform.Visible Then moTheUserform.Show vbModeless
End Sub

'Private Sub XLApp_WorkbookBeforeClose(ByVal Wb As Workbook, Cancel As Boolean)
'    SetWindowLongA mhwndForm, GWL_HWNDPARENT, 0&
'End Sub
Private Sub XLApp_WindowDeactivate(ByVal Wb As Workbook, ByVal Wn As Excel.Window)
    SetWindowLongA mhwndForm, GWL_HWNDPARENT, 0&
End Sub

Public Property Set TheUserform(ByVal oNewValue As Object)
    Set moTheUserform = oNewValue
    mhwndForm = FindWindowA("ThunderDFrame", moTheUserform.Caption)
End Property
