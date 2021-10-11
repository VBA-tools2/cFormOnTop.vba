VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufWorkbooks_Normal 
   Caption         =   "List of open Workbooks"
   ClientHeight    =   3180
   ClientLeft      =   2040
   ClientTop       =   2370
   ClientWidth     =   4710
   OleObjectBlob   =   "ufWorkbooks_Normal.frx":0000
End
Attribute VB_Name = "ufWorkbooks_Normal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub UserForm_Initialize()
    Dim oWb As Workbook
    For Each oWb In Workbooks
        lbxWorkbooks.AddItem oWb.Name
    Next
End Sub

Private Sub lbxWorkbooks_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Workbooks(lbxWorkbooks.Value).Windows(1).Activate
    Application.Goto Workbooks(lbxWorkbooks.Value).Worksheets(1).Range("A1")
End Sub

