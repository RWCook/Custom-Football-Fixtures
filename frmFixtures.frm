VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmFixtures 
   Caption         =   "Add Fixtures"
   ClientHeight    =   4695
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9330
   OleObjectBlob   =   "frmFixtures.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmFixtures"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'==============================
'Name:btnCancel_Click
'Purpose: Cancel button closes form.
'
'==============================
Private Sub btnCancel_Click()
Unload Me


End Sub

'==============================
'Name:btnOk_Click
'Purpose: Passes the user form data on to add data to the custom fixtures list.
'==============================

Private Sub btnOk_Click()
Dim strCompetition As String
Dim strDate As String
Dim booFiltered As Boolean
Dim booCheckCompetition As Boolean
Dim booCheckDate As Boolean
Dim booIncludeDate As Boolean

strCompetition = cmbCompetition.value
strDate = cmbDate.value
booFiltered = chkFilter.value
booIncludeDate = chkIncludeDate.value

booCheckCompetition = ValidateInput(strCompetition)  'check that something has been selected
If booCheckCompetition = False Then
    MsgBox prompt:="Please select a competition."
    Exit Sub
End If

booCheckDate = ValidateInput(strDate)          'check that something has been selected
If booCheckDate = False Then
    MsgBox prompt:="Please select a date."
    Exit Sub
End If

If booCheckDate = True And booCheckCompetition = True Then
    Unload Me
    AddFixtures strCompetition, strDate, booFiltered, booIncludeDate
    
End If

End Sub

'==============================
'Name:ValidateInput
'Purpose: The Competition and Date drop downs are both constrained to the list of values.
'           This function validates that something has been selected from the list and it is not empty.
'==============================

Private Function ValidateInput(ByVal strToTest As String) As Boolean
If strToTest = vbNullString Then
ValidateInput = False
Else:
ValidateInput = True
End If
End Function

'==============================
'Name:UserForm_Initialize
'Purpose: Standard form initialisation.
'==============================

Private Sub UserForm_Initialize()


chkFilter.Caption = Application.ThisWorkbook.Worksheets("Filter").Range("a1").value

Dim i As Integer
chkIncludeDate.value = True
For i = 1 To UBound(arrCompetitions)
cmbCompetition.AddItem (arrCompetitions(i))
Next i

For i = 1 To UBound(arrDates)
cmbDate.AddItem (arrDates(i))
Next i

End Sub
