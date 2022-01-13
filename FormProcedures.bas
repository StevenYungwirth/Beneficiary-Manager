Attribute VB_Name = "FormProcedures"
Option Explicit
'Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
'Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
'Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
'Private Declare Function DrawMenuBar Lib "user32" (ByVal hWnd As Long) As Long

Private Property Get manualSheet() As Worksheet
    Set manualSheet = ThisWorkbook.Sheets("Manual Beneficiaries")
End Property

Private Property Get headerLine() As Range
    Set headerLine = manualSheet.Range("1:1")
End Property

Public Sub AddToSheet(acctToAdd As clsAccount, beneToAdd As clsBeneficiary, action As String)
    'Get the first empty row
    Dim emptyRow As Integer
    emptyRow = FirstEmptyRow
    
    'Put the beneficiary information onto the row
    AddInfoToSheet emptyRow, beneToAdd
    
    'Add the beneficiary to the selected account, and the list of new beneficiaries
    acctToAdd.AddBene beneToAdd
    
    'Add the tracking information to the row
    AddTrackingInfoToSheet emptyRow, beneToAdd.AddDate, beneToAdd.AddedBy, action
End Sub
        
'Public Sub HideBar(frm As Object)
'    'Remove the title bar from the given form
'    Dim Style As Long, Menu As Long, hWndForm As Long
'    hWndForm = FindWindow("ThunderDFrame", frm.Caption)
'    Style = GetWindowLong(hWndForm, &HFFF0)
'    Style = Style And Not &HC00000
'    SetWindowLong hWndForm, &HFFF0, Style
'    DrawMenuBar hWndForm
'End Sub

Private Function FirstEmptyRow() As Integer
    'Return the first row that has no values in it
    If Len(manualSheet.Range("A1").Value2) = 0 Then
        If Len(manualSheet.Range("A2").Value2) = 0 Then
            FirstEmptyRow = 2
        Else
            FirstEmptyRow = 1
        End If
    Else
        FirstEmptyRow = manualSheet.Range("A1").End(xlDown).Row + 1
    End If
End Function

Private Sub AddInfoToSheet(emptyRow As Integer, beneToAdd As clsBeneficiary)
    'Get the columns to put the beneficiary information into
    Dim acctNameCol As Integer, acctNumCol As Integer, acctIDCol As Integer, beneIDCol As Integer, beneNameCol As Integer, beneLevelCol As Integer, benePercentCol As Integer
    acctNameCol = GetHeaderColumn(headerLine, "Account Name/ID")
    acctNumCol = GetHeaderColumn(headerLine, "Account#")
    acctIDCol = GetHeaderColumn(headerLine, "Account ID")
    beneIDCol = GetHeaderColumn(headerLine, "Bene ID")
    beneNameCol = GetHeaderColumn(headerLine, "Name")
    beneLevelCol = GetHeaderColumn(headerLine, "BeneLevel")
    benePercentCol = GetHeaderColumn(headerLine, "Percentage")
    
    'Put the beneficiary information onto the row
    With manualSheet
        .Cells(emptyRow, acctNameCol).Value2 = beneToAdd.account.NameOfAccount
        .Cells(emptyRow, acctNumCol).Value2 = beneToAdd.account.Number
        .Cells(emptyRow, acctIDCol).Value2 = beneToAdd.account.ID
        .Cells(emptyRow, beneIDCol).Value2 = beneToAdd.ID
        .Cells(emptyRow, beneNameCol).Value2 = beneToAdd.NameOfBeneficiary
        .Cells(emptyRow, beneLevelCol).Value2 = beneToAdd.Level
        .Cells(emptyRow, benePercentCol).Value2 = beneToAdd.Percent
    End With
End Sub

Private Sub AddTrackingInfoToSheet(emptyRow As Integer, beneAddDate As String, beneAddedBy As String, action As String)
    'Get the columns to put the account information into
    Dim actionCol As Integer, timeCol As Integer, byCol As Integer
    actionCol = GetHeaderColumn(headerLine, "Action")
    timeCol = GetHeaderColumn(headerLine, "Added")
    byCol = GetHeaderColumn(headerLine, "By")
    
    'Add the tracking information to the row
    With manualSheet
        .Cells(emptyRow, actionCol).Value2 = action
        .Cells(emptyRow, timeCol).Value2 = beneAddDate
        .Cells(emptyRow, byCol).Value2 = beneAddedBy
    End With
End Sub

Private Function GetHeaderColumn(headerLine As Range, Header As String) As Integer
    GetHeaderColumn = headerLine.Find(Header, lookat:=xlWhole).Column
End Function
