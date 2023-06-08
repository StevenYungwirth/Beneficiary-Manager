Attribute VB_Name = "FormProcedures"
Option Explicit
'Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
'Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
'Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
'Private Declare Function DrawMenuBar Lib "user32" (ByVal hWnd As Long) As Long
        
'Public Sub HideBar(frm As Object)
'    'Remove the title bar from the given form
'    Dim Style As Long, Menu As Long, hWndForm As Long
'    hWndForm = FindWindow("ThunderDFrame", frm.Caption)
'    Style = GetWindowLong(hWndForm, &HFFF0)
'    Style = Style And Not &HC00000
'    SetWindowLong hWndForm, &HFFF0, Style
'    DrawMenuBar hWndForm
'End Sub

Private Property Get manualSheet() As Worksheet
    Set manualSheet = ThisWorkbook.Sheets("Manual Beneficiaries")
End Property

Private Property Get HeaderLine() As Range
    Set HeaderLine = manualSheet.Range("1:1")
End Property

Public Sub AddToSheet(acctToAdd As clsAccount, beneToAdd As clsBeneficiary, action As String)
    'Get the first empty row
    Dim emptyRow As Integer
    emptyRow = FirstEmptyRow
    
    'Unlock the manual bene sheet
    manualSheet.Unprotect password:=ProjectGlobals.manualSheetPassword
    
    'Put the beneficiary information onto the row
    AddInfoToSheet emptyRow, beneToAdd
    
    'Add the beneficiary to the selected account, and the list of new beneficiaries
    acctToAdd.AddBene beneToAdd, True
    
    'Add the tracking information to the row
    AddTrackingInfoToSheet emptyRow, beneToAdd.addDate, beneToAdd.AddedBy, action
    
    'Lock the manual bene sheet
    manualSheet.Protect password:=ProjectGlobals.manualSheetPassword
End Sub

Public Function SortHouseholdList(origHouseholdList As Dictionary) As Dictionary
    'Turn off screen updating
    ProjectGlobals.StateToggle False

    'Create a temporary worksheet
    Dim tempSheet As Worksheet
    Set tempSheet = ThisWorkbook.Sheets.Add
    
    'Put all of the keys from origHouseholdList into the worksheet
    Dim rw As Integer
    Dim dictKey As Variant
    For rw = 1 To origHouseholdList.count
        tempSheet.Cells(rw, 1).Value2 = origHouseholdList.Keys(rw - 1)
    Next rw
    
    'Sort the worksheet
    tempSheet.UsedRange.Sort Key1:=tempSheet.Cells(1, 1), order1:=xlAscending, Header:=xlNo
    
    'Create a temporary dictionary and set the key-item pairs with the sorted list
    Dim tempDict As Dictionary: Set tempDict = New Dictionary
    For rw = 1 To tempSheet.UsedRange.Rows.count
        tempDict.Add tempSheet.Cells(rw, 1).Value2, origHouseholdList.Item(tempSheet.Cells(rw, 1).Value2)
    Next rw
    
    'Overwrite the original dictionary with the temporary one
    Set SortHouseholdList = tempDict
    
    'Delete the hidden worksheet
    Application.DisplayAlerts = False
    tempSheet.Delete
    Application.DisplayAlerts = True

    'Turn screen updating back on
    ProjectGlobals.StateToggle True
End Function

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
    acctNameCol = GetHeaderColumn(HeaderLine, "Account Name/ID")
    acctNumCol = GetHeaderColumn(HeaderLine, "Account#")
    acctIDCol = GetHeaderColumn(HeaderLine, "Morningstar ID")
    beneIDCol = GetHeaderColumn(HeaderLine, "Bene ID")
    beneNameCol = GetHeaderColumn(HeaderLine, "Name")
    beneLevelCol = GetHeaderColumn(HeaderLine, "BeneLevel")
    benePercentCol = GetHeaderColumn(HeaderLine, "Percentage")
    
    'Put the beneficiary information onto the row
    With manualSheet
        .Cells(emptyRow, acctNameCol).Value2 = beneToAdd.account.NameOfAccount
        .Cells(emptyRow, acctNumCol).Value2 = beneToAdd.account.Number
        .Cells(emptyRow, acctIDCol).Value2 = beneToAdd.account.morningstarID
        .Cells(emptyRow, beneIDCol).Value2 = beneToAdd.id
        .Cells(emptyRow, beneNameCol).Value2 = beneToAdd.NameOfBeneficiary
        .Cells(emptyRow, beneLevelCol).Value2 = beneToAdd.Level
        .Cells(emptyRow, benePercentCol).Value2 = beneToAdd.Percent
    End With
End Sub

Private Sub AddTrackingInfoToSheet(emptyRow As Integer, beneAddDate As String, beneAddedBy As String, action As String)
    'Get the columns to put the account information into
    Dim actionCol As Integer, timeCol As Integer, byCol As Integer
    actionCol = GetHeaderColumn(HeaderLine, "Action")
    timeCol = GetHeaderColumn(HeaderLine, "Added")
    byCol = GetHeaderColumn(HeaderLine, "By")
    
    'Add the tracking information to the row
    With manualSheet
        .Cells(emptyRow, actionCol).Value2 = action
        .Cells(emptyRow, timeCol).Value2 = beneAddDate
        .Cells(emptyRow, byCol).Value2 = beneAddedBy
        
        If action = "Delete" Then
            .Cells(emptyRow, timeCol).Value2 = Format(Now(), "m/d/yy h:mm;@")
            .Cells(emptyRow, byCol).Value2 = VBA.Environ("username")
        End If
    End With
End Sub

Private Function GetHeaderColumn(HeaderLine As Range, Header As String) As Integer
    GetHeaderColumn = HeaderLine.Find(Header, lookat:=xlWhole).Column
End Function
