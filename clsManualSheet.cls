VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsManualSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_datasheet As clsDataSheet
Private Const worksheetName As String = "Manual Beneficiaries"
Private Const accountNameHeader As String = "Account Name/ID"
Private Const accountNumberHeader As String = "Account#"
Private Const accountIDHeader As String = "Account ID"
Private Const beneIDHeader As String = "Bene ID"
Private Const beneNameHeader As String = "Name"
Private Const beneLevelHeader As String = "BeneLevel"
Private Const benePercentHeader As String = "Percentage"
Private Const actionHeader As String = "Action"
Private Const addDateHeader As String = "Added"
Private Const addedByHeader As String = "By"

Public Sub ClassBuilder()
    'Set the required headers
    Dim requiredHeaders() As String
    requiredHeaders = Split(accountNameHeader & "," & accountNumberHeader & "," & accountIDHeader & "," _
                          & beneIDHeader & "," & beneNameHeader & "," & beneLevelHeader & "," & benePercentHeader & "," _
                          & actionHeader & "," & addDateHeader & "," & addedByHeader, ",")
    
    'Set the datasheet in order to get the data
    Set m_datasheet = ClassConstructor.NewDataSheet(filePath:=vbNullString, wkstName:=worksheetName, reqHeaders:=requiredHeaders)
End Sub

Public Sub AddManualBenes(accountDict As Dictionary)
    If UBound(m_datasheet.data, 1) <= 1 Then
        'There are no manual beneficiaries
        Exit Sub
    End If
    
    'For each row in Manual Beneficiaries, find the account in the dictionary and add the beneficiary
    Dim rw As Integer
    For rw = 1 To UBound(m_datasheet.data, 1) - 1
        Dim foundAccount As clsAccount
        Set foundAccount = GetAccountFromRow(rw, accountDict)
        
        If foundAccount Is Nothing Then
            'The account wasn't found
            Exit Sub
        End If
        
        'The account was found in the dictionary. Get the beneficiary from the row
        Dim foundBene As clsBeneficiary
        Set foundBene = GetBeneFromRow(rw)
        
        'Given the row's action, call the corresponding function
        Dim action As String
        action = GetData(actionHeader, rw)
        If action = "Added" Then
            'Add the beneficiary to the account
            foundAccount.AddBene foundBene
        ElseIf action = "Updated" Then
            'Update the beneficiary's information
            foundAccount.UpdateBene GetData(beneIDHeader, rw), foundBene
        ElseIf action = "Deleted" Then
            'Remove the beneficiary from the account
            foundAccount.RemoveBene foundBene
        End If
    Next rw
End Sub

Public Function GetData(headerName As String, datapoint As Integer) As String
    GetData = m_datasheet.GetData(headerName, datapoint)
End Function

Public Sub SetData(newData As Variant, headerName As String, datapoint As Integer)
    m_datasheet.SetData newData, headerName, datapoint
End Sub

Public Sub FillWorksheet(sht As Worksheet)
    m_datasheet.FillWorksheet sht
End Sub

Private Function GetAccountFromRow(rw As Integer, accountDict As Dictionary) As clsAccount
    'Get the account name and number
    Dim accountName As String, accountNumber As String
    accountName = GetData(accountNameHeader, rw)
    accountNumber = GetData(accountNumberHeader, rw)
    
    'Find the account in the dictionary
    If accountDict.Exists(accountNumber) Then
        Set GetAccountFromRow = accountDict(accountNumber)
    ElseIf accountDict.Exists(accountName & accountNumber) Then
        Set GetAccountFromRow = accountDict(accountName & accountNumber)
    End If
End Function

Private Function GetBeneFromRow(rw As Integer) As clsBeneficiary
    'Initialize the beneficiary using the row's data
    Set GetBeneFromRow = ClassConstructor.NewBene(GetData(beneNameHeader, rw), GetData(beneLevelHeader, rw), GetData(benePercentHeader, rw))
    
    'Get the beneficiary ID, account name and number
    With GetBeneFromRow
        If Len(GetData(beneIDHeader, rw)) > 0 Then
            .ID = GetData(beneIDHeader, rw)
        End If
        .account.NameOfAccount = GetData(accountNameHeader, rw)
        .account.Number = GetData(accountNumberHeader, rw)
    End With
End Function
