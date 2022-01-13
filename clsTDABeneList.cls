VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTDABeneList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_datasheet As clsDataSheet
Private Const worksheetName As String = "TDA Bene List"
Private Const accountNumberHeader As String = "Account#"
Private Const accountTypeHeader As String = "AcctDescription"
Private Const openDateHeader As String = "OpenDate"
Private Const closeDateHeader As String = "DateClosed"
Private Const name1Header As String = "NameLine1"
Private Const name2Header As String = "NameLine2"
Private Const name3Header As String = "NameLine3"
Private Const beneNameHeader As String = "Name"
Private Const beneRelationshipHeader As String = "Relationship"
Private Const beneLevelHeader As String = "BeneLevel"
Private Const benePercentHeader As String = "Percentage"

Public Sub ClassBuilder(path As String)
    'Set the required headers
    Dim requiredHeaders() As String
    requiredHeaders = Split(accountNumberHeader & "," & accountTypeHeader & "," & openDateHeader & "," & closeDateHeader _
                            & "," & name1Header & "," & name2Header & "," & name3Header _
                            & "," & beneNameHeader & "," & beneRelationshipHeader & "," & beneLevelHeader & "," & benePercentHeader, ",")
    
    'Set the datasheet in order to get the data
    Set m_datasheet = ClassConstructor.NewDataSheet(filePath:=path, wkstName:=worksheetName, reqHeaders:=requiredHeaders)
End Sub

Public Function ReadAccounts() As Dictionary
    'Initialize a dictionary to return
    Dim tempAccounts As Dictionary
    Set tempAccounts = New Dictionary
    
    'Loop through each row to build the dictionary
    Dim rw As Integer
    For rw = 2 To UBound(m_datasheet.data, 1) - 1
        'Get the account number on this row
        Dim accountNumber As String
        accountNumber = GetData(accountNumberHeader, rw)
        
        'See if the account number is in the dictionary already
        Dim tempAccount As clsAccount
        If tempAccounts.Exists(accountNumber) Then
            'The account number is already in the dictionary. Add the new beneficiary to it
            Set tempAccount = tempAccounts.Item(accountNumber)
            tempAccount.AddBene GetBeneFromTDRow(rw)
        Else
            'The account number is not already in the dictionary. Add it
            Set tempAccount = GetAccountFromTDRow(rw)
            
            'Add the account to the dictionary
            tempAccounts.Add tempAccount.Number, tempAccount
        End If
    Next rw
    
    'Return the dictionary
    Set ReadAccounts = tempAccounts
End Function

Public Sub AddBenesFromTD(accountDict As Dictionary)
    'Set a string to write to the error file
    Dim errorString As String
    
    'Set an array for missing accounts
    Dim missingAccounts() As String
    ReDim missingAccounts(0) As String
    
    'For each row, find the account and add the beneficiaries
    Dim rw As Integer
    For rw = 1 To UBound(m_datasheet.data, 1) - 1
        Dim accountNumber As String
        accountNumber = GetData(accountNumberHeader, rw)
        If accountDict.Exists(accountNumber) Then
            'The account is in the dictionary. Add the bene if it's there and mark it as active
            If Len(GetData(beneNameHeader, rw)) > 0 Then
                Dim beneToAdd As clsBeneficiary
                Set beneToAdd = ClassConstructor.NewBene(beneName:=GetData(beneNameHeader, rw), beneLevel:=GetData(beneLevelHeader, rw), _
                                                         benePercent:=CDbl(GetData(benePercentHeader, rw)), beneRelation:=GetData(beneRelationshipHeader, rw))
                beneToAdd.account.Number = accountNumber
                accountDict.Item(accountNumber).AddBene beneToAdd
            End If
            
            'Check if the account is closed
            If Len(GetData(closeDateHeader, rw)) > 0 Then
                'The account is closed. Mark it as inactive
                accountDict.Item(accountNumber).Active = False
            Else
                'The account is open. Mark it as active
                accountDict.Item(accountNumber).Active = True
            End If
        Else
            'The account isn't in the dictionary. Add it to the missing accounts
            Dim missingAccount As String
            missingAccount = accountNumber & " - " & GetData(name1Header, rw) & " " & GetData(name2Header, rw) & " " & GetData(name3Header, rw)
            If Not IsInArray(missingAccount, missingAccounts) Then
                If Len(missingAccounts(0)) = 0 Then
                    missingAccounts(0) = missingAccount
                Else
                    ReDim Preserve missingAccounts(0 To UBound(missingAccounts) + 1)
                    missingAccounts(UBound(missingAccounts)) = missingAccount
                End If
            End If
        End If
    Next rw
    
    'Log all errors to the error file
    ErrorHandling.LogErrorToFile ErrorHandling.AccountsNotInMSError(missingAccounts)
End Sub

Public Function GetData(headerName As String, datapoint As Integer) As String
    GetData = m_datasheet.GetData(headerName, datapoint)
End Function

Public Sub SetData(newData As String, headerName As String, datapoint As Integer)
    m_datasheet.SetData newData, headerName, datapoint
End Sub

Public Sub FillWorksheet(sht As Worksheet)
    m_datasheet.FillWorksheet sht
End Sub

Private Function GetAccountFromTDRow(rw As Integer) As clsAccount
    'Set the account on this row
    Dim tempAccount As clsAccount
    Set tempAccount = ClassConstructor.NewAccount(accountName:=vbNullString, accountNumber:=GetData(accountNumberHeader, rw), _
                                                  accountType:=GetData(accountTypeHeader, rw), accountCustodian:="TD Ameritrade Institutional", _
                                                  accountTag:=vbNullString)
    
    'Add the close date to the account
    If IsDate(GetData(closeDateHeader, rw)) Then
        tempAccount.CloseDate = CDate(GetData(closeDateHeader, rw))
    End If
    
    'Add the beneficiary on this row
    If Not GetBeneFromTDRow(rw) Is Nothing Then
        tempAccount.AddBene GetBeneFromTDRow(rw)
    End If
    
    'Return the account
    Set GetAccountFromTDRow = tempAccount
End Function

Private Function GetBeneFromTDRow(rw As Integer) As clsBeneficiary
    'Get the beneficiary on this row
    Dim tempBene As clsBeneficiary
    Set tempBene = ClassConstructor.NewBene(beneName:=GetData(beneNameHeader, rw), beneLevel:=GetData(beneLevelHeader, rw), _
                                            benePercent:=Val(GetData(benePercentHeader, rw)), beneRelation:=GetData(beneRelationshipHeader, rw))
    tempBene.account.Number = GetData(accountNumberHeader, rw)
                                            
    'Return the beneficiary, if there's one on this row
    If Len(tempBene.NameOfBeneficiary) > 0 Then
        Set GetBeneFromTDRow = tempBene
    End If
End Function

Private Function IsInArray(searchTerm As String, arr() As String) As Boolean
    'Check the array for the search term
    Dim filteredList() As String
    filteredList = Filter(arr, searchTerm)
    Dim name As Integer
    Do While name <= UBound(filteredList) And Not IsInArray
        If filteredList(name) = searchTerm Then
            'The search term was found in the array
            IsInArray = True
        End If
        name = name + 1
    Loop
End Function
