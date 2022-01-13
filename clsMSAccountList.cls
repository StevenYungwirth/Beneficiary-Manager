VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMSAccountList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_datasheet As clsDataSheet
Private Const worksheetName As String = "MS Accounts"
Private Const accountNameHeader As String = "Account Name/ID"
Private Const accountNumberHeader As String = "Account Number"
Private Const custodianHeader As String = "Current Custodian"
Private Const householdHeader As String = "Client / Prospect Name"
Private Const ownerHeader As String = "Account Owner"
Private Const accountTypeHeader As String = "Account Type"
Private Const defaultCustodian As String = "TD Ameritrade Institutional"

Private Property Get marketValueHeader() As String
    marketValueHeader = "Market Value " & Chr(10) & "USD"
End Property

Public Sub ClassBuilder(path As String, sort As Boolean)
    'Set the required headers
    Dim requiredHeaders() As String
    requiredHeaders = Split(accountNameHeader & "," & accountNumberHeader & "," & custodianHeader & "," _
                          & householdHeader & "," & marketValueHeader & "," & ownerHeader & "," & accountTypeHeader, ",")
    
    'Set the datasheet in order to get the data
    If sort Then
        Set m_datasheet = ClassConstructor.NewDataSheet(filePath:=path, wkstName:=worksheetName, reqHeaders:=requiredHeaders, sortColumnHeader:=accountNameHeader)
    Else
        Set m_datasheet = ClassConstructor.NewDataSheet(filePath:=path, wkstName:=worksheetName, reqHeaders:=requiredHeaders, sortColumnHeader:=vbNullString)
    End If
End Sub

Public Function GetAccounts(Optional households As Dictionary) As Dictionary
    'Load the array of Associated Bank account names
    Dim associatedAccounts() As String
    associatedAccounts = LoadAssociatedAccounts

    'Initialize a temporary dictionary to return
    Dim tempAccountDict As Dictionary
    Set tempAccountDict = New Dictionary

    'For each index in the arrays, add the accounts to the households
    Dim rw As Integer
    For rw = 1 To UBound(m_datasheet.data, 1) - 1
        'Get the account from the row
        Dim tempAccount As clsAccount
        Set tempAccount = ClassConstructor.NewAccount(accountName:=GetData(accountNameHeader, rw), accountNumber:=GetData(accountNumberHeader, rw), _
                                                      accountType:=GetData(accountTypeHeader, rw), accountCustodian:=GetData(custodianHeader, rw), _
                                                      accountTag:=AutoTag(GetData(accountNameHeader, rw), associatedAccounts), marketValue:=Val(GetData(marketValueHeader, rw)))
        If tempAccount.custodian <> defaultCustodian And tempAccount.Balance > 1 Then
            'Heldaway accounts with a near-zero balance (1 is an arbitrary number) aren't active. Mark the other ones as active (TD accounts will be marked as active later)
            tempAccount.Active = True
        End If
        
        If Not households Is Nothing Then
            'Add the account to the member on this row
            GetMemberFromRow(households, rw).AddAccount tempAccount
        End If
        
        'Add the account to the dictionary, using different keys for heldaway accounts
        If tempAccount.custodian = defaultCustodian Then
            If tempAccountDict.Exists(tempAccount.Number) Then Stop
            tempAccountDict.Add tempAccount.Number, tempAccount
        Else
            If tempAccountDict.Exists(tempAccount.NameOfAccount & tempAccount.Number) Then Stop
            tempAccountDict.Add tempAccount.NameOfAccount & tempAccount.Number, tempAccount
        End If
    Next rw
    
    'Return the dictionary
    Set GetAccounts = tempAccountDict
End Function

Public Function GetData(headerName As String, datapoint As Integer) As String
    GetData = m_datasheet.GetData(headerName, datapoint)
End Function

Public Sub SetData(newData As String, headerName As String, datapoint As Integer)
    m_datasheet.SetData newData, headerName, datapoint
End Sub

Public Sub FillWorksheet(sht As Worksheet)
    m_datasheet.FillWorksheet sht
End Sub

Private Function LoadAssociatedAccounts() As String()
    'Set the location of the file with the list of Associated Bank accounts
    Dim AssociatedFileLocation As String
    AssociatedFileLocation = "Z:\YungwirthSteve\Beneficiary Report\Assets\associated accounts.txt"

    If Dir(AssociatedFileLocation) <> vbNullString Then
        'The file exists; load it
        Dim fs As FileSystemObject
        Set fs = New FileSystemObject
        Dim associatedFile As TextStream
        Set associatedFile = fs.OpenTextFile(AssociatedFileLocation, ForReading, True)
        
        'Return the array of Associated account names
        LoadAssociatedAccounts = Split(associatedFile.ReadAll, vbLf)
        
        'Close the file
        associatedFile.Close
    Else
        'Return an empty string in the first index
        ReDim LoadAssociatedAccounts(0) As String
    End If
End Function

Private Function AutoTag(accountName As String, associatedAccountNames() As String) As String
    'Add WEC or Associated tags if it's easily identifiable or in the list of known Associated Bank account names
    If Len(accountName) > 0 And (UBound(Filter(associatedAccountNames, accountName)) > -1 Or InStr(accountName, " Associated ") > 0) Then
        AutoTag = "Associated"
    ElseIf InStr(accountName, " WEC ") > 0 Then
        AutoTag = "WEC"
    Else
        AutoTag = vbNullString
    End If
End Function

Private Function GetMemberFromRow(households As Dictionary, rw As Integer) As clsMember
    'Get the household to look at
    Dim tempHousehold As clsHousehold
    Set tempHousehold = GetHouseholdFromRow(households, rw)
    
    'Get the member
    Dim acctOwner As String
    acctOwner = GetData(ownerHeader, rw)
    Dim tempMember As clsMember
    Set tempMember = tempHousehold.GetMember(acctOwner)
    If tempMember Is Nothing Then
        'Add a new member to the household
        Set tempMember = ClassConstructor.NewMemberNameOnly(acctOwner)
        tempMember.Active = True
        tempHousehold.AddMember tempMember
    End If
    
    Set GetMemberFromRow = tempMember
End Function

Private Function GetHouseholdFromRow(households As Dictionary, rw As Integer) As clsHousehold
    Dim householdName As String
    householdName = GetData(householdHeader, rw)
    
    'See if the household is in the given dictionary
    If households.Exists(householdName) Then
        'The family is in the dictionary
        Set GetHouseholdFromRow = households(householdName)
    Else
        'The family's not in the dictionary, add it
        Set GetHouseholdFromRow = ClassConstructor.NewHousehold(householdName:=householdName)
        households.Add householdName, GetHouseholdFromRow
    End If
End Function
