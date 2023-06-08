VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMSHouseholdExport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_datasheet As clsDataSheet
Private Const worksheetName As String = "MS Household Export"
Private Const householdIDHeader As String = "HouseholdId"
Private Const accountIDHeader As String = "AccountId"
Private Const accountNumberHeader As String = "AccountNumber"
Private Const accountNameHeader As String = "AccountName"
Private Const accountTypeHeader As String = "AccountType"
Private Const custodianHeader As String = "AccountCustodian"
Private Const householdNameHeader As String = "HouseholdDescription"
Private Const newHouseholdDictName As String = "New Household"
Private Const newAccountDictName As String = "New Account"
Private Const dupeHouseholdDictName As String = "Duplicate Household"
Private Const dupeAccountDictName As String = "Duplicate Account"
Private Const diffHouseholdIDDictName As String = "Different Household ID"
Private Const diffHouseholdNameDictName As String = "Different Household Name"
Private Const diffAccountIDDictName As String = "Different Account ID"
Private Const diffAccountNameDictName As String = "Different Account Name"
Private Const diffAccountNumberDictName As String = "Different Account Number"
Private Const diffAccountTypeDictName As String = "Different Account Type"
Private Const diffCustodianDictName As String = "Different Custodian"
Private Const movedAccountDictName As String = "Moved Account"

Public Property Get DataSize() As Integer
    DataSize = m_datasheet.DataSize
End Property

Private Property Get ErrorDictKeys() As String()
    Dim keyArray(0 To 11) As String
    keyArray(0) = newHouseholdDictName
    keyArray(1) = newAccountDictName
    keyArray(2) = dupeHouseholdDictName
    keyArray(3) = dupeAccountDictName
    keyArray(4) = diffHouseholdIDDictName
    keyArray(5) = diffHouseholdNameDictName
    keyArray(6) = diffAccountIDDictName
    keyArray(7) = diffAccountNameDictName
    keyArray(8) = diffAccountNumberDictName
    keyArray(9) = diffAccountTypeDictName
    keyArray(10) = diffCustodianDictName
    keyArray(11) = movedAccountDictName
    ErrorDictKeys = keyArray
End Property

Private Property Get NewHouseholdDictMessage(householdName As String) As String
    NewHouseholdDictMessage = "Household Added - " & householdName & " with its accounts in an empty member"
End Property

Private Property Get NewAccountDictMessage(accountName As String, accountNumber As String, householdName As String) As String
    NewAccountDictMessage = "Account Added - " & accountName & " - " & accountNumber & " | Household - " & householdName & " under an empty member"
End Property

Private Property Get DupeHouseholdDictMessage(householdName As String) As String
    DupeHouseholdDictMessage = "Household Duplicate - " & householdName & " - Multiple instances found in XML"
End Property

Private Property Get DupeAccountDictMessage(accountName As String, accountNumber As String) As String
    DupeAccountDictMessage = "Account Duplicate - " & accountName & " - " & accountNumber & ": Multiple instances found in XML"
End Property

Private Property Get MovedAccountDictMessage(accountName As String, accountNumber As String, originalHouseholdName As String, newHouseholdName As String) As String
    MovedAccountDictMessage = "Account Moved - " & accountName & " - " & accountNumber & ": from household " & originalHouseholdName & " to " & newHouseholdName
End Property

Private Property Get EmptyMemberName() As String
    EmptyMemberName = ProjectGlobals.m_emptyMemberName
End Property

Private Property Get sheetName() As String
    sheetName = ProjectGlobals.m_msExportName
End Property

Private Property Get XMLClientList() As DOMDocument60
    Set XMLClientList = ProjectGlobals.ClientListFile
End Property

Public Sub ClassBuilder(path As String)
    'Set the required headers
    Dim requiredHeaders() As String
    requiredHeaders = Split(householdIDHeader & "," & accountIDHeader & "," & accountNumberHeader & "," _
                          & accountNameHeader & "," & custodianHeader & "," & householdNameHeader, ",")
    
    'Set the datasheet in order to get the data
    Set m_datasheet = ClassConstructor.NewDataSheet(filePath:=path, wkstName:=worksheetName, reqHeaders:=requiredHeaders)
End Sub

Public Function ImportToXML(Optional testingLimit As Integer) As Dictionary
    'Contains master Morningstar household/account ID, account name
    'Also contains account number, account taxable, custodian, account type, household name
    
    'Set up dictionaries for the possible errors to return
    Dim errorDict As Dictionary
    Set errorDict = SetUpDictionaries

    'Flag every household, member, and account as not in the list
    FlagAsNotInList "Household"
    FlagAsNotInList "Account"
    
    'Iterate through each row of the MS Household Export list
    Dim rw As Integer
    For rw = 1 To m_datasheet.DataSize - 1
        If rw Mod 100 = 0 Then
            DoEvents
            Debug.Print sheetName & " " & rw & " / " & m_datasheet.DataSize - 1
            Application.StatusBar = sheetName & " " & rw & " / " & m_datasheet.DataSize - 1
        End If
        
        'Update or add the household
        Dim householdNodeList As IXMLDOMNodeList
        Set householdNodeList = UpdateHouseholdNodes(rw, errorDict)
        
        'Update or add the account
        UpdateAccountNodes rw, householdNodeList, errorDict
        
        'Exit the loop if the testing limit has been reached
        If rw = testingLimit Then: rw = m_datasheet.DataSize - 1
    Next rw
    
    'Reset the status bar
    Application.StatusBar = False
    
    Set ImportToXML = errorDict
End Function

Public Function GetData(headerName As String, datapoint As Integer) As String
    GetData = m_datasheet.GetData(headerName, datapoint)
End Function

Public Sub FillWorksheet(sht As Worksheet)
    m_datasheet.FillWorksheet sht
End Sub

Private Function UpdateHouseholdNodes(rw As Integer, errorDict As Dictionary) As IXMLDOMNodeList
    'Get the household from this row
    Dim msHousehold As clsHousehold
    Set msHousehold = GetHouseholdFromRow(rw)
    
    'Set the key for the error dictionaries
    Dim dictKey As String
    dictKey = msHousehold.NameOfHousehold
    
    'Attempt to find the household in the XML
    Dim msHouseholdList As IXMLDOMNodeList
    Set msHouseholdList = XMLRead.FindHouseholds(msHousehold.NameOfHousehold, msHousehold.morningstarID)
    
    'Add the household if it wasn't found, update it if it was
    If msHouseholdList.Length = 0 Then
        'The household wasn't found. Add it to the client list
        AddHouseholdToClientList msHousehold, XMLClientList.SelectSingleNode("Client_List"), errorDict(newHouseholdDictName), dictKey
    Else
        'At least one household was found, update the info for each one
        Dim householdFound As Integer
        For householdFound = 0 To msHouseholdList.Length - 1
            'Update the household node
            UpdateHouseholdNode msHousehold, msHouseholdList(householdFound), errorDict, dictKey
        Next householdFound
        
        'Make note if multiple households were found
        If msHouseholdList.Length > 1 And Not errorDict(dupeHouseholdDictName).Exists(dictKey) Then
            errorDict(dupeHouseholdDictName).Add dictKey, DupeHouseholdDictMessage(msHousehold.NameOfHousehold)
        End If
    End If
    
    'Return the household nodes
    Set UpdateHouseholdNodes = msHouseholdList
End Function

Private Sub UpdateAccountNodes(rw As Integer, householdNodeList As IXMLDOMNodeList, errorDict As Dictionary)
    'Exit the sub if there are no household nodes
    If householdNodeList.Length = 0 Then Exit Sub

    'Get the account from this row
    Dim msAccount As clsAccount
    Set msAccount = GetAccountFromRow(rw)
    
    'Set the key for the error dictionaries
    Dim dictKey As String
    dictKey = msAccount.morningstarID
    
    'Attempt to find the account in the XML
    Dim msAccountList As IXMLDOMNodeList
    Set msAccountList = XMLRead.FindAccounts(accountNumber:=msAccount.Number, accountName:=msAccount.NameOfAccount, morningstarID:=msAccount.morningstarID)
    
    'Put the resulting nodes into an array. If the accounts move, then their updated location in the XML will be preserved
    Dim msAccountArray() As IXMLDOMNode
    If msAccountList.Length = 0 Then
        ReDim msAccountArray(0) As IXMLDOMNode
        Set msAccountArray(0) = Nothing
    Else
        ReDim msAccountArray(0 To msAccountList.Length - 1) As IXMLDOMNode
        Dim accountNode As Integer
        For accountNode = 0 To msAccountList.Length - 1
            Set msAccountArray(accountNode) = msAccountList(accountNode)
        Next accountNode
    End If
    
    'Add/update the account for each household node
    Dim householdNodeFound As Variant
    Dim householdNode As IXMLDOMNode
    For Each householdNodeFound In householdNodeList
        Set householdNode = householdNodeFound
        'Add the account if it wasn't found in the XML, update it if it was
        If msAccountArray(0) Is Nothing Then
            'The account wasn't found. Add it to the household
            AddAccountToHousehold msAccount, householdNode, errorDict(newAccountDictName), dictKey
        Else
            'At least one account was found, update the info for each one
            Dim accountFound As Integer
            For accountFound = 0 To UBound(msAccountArray)
                UpdateAccountInHousehold msAccount, householdNode, msAccountArray(accountFound), errorDict, dictKey
            Next accountFound
            
            'Make note if multiple accounts were found
            If UBound(msAccountArray) > 0 And Not errorDict(dupeAccountDictName).Exists(dictKey) Then
                errorDict(dupeAccountDictName).Add dictKey, DupeAccountDictMessage(msAccount.NameOfAccount, msAccount.Number)
            End If
        End If
    Next householdNodeFound
End Sub

Private Function GetHouseholdFromRow(rw As Integer) As clsHousehold
    'Create a household from the name and ID
    Dim returnHousehold As clsHousehold
    Set returnHousehold = ClassConstructor.NewHousehold(householdName:=m_datasheet.GetData(householdNameHeader, rw), morningstarID:=m_datasheet.GetData(householdIDHeader, rw))
    
    'Member information isn't available in this sheet. Create a dummy member for the account to be added to
    Dim emptyMember As clsMember
    Set emptyMember = ClassConstructor.NewMemberNameOnly(EmptyMemberName)
    
    'Create the account from the information on the row
    Dim accountOnRow As clsAccount
    Set accountOnRow = GetAccountFromRow(rw)
    
    'Add the account to the member, and member to the household
    emptyMember.AddAccount accountOnRow
    returnHousehold.AddMember emptyMember
    
    'Return the household
    Set GetHouseholdFromRow = returnHousehold
End Function

Private Function GetAccountFromRow(rw As Integer) As clsAccount
    'Create an account from the row's information
    Dim returnAccount As clsAccount
    Set returnAccount = ClassConstructor.NewAccount(accountName:=m_datasheet.GetData(accountNameHeader, rw), accountNumber:=m_datasheet.GetData(accountNumberHeader, rw), _
                                                accountType:=m_datasheet.GetData(accountTypeHeader, rw), accountCustodian:=m_datasheet.GetData(custodianHeader, rw), _
                                                accountTag:=vbNullString)
    returnAccount.morningstarID = m_datasheet.GetData(accountIDHeader, rw)
    Set GetAccountFromRow = returnAccount
End Function

Private Function GetEmptyMemberNode(householdNode As IXMLDOMNode) As IXMLDOMNode
    'Find the empty member node in the household
    Dim returnNode As IXMLDOMNode
    Set returnNode = householdNode.SelectSingleNode("./Member[./Full_Name[text()='" & EmptyMemberName & "']]")
    
    If returnNode Is Nothing Then
        'Create an empty member node within the household node
        Dim emptyMember As clsMember
        Set emptyMember = ClassConstructor.NewMemberNameOnly(EmptyMemberName)
        Set returnNode = XMLWrite.AddMemberToNode(emptyMember, householdNode, sheetName)
    End If
    
    'Return the empty member node
    Set GetEmptyMemberNode = returnNode
End Function

Private Sub AddHouseholdToClientList(householdToAdd As clsHousehold, clientListNode As IXMLDOMNode, newHouseholdDict As Dictionary, dictKey As String)
    'The household wasn't found. Add it to the XML
    Dim addedHouseholdNode As IXMLDOMNode
    Set addedHouseholdNode = XMLWrite.AddHouseholdToNode(householdToAdd, clientListNode, sheetName)
    
    'Make note that it's being added
    If Not newHouseholdDict.Exists(dictKey) Then
        newHouseholdDict.Add dictKey, NewHouseholdDictMessage(householdToAdd.NameOfHousehold)
    End If
End Sub

Private Sub AddAccountToHousehold(accountToAdd As clsAccount, householdNode As IXMLDOMNode, newAccountDict As Dictionary, dictKey As String)
    'Get the empty member node in this household (or create it if it doesn't exist)
    Dim emptyMemberNode As IXMLDOMNode
    Set emptyMemberNode = GetEmptyMemberNode(householdNode)
    
    'Add the account to the household node through the empty member
    XMLWrite.AddAccountToNode accountToAdd, emptyMemberNode, sheetName

    'Make note that it's being added, unless the household was already added in the import process
    Dim householdName As String
    householdName = householdNode.SelectSingleNode("./Name").Text
    If Not newAccountDict.Exists(dictKey & householdName) And XMLProcedures.GetAddDate(householdNode) < ProjectGlobals.ImportTime Then
        newAccountDict.Add dictKey & householdName, NewAccountDictMessage(accountToAdd.NameOfAccount, accountToAdd.Number, householdName)
    End If
End Sub

Private Function MoveAccountToCorrectHousehold(accountNode As IXMLDOMNode, householdNode As IXMLDOMNode, accountName As String, accountNumber As String, _
                                            accountID As String, movedAccountDict As Dictionary, dictKey As String) As Boolean
    If Not XMLRead.FindAccounts(accountNumber, morningstarID:=accountID, householdNode:=householdNode) Is Nothing Then
        'The account is already in the correct household
        MoveAccountToCorrectHousehold = False
        Exit Function
    End If
    
    'The account wasn't found in the household node. Add it, remove it from its other parent node, and make note of it
    Dim originalHouseholdName As String, newHouseholdName As String
    originalHouseholdName = accountNode.SelectSingleNode("../../Name").Text
    newHouseholdName = householdNode.SelectSingleNode("./Name").Text
    
    'Move the account node to the correct household
    XMLProcedures.MoveNode accountNode, GetEmptyMemberNode(householdNode)
    
    'Make note that the account moved
    If Not movedAccountDict.Exists(dictKey) Then
        movedAccountDict.Add dictKey, MovedAccountDictMessage(accountName, accountNumber, originalHouseholdName, newHouseholdName)
    End If
    
    'Return that the account was moved
    MoveAccountToCorrectHousehold = True
End Function

Private Sub UpdateHouseholdNode(msHousehold As clsHousehold, householdNode As IXMLDOMNode, errorDict As Dictionary, dictKey As String)
    With msHousehold
        'Update the information in the household node
        Dim listID As String, listName As String
        XMLUpdate.UpdateHouseholdFromMSExport householdNode, .NameOfHousehold, .morningstarID, listID, listName
        
        'Check for differing information in the household
        CheckForDifferingInfo "Morningstar_ID", .morningstarID, listID, errorDict(diffHouseholdIDDictName), dictKey, householdNode, .NameOfHousehold
        CheckForDifferingInfo "Name", .NameOfHousehold, listName, errorDict(diffHouseholdNameDictName), dictKey, householdNode
    End With
End Sub

Private Sub UpdateAccountInHousehold(msAccount As clsAccount, householdNode As IXMLDOMNode, accountNode As IXMLDOMNode, errorDict As Dictionary, dictKey As String)
    'Check that the account is in the correct household
    Dim isAccountMoved As Boolean
    isAccountMoved = MoveAccountToCorrectHousehold(accountNode, householdNode, msAccount.NameOfAccount, msAccount.Number, msAccount.morningstarID, _
                                                   errorDict(movedAccountDictName), dictKey)
    
    'Update the information in the account node
    Dim listAccountID As String, listAccountName As String, listAccountNumber As String, listAccountType As String, listAccountCustodian As String
    With msAccount
        XMLUpdate.UpdateAccountFromMSExport accountNode, .morningstarID, .NameOfAccount, .Number, .TypeOfAccount, .custodian, _
                                            listAccountID, listAccountName, listAccountNumber, listAccountType, listAccountCustodian
        
        'Check for differing information
        CheckForDifferingInfo "Morningstar_ID", .morningstarID, listAccountID, errorDict(diffAccountIDDictName), .morningstarID, accountNode, .Number
        CheckForDifferingInfo "Name", .NameOfAccount, listAccountName, errorDict(diffAccountNameDictName), .morningstarID, accountNode, .Number
        CheckForDifferingInfo "Number", .Number, listAccountNumber, errorDict(diffAccountNumberDictName), .morningstarID, accountNode, .NameOfAccount
        CheckForDifferingInfo "Type", .TypeOfAccount, listAccountType, errorDict(diffAccountTypeDictName), .morningstarID, accountNode, .Number
        CheckForDifferingInfo "Custodian", .custodian, listAccountCustodian, errorDict(diffCustodianDictName), .morningstarID, accountNode, .Number
    End With
End Sub

Private Function SetUpDictionaries() As Dictionary
    'Declare a dictionary to return
    Dim returnDict As Dictionary: Set returnDict = New Dictionary
    Dim errorKey As Variant
    For Each errorKey In ErrorDictKeys
        If Not returnDict.Exists(errorKey) Then
            returnDict.Add errorKey, New Dictionary
        End If
    Next errorKey
    
    'Return the dictionary
    Set SetUpDictionaries = returnDict
End Function

Private Sub FlagAsNotInList(nodeName As String)
    'Add attribute to each node saying it's not in this sheet
    XMLProcedures.FlagNodeTypeInList nodeName, sheetName, "False"
End Sub

Private Sub CheckForDifferingInfo(infoType As String, sheetInfo As Variant, listInfo As Variant, DifferingInfoDict As Dictionary, dictKey As Variant, _
                                  componentNode As IXMLDOMNode, Optional IdentifyingData As String)
    'Make note if the info on this row is different than what's in the list
    XMLProcedures.DifferingInfoCheck infoType, sheetInfo, listInfo, DifferingInfoDict, dictKey, sheetName, componentNode, componentNode.BaseName, IdentifyingData
End Sub

Private Function SearchWrapper(value As String) As String
    'Create a search string with appropriate single or double quotes
    SearchWrapper = XMLProcedures.XPathExpression(value)
End Function
