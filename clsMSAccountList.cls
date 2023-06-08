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
Private Const householdNameHeader As String = "Client / Prospect Name"
Private Const ownerHeader As String = "Account Owner"
Private Const accountTypeHeader As String = "Account Type"
Private Const newHouseholdDictName As String = "New Household"
Private Const newMemberDictName As String = "New Member"
Private Const newAccountDictName As String = "New Account"
Private Const dupeHouseholdDictName As String = "Duplicate Household"
Private Const dupeMemberDictName As String = "Duplicate Member"
Private Const dupeAccountDictName As String = "Duplicate Account"
Private Const diffHouseholdNameDictName As String = "Different Household Name"
Private Const diffMemberFirstNameDictName As String = "Different First Name"
Private Const diffMemberLastNameDictName As String = "Different Last Name"
Private Const diffAccountNameDictName As String = "Different Account Name"
Private Const diffAccountTypeDictName As String = "Different Account Type"
Private Const diffOwnerDictName As String = "Different Owner"
Private Const diffCustodianDictName As String = "Different Custodian"
Private Const diffDiscretionaryDictName As String = "Different Discretionary"
Private Const movedAccountDictName As String = "Moved Account"
Private Const associatedFileLocation As String = "Z:\FPIS - Operations\Beneficiary Project\Assets\associated accounts.txt"

Public Property Get DataSize() As Integer
    DataSize = m_datasheet.DataSize
End Property

Private Property Get ErrorDictKeys() As String()
    Dim keyArray(0 To 14) As String
    keyArray(0) = newHouseholdDictName
    keyArray(1) = newMemberDictName
    keyArray(2) = newAccountDictName
    keyArray(3) = dupeHouseholdDictName
    keyArray(4) = dupeMemberDictName
    keyArray(5) = dupeAccountDictName
    keyArray(6) = diffHouseholdNameDictName
    keyArray(7) = diffMemberFirstNameDictName
    keyArray(8) = diffMemberLastNameDictName
    keyArray(9) = diffAccountNameDictName
    keyArray(10) = diffAccountTypeDictName
    keyArray(11) = diffOwnerDictName
    keyArray(12) = diffCustodianDictName
    keyArray(13) = diffDiscretionaryDictName
    keyArray(14) = movedAccountDictName
    ErrorDictKeys = keyArray
End Property

Private Property Get NewHouseholdDictMessage(householdName As String) As String
    NewHouseholdDictMessage = "Household Added - " & householdName & " with its members and accounts"
End Property

Private Property Get NewMemberDictMessage(memberName As String, householdName As String) As String
    NewMemberDictMessage = "Member Added - " & memberName & " | Household - " & householdName
End Property

Private Property Get NewAccountDictMessage(accountName As String, accountNumber As String, memberName As String, householdName As String) As String
    If accountNumber <> vbNullString Then
        NewAccountDictMessage = "Account Added - " & accountName & " - " & accountNumber & " | Member - " & memberName & " | Household - " & householdName
    Else
        NewAccountDictMessage = "Account Added - " & accountName & " | Member - " & memberName & " | Household - " & householdName
    End If
End Property

Private Property Get DupeHouseholdDictMessage(householdName As String) As String
    DupeHouseholdDictMessage = "Household Duplicate - " & householdName & " - Multiple instances found in XML"
End Property

Private Property Get DupeMemberDictMessage(memberName As String) As String
    DupeMemberDictMessage = "Member Duplicate - " & memberName & " - Multiple instances found in XML"
End Property

Private Property Get DupeAccountDictMessage(accountName As String, accountNumber As String) As String
    DupeAccountDictMessage = "Account Duplicate - " & accountName & " - " & accountNumber & ": Multiple instances found in XML"
End Property

Private Property Get MovedAccountHouseholdDictMessage(accountName As String, accountNumber As String, originalHouseholdName As String, newHouseholdName As String) As String
    MovedAccountHouseholdDictMessage = "Account Moved - " & accountName & " - " & accountNumber & ": from household " & originalHouseholdName & " to " & newHouseholdName
End Property

Private Property Get MovedAccountMemberDictMessage(accountName As String, accountNumber As String, originalMemberName As String, newMemberName As String) As String
    If originalMemberName = EmptyMemberName Then
        MovedAccountMemberDictMessage = "Account Moved - " & accountName & " - " & accountNumber & ": from member " & originalMemberName & " to " & newMemberName
    Else
        MovedAccountMemberDictMessage = "Account Moved - " & accountName & " - " & accountNumber & ": from empty member to " & newMemberName
    End If
End Property

Private Property Get EmptyMemberName() As String
    EmptyMemberName = ProjectGlobals.m_emptyMemberName
End Property

Private Property Get marketValueHeader() As String
    marketValueHeader = "Market Value " & Chr(10) & "USD"
End Property

Private Property Get sheetName() As String
    sheetName = ProjectGlobals.m_msAccountName
End Property

Private Property Get XMLClientList() As DOMDocument60
    Set XMLClientList = ProjectGlobals.ClientListFile
End Property

Public Sub ClassBuilder(path As String)
    'Set the required headers
    Dim requiredHeaders() As String
    requiredHeaders = Split(accountNameHeader & "," & accountNumberHeader & "," & custodianHeader & "," _
                          & householdNameHeader & "," & marketValueHeader & "," & ownerHeader & "," & accountTypeHeader, ",")
    
    'Set the datasheet in order to get the data
    Set m_datasheet = ClassConstructor.NewDataSheet(filePath:=path, wkstName:=worksheetName, reqHeaders:=requiredHeaders)
End Sub

Public Function ImportToXML(Optional testingLimit As Integer) As Dictionary
    'Contains master account names, market values
    'Contains account number, custodian, household name, owner name, account type, discretionary

    'Set up dictionaries for the possible errors to return
    Dim errorDict As Dictionary
    Set errorDict = SetUpDictionaries

    'Flag households, members, and accounts as not in the list
    FlagAsNotInList "Household"
    FlagAsNotInList "Member"
    FlagAsNotInList "Account"
    
    'Iterate through each row of the MS account list
    Dim householdsUpdated As Dictionary: Set householdsUpdated = New Dictionary
    Dim rw As Integer
    For rw = 1 To m_datasheet.DataSize - 1
        If rw Mod 100 = 0 Then
            DoEvents
            Debug.Print sheetName & " " & rw & " / " & m_datasheet.DataSize - 1
            Application.StatusBar = sheetName & " " & rw & " / " & m_datasheet.DataSize - 1
        End If
        
        'Update or add the household
        Dim householdNodeList As IXMLDOMNodeList
        Set householdNodeList = UpdateHouseholdNodes(rw, errorDict, householdsUpdated)
        
        'Update or add the member
        Dim memberNodeList As IXMLDOMNodeList
        Set memberNodeList = UpdateMemberNodes(rw, householdNodeList, errorDict)
        
        'Update or add the account
        UpdateAccountNodes rw, memberNodeList, errorDict
        
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

Private Function GetAccountFromRow(rw As Integer) As clsAccount
    Dim returnAccount As clsAccount
    Set returnAccount = ClassConstructor.NewAccount(accountName:=m_datasheet.GetData(accountNameHeader, rw), accountNumber:=m_datasheet.GetData(accountNumberHeader, rw), _
                                                    accountType:=m_datasheet.GetData(accountTypeHeader, rw), accountCustodian:=m_datasheet.GetData(custodianHeader, rw), _
                                                    accountTag:=vbNullString, marketValue:=Val(GetData(marketValueHeader, rw)))
    returnAccount.owner = ClassConstructor.NewMemberNameOnly(GetData(ownerHeader, rw))
    Set GetAccountFromRow = returnAccount
End Function

Private Function UpdateHouseholdNodes(rw As Integer, errorDict As Dictionary, householdsUpdated As Dictionary) As IXMLDOMNodeList
    'Get the household from this row
    Dim msHousehold As clsHousehold
    Set msHousehold = GetFullHouseholdFromRow(rw)
    
    'Set the key for the error dictionaries
    Dim dictKey As String
    dictKey = msHousehold.NameOfHousehold
    
    'Attempt to find the household in the XML
    Dim msHouseholdList As IXMLDOMNodeList
    Set msHouseholdList = XMLRead.FindHouseholds(msHousehold.NameOfHousehold)
    
    'Add the household if it wasn't found, update it if it was
    If msHouseholdList.Length = 0 Then
        'The household wasn't found. Add it to the XML
        AddHouseholdToClientList msHousehold, XMLClientList.SelectSingleNode("Client_List"), errorDict(newHouseholdDictName), dictKey
    Else
        'At least one household was found, update the info for each one if it hasn't yet been updated
        If Not householdsUpdated.Exists(dictKey) Then
            'The household hasn't been updated yet. Update it
            Dim householdFound As Integer
            For householdFound = 0 To msHouseholdList.Length - 1
                'Update the household
                UpdateHouseholdNode msHousehold, msHouseholdList(householdFound), errorDict, dictKey
            Next householdFound
            
            If msHouseholdList.Length > 1 And Not errorDict(dupeHouseholdDictName).Exists(dictKey) Then
                'Make note that multiple households were found
                errorDict(dupeHouseholdDictName).Add dictKey, DupeHouseholdDictMessage(msHousehold.NameOfHousehold)
            End If
            
            'Add the household to the updated households dictionary
            householdsUpdated.Add dictKey, dictKey
        End If
    End If
    
    'Return the household nodes
    Set UpdateHouseholdNodes = msHouseholdList
End Function

Private Function UpdateMemberNodes(rw As Integer, householdNodeList As IXMLDOMNodeList, errorDict As Dictionary) As IXMLDOMNodeList
    If householdNodeList.Length = 0 Then
        'There are no household nodes. Return a list of 0 and exit the function
        Set UpdateMemberNodes = XMLClientList.createNode(1, "Household", vbNullString).SelectNodes("not //*")
        Exit Function
    End If
    
    'Get the member from this row
    Dim msMember As clsMember
    Set msMember = ClassConstructor.NewMemberNameOnly(GetData(ownerHeader, rw))
    
    'Set the key for the error dictionaries
    Dim dictKey As String
    If msMember.fName <> vbNullString And msMember.lName <> vbNullString Then
        dictKey = msMember.fName & msMember.lName
    Else
        dictKey = msMember.NameOfMember
    End If
    
    'Add/update the member for each household node
    Dim householdNodeFound As Variant
    Dim householdNode As IXMLDOMNode
    For Each householdNodeFound In householdNodeList
        'Attempt to find the member in the XML
        Set householdNode = householdNodeFound
        Dim msMemberList As IXMLDOMNodeList
        Set msMemberList = XMLRead.FindMembersInHousehold(householdNode:=householdNode, fName:=msMember.fName, lName:=msMember.lName, _
                                                          fullName:=msMember.NameOfMember)

        'Add the member if it wasn't found, update it if it was
        If msMemberList.Length = 0 Then
            'The member wasn't found. Add it to the XML under the household node
            AddMemberToHousehold msMember, householdNode, errorDict(newMemberDictName), dictKey
        Else
            'At least one member was found, update the info for each one
            Dim memberFound As Integer
            For memberFound = 0 To msMemberList.Length - 1
                'Update the member
                UpdateMemberNode msMember, msMemberList(memberFound), errorDict, dictKey
            Next memberFound

            If msMemberList.Length > 1 And Not errorDict(dupeMemberDictName).Exists(dictKey) Then
                'Make note that multiple members were found
                errorDict(dupeMemberDictName).Add dictKey, DupeMemberDictMessage(msMember.NameOfMember)
            End If
        End If
    Next householdNodeFound
    
    'Return the member nodes
    Set UpdateMemberNodes = msMemberList
End Function

Private Sub UpdateAccountNodes(rw As Integer, memberNodeList As IXMLDOMNodeList, errorDict As Dictionary)
    'Exit the sub if there are no member nodes
    If memberNodeList.Length = 0 Then Exit Sub
    
    'Get the account from this row
    Dim msAccount As clsAccount
    Set msAccount = GetAccountFromRow(rw)

    'Set the key for the dictionaries
    Dim dictKey As String
    dictKey = msAccount.NameOfAccount & msAccount.Number

    'Attempt to find the account(s) in the XML by its name/number
    Dim msAccountList As IXMLDOMNodeList
    Set msAccountList = XMLRead.FindAccounts(msAccount.Number, msAccount.NameOfAccount)
    
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

    'Add/update the account for each member node
    Dim memberNodeFound As Variant
    Dim memberNode As IXMLDOMNode
    For Each memberNodeFound In memberNodeList
        Set memberNode = memberNodeFound
        If msAccountArray(0) Is Nothing Then
            'The account wasn't found, add it as its own node in the member
            AddAccountToMember msAccount, memberNode, errorDict(newAccountDictName), dictKey
        Else
            'At least one account node was found, update the info for each one
            Dim nodeFound As Integer
            For nodeFound = 0 To UBound(msAccountArray)
                UpdateAccountInMember msAccount, memberNode, msAccountArray(nodeFound), errorDict, dictKey
            Next nodeFound
    
            'Check for duplicate accounts
            If UBound(msAccountArray) > 0 And Not errorDict.Item(dupeAccountDictName).Exists(dictKey) Then
                'Make note that there are multiple account nodes with the same name/number
                errorDict.Item(dupeAccountDictName).Add dictKey, DupeAccountDictMessage(msAccount.NameOfAccount, msAccount.Number)
            End If
        End If
    Next memberNodeFound
End Sub

Private Function GetFullHouseholdFromRow(rw As Integer) As clsHousehold
    'Create a household from the name and ID
    Dim returnHousehold As clsHousehold
    Set returnHousehold = ClassConstructor.NewHousehold(householdName:=m_datasheet.GetData(householdNameHeader, rw))
    
    'Member information isn't available in this sheet. Create a dummy member for the account to be added to
    Dim householdMember As clsMember
    Set householdMember = ClassConstructor.NewMemberNameOnly(GetData(ownerHeader, rw))
    
    'Create the account from the information on the row
    Dim accountOnRow As clsAccount
    Set accountOnRow = GetAccountFromRow(rw)
    
    'Add the account to the member, and member to the household
    householdMember.AddAccount accountOnRow
    returnHousehold.AddMember householdMember
    
    'Return the household
    Set GetFullHouseholdFromRow = returnHousehold
End Function

Private Function GetHouseholdFromRow(households As Dictionary, rw As Integer) As clsHousehold
    Dim householdName As String
    householdName = GetData(householdNameHeader, rw)
    
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

Private Function GetMemberFromRow(households As Dictionary, rw As Integer) As clsMember
    'Get the household to look at
    Dim tempHousehold As clsHousehold
    Set tempHousehold = GetHouseholdFromRow(households, rw)
    
    If Not tempHousehold Is Nothing Then
        'Get the member
        Dim acctOwner As String
        acctOwner = GetData(ownerHeader, rw)
        Dim tempmember As clsMember
        Set tempmember = tempHousehold.GetMemberFromLastCommaFirst(acctOwner)
        If tempmember Is Nothing Then
            'Add a new member to the household
            Set tempmember = ClassConstructor.NewMemberNameOnly(acctOwner)
            tempHousehold.AddMember tempmember
        End If
    
        Set GetMemberFromRow = tempmember
    End If
End Function

Private Sub AddHouseholdToClientList(householdToAdd As clsHousehold, clientListNode As IXMLDOMNode, newHouseholdDict As Dictionary, dictKey As String)
    'Add the household to the client list node
    Dim addedHouseholdNode As IXMLDOMNode
    Set addedHouseholdNode = XMLWrite.AddHouseholdToNode(householdToAdd, clientListNode, sheetName)
    
    'Make note that it's being added
    If Not newHouseholdDict.Exists(dictKey) Then
        newHouseholdDict.Add dictKey, NewHouseholdDictMessage(householdToAdd.NameOfHousehold)
    End If
End Sub

Private Sub AddMemberToHousehold(msMember As clsMember, householdNode As IXMLDOMNode, newMemberDict As Dictionary, dictKey As String)
    'Add the member to the household node
    XMLWrite.AddMemberToNode msMember, householdNode, sheetName

    'Make note that it's being added, unless the household was already added in the import process
    Dim householdName As String
    householdName = householdNode.SelectSingleNode("./Name").Text
    If Not newMemberDict.Exists(dictKey & householdName) And XMLProcedures.GetAddDate(householdNode) < ProjectGlobals.ImportTime Then
        newMemberDict.Add dictKey & householdName, NewMemberDictMessage(msMember.NameOfMember, householdName)
    End If
End Sub

Private Sub AddAccountToMember(msAccount As clsAccount, memberNode As IXMLDOMNode, newAccountDict As Dictionary, dictKey As String)
    'Add the account to the member node
    XMLWrite.AddAccountToNode msAccount, memberNode, sheetName

    'Make note that the account was added unless the member was already added in the process
    Dim householdName As String
    householdName = memberNode.SelectSingleNode("../Name").Text
    If Not newAccountDict.Exists(dictKey) And XMLProcedures.GetAddDate(memberNode) < ProjectGlobals.ImportTime Then
        newAccountDict.Add dictKey, NewAccountDictMessage(msAccount.NameOfAccount, msAccount.Number, msAccount.owner.NameOfMember, householdName)
    End If
End Sub

Private Function MoveAccountToCorrectMember(accountNode As IXMLDOMNode, memberNode As IXMLDOMNode, accountName As String, accountNumber As String, _
                                            movedAccountDict As Dictionary, dictKey As String) As Boolean
    If Not XMLRead.FindAccountInMember(memberNode, accountNumber, accountName) Is Nothing Then
        'The account is already in the correct member
        MoveAccountToCorrectMember = False
        Exit Function
    End If
    
    'Add the account to the member node, remove it from its other parent node, and make note of it
    'Get the old and new member names
    Dim originalMemberName As String, newMemberName As String
    If accountNode.parentNode Is Nothing Then
        'The account's parent doesn't exist. Return nothing
        originalMemberName = vbNullString
    ElseIf accountNode.SelectSingleNode("..").BaseName <> "Member" Then
        'The account's parent isn't a member. Return nothing
        originalMemberName = vbNullString
    ElseIf accountNode.SelectSingleNode("../Full_Name") Is Nothing Then
        'The member exists and doesn't have a full name node. Concatenate the last and first names, separated by a comma
        originalMemberName = accountNode.SelectSingleNode("../Last_Name").Text & ", " & accountNode.SelectSingleNode("../First_Name").Text
    Else
        'The member has a full name node. Return it
        originalMemberName = accountNode.SelectSingleNode("../Full_Name").Text
    End If
    newMemberName = memberNode.SelectSingleNode("./Full_Name").Text
    
    'Get the old and new household names
    Dim originalHouseholdName As String, newHouseholdName As String
    If accountNode.parentNode Is Nothing Then
        'The account has no parent. Return nothing
        originalHouseholdName = vbNullString
    ElseIf accountNode.SelectSingleNode("../..").BaseName <> "Household" Then
        'The account's grandparent isn't a household. Return nothing
        originalHouseholdName = vbNullString
    Else
        'The household node exists. Get the household name
        originalHouseholdName = accountNode.SelectSingleNode("../../Name").Text
    End If
    newHouseholdName = memberNode.SelectSingleNode("../Name").Text
    
    'Change the account's owner to be the member name
    accountNode.SelectSingleNode("Owner").Text = newMemberName

    'Move the account node to the correct member
    Set accountNode = XMLProcedures.MoveNode(accountNode, memberNode)

    'Make note that the account moved
    If Not movedAccountDict.Exists(dictKey) Then
        If originalHouseholdName <> newHouseholdName And XMLProcedures.GetAddDate(memberNode.parentNode) < ProjectGlobals.ImportTime Then
            'The account is changing households and the household wasn't yet added in the process
            movedAccountDict.Add dictKey, MovedAccountHouseholdDictMessage(accountName, accountNumber, originalHouseholdName, newHouseholdName)
        ElseIf originalHouseholdName = newHouseholdName Then
            'The account is changing members within a household
            movedAccountDict.Add dictKey, MovedAccountMemberDictMessage(accountName, accountNumber, originalMemberName, newMemberName)
        Else
            'The account is changing households, but the household was already added in the process
        End If
    End If
    
    'Return that the account was moved
    MoveAccountToCorrectMember = True
End Function

Private Sub UpdateHouseholdNode(msHousehold As clsHousehold, householdNode As IXMLDOMNode, errorDict As Dictionary, dictKey As String)
    With msHousehold
        'Update the information in the household node
        Dim listName As String
        XMLUpdate.UpdateHouseholdFromMSAccounts householdNode, .NameOfHousehold, listName
        
        'Check for differing information in the household unless the household was already added in the process
        CheckForDifferingInfo "Name", .NameOfHousehold, listName, errorDict(diffHouseholdNameDictName), dictKey, householdNode
    End With
End Sub

Private Sub UpdateMemberNode(msMember As clsMember, memberNode As IXMLDOMNode, errorDict As Dictionary, dictKey As String)
    With msMember
        'Update the information in the member node
        Dim listFirstName As String, listLastName As String
        XMLUpdate.UpdateMemberFromMSAccounts memberNode, .fName, .lName, .NameOfMember, listFirstName, listLastName
    
        'Check for differing information in the member
        CheckForDifferingInfo "First_Name", .fName, listFirstName, errorDict(diffMemberFirstNameDictName), dictKey, memberNode, .NameOfMember
        CheckForDifferingInfo "Last_Name", .lName, listLastName, errorDict(diffMemberLastNameDictName), dictKey, memberNode, .NameOfMember
    End With
End Sub

Private Function UpdateAccountInMember(msAccount As clsAccount, memberNode As IXMLDOMNode, accountNode As IXMLDOMNode, errorDict As Dictionary, dictKey As String) As IXMLDOMNode
    'Check that the account is in the correct member
    Dim isAccountMoved As Boolean
    isAccountMoved = MoveAccountToCorrectMember(accountNode, memberNode, msAccount.NameOfAccount, msAccount.Number, errorDict(movedAccountDictName), dictKey)
    
    'Update the account information
    Dim listAccountName As String, listAccountType As String, listOwnerName As String, listCustodian As String
    Dim listDiscretionary As Boolean
    With msAccount
        XMLUpdate.UpdateAccountFromMSAccountList accountNode, .NameOfAccount, .TypeOfAccount, .Balance, .owner.NameOfMember, .Discretionary, _
                                                 .custodian, listAccountName, listAccountType, listOwnerName, listDiscretionary, listCustodian

        'Check for differing information
        Dim IdentifyingData As String
        IdentifyingData = msAccount.NameOfAccount & " - " & msAccount.Number
        CheckForDifferingInfo "Name", .NameOfAccount, listAccountName, errorDict(diffAccountNameDictName), dictKey, accountNode, .Number
        CheckForDifferingInfo "Type", .TypeOfAccount, listAccountType, errorDict(diffAccountTypeDictName), dictKey, accountNode, IdentifyingData
        CheckForDifferingInfo "Custodian", .custodian, listCustodian, errorDict(diffCustodianDictName), dictKey, accountNode, IdentifyingData
        CheckForDifferingInfo "Discretionary", .Discretionary, listDiscretionary, errorDict(diffDiscretionaryDictName), dictKey, accountNode, IdentifyingData
        If Not isAccountMoved Then
            'Don't check the owner if the account moved to a different member. If it moved, then this will only create a redundant error message
            CheckForDifferingInfo "Owner", .owner.NameOfMember, listOwnerName, errorDict(diffOwnerDictName), dictKey, accountNode, IdentifyingData
        End If
    End With
    
    Set UpdateAccountInMember = accountNode
End Function

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

Private Sub CheckForDifferingInfo(infoType As String, sheetInfo As Variant, listInfo As Variant, DifferingInfoDict As Dictionary, dictKey As Variant, _
                                  componentNode As IXMLDOMNode, Optional IdentifyingData As String)
    'Make note if the info on this row is different than what's in the list
    XMLProcedures.DifferingInfoCheck infoType, sheetInfo, listInfo, DifferingInfoDict, dictKey, sheetName, componentNode, componentNode.BaseName, IdentifyingData
End Sub

Private Sub FlagAsNotInList(nodeName As String)
    'Add attribute to each node saying it's not in this sheet
    XMLProcedures.FlagNodeTypeInList nodeName, sheetName, "False"
End Sub

Private Function SearchWrapper(value As String) As String
    'Create a search string with appropriate single or double quotes
    SearchWrapper = XMLProcedures.XPathExpression(value)
End Function
