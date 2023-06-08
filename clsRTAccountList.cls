VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsRTAccountList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'ImportToXML and all its associated subs/functions are not ready for use and not planned on being worked on
'RT Accounts sheet doesn't appear to add necessary information.
'It contains account owner, number, custodian, and type, but these don't need to be closely monitored and updated in Redtail
'Consider removing
Private m_datasheet As clsDataSheet
Private Const worksheetName As String = "RT Accounts"
Private Const idHeader As String = "Account Id" '(sic)
Private Const nameHeader As String = "Contact Name"
Private Const accountNumberHeader As String = "Account Number"
Private Const accountTypeHeader As String = "Type"
Private Const custodianHeader As String = "Company"
Private Const newMemberDictName As String = "New Member"
Private Const newAccountDictName As String = "New Account"
Private Const dupeAccountDictName As String = "Duplicate Account"
Private Const diffAccountTypeDictName As String = "Different Account Type"
Private Const diffOwnerDictName As String = "Different Owner"
Private Const diffCustodianDictName As String = "Different Custodian"

Private Property Get ErrorDictKeys() As String()
    Dim keyArray(0 To 5) As String
    keyArray(0) = newMemberDictName
    keyArray(1) = newAccountDictName
    keyArray(2) = dupeAccountDictName
    keyArray(3) = diffOwnerDictName
    keyArray(4) = diffAccountTypeDictName
    keyArray(5) = diffCustodianDictName
    ErrorDictKeys = keyArray
End Property

Public Property Get DataSize() As Integer
    DataSize = m_datasheet.DataSize
End Property

Private Property Get sheetName() As String
    sheetName = ProjectGlobals.m_rtAccountName
End Property

Private Property Get XMLClientList() As DOMDocument60
    Set XMLClientList = ProjectGlobals.ClientListFile
End Property

Public Sub ClassBuilder(path As String)
    'Set the required headers
    Dim requiredHeaders() As String
    requiredHeaders = Split(idHeader & "," & nameHeader & "," & accountNumberHeader & "," & accountTypeHeader & "," & custodianHeader, ",")
    
    'Set the datasheet in order to get the data
    Set m_datasheet = ClassConstructor.NewDataSheet(filePath:=path, wkstName:=worksheetName, reqHeaders:=requiredHeaders)
End Sub

Public Function ImportToXML(Optional testingLimit As Integer) As Dictionary
    'Contains master Redtail Account ID
    'Contains owner name, account number, custodian, account type, account active status, market value
    
    'Set up dictionaries for the possible errors to return
    Dim errorDict As Dictionary
    Set errorDict = SetUpDictionaries

    'Flag every household, member, and account as not in the list
    FlagAsNotInList "Household"
    FlagAsNotInList "Member"
    FlagAsNotInList "Account"
    
    'Iterate through each row of the RT account list
    Dim rw As Integer
    For rw = 1 To m_datasheet.DataSize - 1
        If rw Mod 100 = 0 Then
            DoEvents
            Debug.Print sheetName & " " & rw & " / " & m_datasheet.DataSize - 1
            Application.StatusBar = sheetName & " " & rw & " / " & m_datasheet.DataSize - 1
        End If
        
        'Update or add the account
        Dim accountNodeList As IXMLDOMNodeList
'        Set accountNodeList = UpdateAccountNodes(rw, memberNodeList, errordict)
        
        'Update or add the member
'        UpdateMemberNodes rw, householdNodeList, errordict
        
        'Exit the loop if the testing limit has been reached
        If rw = testingLimit Then: rw = m_datasheet.DataSize - 1
    Next rw
    
    'Reset the status bar
    Application.StatusBar = False
    
    Set ImportToXML = errorDict
End Function

Private Function UpdateAccountNodes(rw As Integer, memberNodeList As IXMLDOMNodeList, errorDict As Dictionary) As IXMLDOMNodeList
'    'Get the account from this row
'    Dim rtAccount As clsAccount
'    Set rtAccount = GetAccountFromRow(rw)
'
'    'Attempt to find the account(s) in the XML by its Redtail ID, number, and owner name
'    Dim accountNodeList As IXMLDOMNodeList
'    Set accountNodeList = GetAccountsFromXML(rtAccount.redtailID, rtAccount.Number)
'
'    If accountNodeList.Length = 0 Then
'        'The account wasn't found. Attempt to find the owner node
''        Dim memberNodeList As IXMLDOMNodeList
''        Set memberNodeList = GetMembersFromXML(rtAccount.owner.lName, rtAccount.owner.fName, rtAccount.owner.Nickname)
'
'        If memberNodeList.Length = 0 Then
'            'The member wasn't found. Add it as its own node to the client list
'            XMLWrite.AddMemberToNode rtAccount.owner, XMLClientList.SelectSingleNode("Client_List"), sheetName
'
'            'Make note that the member with the account was added
'            If Not errorDict.Item(newMemberDictName).Exists(rtAccount.owner.NameOfMember) Then
'                errorDict.Item(newMemberDictName).Add rtAccount.owner.NameOfMember, rtAccount.owner.NameOfMember & ": Added to XML"
'            End If
'        Else
'            'At least one member was found. Add the account to each one
'            Dim memberNodeItem As Variant
'            For Each memberNodeItem In memberNodeList
'                Dim memberNodeFound As IXMLDOMElement: Set memberNodeFound = memberNodeItem
'                XMLWrite.AddAccountToNode rtAccount, memberNodeFound, sheetName
'
'                'Make note that the account was added
'                If Not errorDict.Item(newAccountDictName).Exists(rtAccount.redtailID) Then
'                    errorDict.Item(newAccountDictName).Add rtAccount.redtailID, rtAccount.Number & ": Added to XML"
'                End If
'            Next memberNodeItem
'        End If
'    Else
'        'At least one account was found. Update the info for each one
'        Dim accountNodeItem As Variant
'        For Each accountNodeItem In accountNodeList
'            Dim accountNodeFound As IXMLDOMElement: Set accountNodeFound = accountNodeItem
'            Dim listOwnerName As String, listAccountType As String, listCustodian As String, listAccountNumber As String
'            With rtAccount
'                UpdateAccountFromRTAccountList accountNodeFound, .redtailID, .TypeOfAccount, .custodian, .Number, .redtailID, listAccountType, listCustodian, listAccountNumber
'            End With
'
'            'Check for differing information
'            CheckForDifferingInfo "Owner", rtAccount.owner.NameOfMember, listOwnerName, errorDict.Item(diffOwnerDictName), accountNodeFound, rtAccount.redtailID
'            CheckForDifferingInfo "Type", rtAccount.TypeOfAccount, listAccountType, errorDict.Item(diffAccountTypeDictName), accountNodeFound, rtAccount.redtailID
'            CheckForDifferingInfo "Custodian", rtAccount.custodian, listCustodian, errorDict.Item(diffCustodianDictName), accountNodeFound, rtAccount.redtailID
'        Next accountNodeItem
'    End If
'
'    If accountNodeList.Length > 1 And Not errorDict.Item(dupeAccountDictName).Exists(rtAccount.redtailID) Then
'        'Make note that there are multiple account nodes with the same name/number
'        errorDict.Item(dupeAccountDictName).Add rtAccount.redtailID, "RT Account ID: " & rtAccount.redtailID & " - Multiple instances found in XML"
'    End If
End Function

Private Sub UpdateMemberNodes(rw As Integer, accountNodeList As IXMLDOMNodeList, errorDict As Dictionary)
'    'Get the member from this row
'    Dim msMember As clsMember
'    Set msMember = ClassConstructor.NewMemberNameOnly(GetData(nameHeader, rw))
'
'    'Set the key for the error dictionaries
'    Dim dictKey As String
'    dictKey = msMember.fName & msMember.lName
'
'    'Attempt to find the member in the XML
'    Dim msMemberList As IXMLDOMNodeList
'    If accountNodeList.Length > 0 Then
'        'Account nodes were found. Attempt to find the member
'        Set msMemberList = GetMembersFromXML(msMember.fName, msMember.lName, msMember.NameOfMember, accountNodeList(0))
'    Else
'        'No account nodes were found. Need to send a dummy account node in order to return a list of length 0
'        Set msMemberList = GetMembersFromXML(msMember.fName, msMember.lName, msMember.NameOfMember, ProjectGlobals.ClientListFile.createNode(1, "Account", vbNullString))
'    End If
'
'    'Add/update the member for each account node
'    Dim accountNodeFound As Variant
'    Dim accountNode As IXMLDOMNode
'    For Each accountNodeFound In accountNodeList
'        Set accountNode = accountNodeFound
'
'        'Add the member if it wasn't found, update it if it was
'        If msMemberList.Length = 0 Then
'            'The member wasn't found. Add it to the XML under the household node
''            XMLWrite.AddMemberToNode msMember, householdNode, sheetName
'
'            'Make note that it's being added, unless the household was already added in the import process
'            Dim householdName As String
''            householdName = householdNode.SelectSingleNode("./Name").Text
''            If Not errordict(newAccountDictName).Exists(dictKey & householdName) And GetAddDate(householdNode) < ProjectGlobals.ImportTime Then
''                errordict(newMemberDictName).Add dictKey, NewMemberDictMessage(msMember.lName, msMember.fName, householdName)
''            End If
'        Else
'            'At least one member was found, update the info for each one
'            Dim memberFound As Integer
'            For memberFound = 0 To msMemberList.Length - 1
'                'Update the member
'                Dim listFirstName As String, listLastName As String
''                XMLUpdate.UpdateMemberFromRTAccounts msMemberList(memberFound), msMember.fName, msMember.lName, msMember.NameOfMember, listFirstName, listLastName
'
'                'Check for differing information in the member
'                CheckForDifferingInfo "First_Name", msMember.fName, listFirstName, errordict(diffMemberFirstNameDictName), dictKey, "Member", msMember.NameOfMember
'                CheckForDifferingInfo "Last_Name", msMember.lName, listLastName, errordict(diffMemberLastNameDictName), dictKey, "Member", msMember.NameOfMember
'            Next memberFound
'
'            If msMemberList.Length > 1 And Not errordict(dupeMemberDictName).Exists(dictKey) Then
'                'Make note that multiple members were found
'                errordict(dupeMemberDictName).Add dictKey, DupeMemberDictMessage(msMember.NameOfMember)
'            End If
'        End If
'    Next accountNodeFound
'
'    'Return the member nodes
'    Set UpdateMemberNodes = msMemberList
End Sub

Private Function GetAccountFromRow(rw As Integer) As clsAccount
    Dim returnAccount As clsAccount
    Set returnAccount = ClassConstructor.NewAccount(accountName:=vbNullString, accountNumber:=m_datasheet.GetData(accountNumberHeader, rw), _
                                                    accountType:=m_datasheet.GetData(accountTypeHeader, rw), accountCustodian:=m_datasheet.GetData(custodianHeader, rw), _
                                                    accountTag:=vbNullString)
    returnAccount.redtailID = m_datasheet.GetData(idHeader, rw)
    returnAccount.owner = ClassConstructor.NewMemberNameOnly(m_datasheet.GetData(nameHeader, rw))
    Set GetAccountFromRow = returnAccount
End Function

Private Function GetAccountsFromXML(redtailID As Long, accountNumber As String) As IXMLDOMNodeList
'    'Declare a list to return
'    Dim accountNodeList As IXMLDOMNodeList
'
'    'Attempt to find the account
'    Dim attempt As Integer: attempt = 0
'    Do
'        attempt = attempt + 1
'        Select Case attempt
'        Case 1
'            'Attempt to find the account by Redtail ID
'            Set accountList = XMLClientList.SelectNodes("//Account[./Redtail_ID[text()='" & redtailID & "']]")
'        Case 2
'            If accountNumber <> "-" Then
'                'Attempt to find the account by number
'                Set accountList = XMLClientList.SelectNodes("//Account[./Number[text()='" & accountNumber & "']]")
'            End If
'        End Select
'    Loop While accountNodeList.Length = 0 And attempt <= 2
'
'    'Return the nodes found
'    Set GetAccountsFromXML = accountNodeList
End Function

Private Function GetMembersFromXML(lastName As String, firstName As String, nickname As String, accountNode As IXMLDOMNode) As IXMLDOMNodeList
    'Attempt to find the member nodes by first(or nickname) and last name
    Dim memberList As IXMLDOMNodeList
    If accountNode.parentNode Is Nothing Then
        Set memberList = XMLClientList.SelectNodes("//Member[./Last_Name[text()=" & SearchWrapper(lastName) & "] and (./First_Name[text()=" _
                         & SearchWrapper(firstName) & "] or ./First_Name[text()=" & SearchWrapper(nickname) & "])]")
    Else
        Set memberList = accountNode.SelectNodes("../[./Last_Name[text()=" & SearchWrapper(lastName) & "] and (./First_Name[text()=" _
                         & SearchWrapper(firstName) & "] or ./First_Name[text()=" & SearchWrapper(nickname) & "])]")
    End If
    
    'Return the nodes found
    Set GetMembersFromXML = memberList
End Function

Public Function GetData(headerName As String, datapoint As Integer) As String
    GetData = m_datasheet.GetData(headerName, datapoint)
End Function

Public Sub FillWorksheet(sht As Worksheet)
    m_datasheet.FillWorksheet sht
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

Private Sub CheckForDifferingInfo(infoType As String, sheetInfo As Variant, listInfo As Variant, DifferingInfoDict As Dictionary, dictKey As Variant, _
                                  dataNode As IXMLDOMNode, Optional IdentifyingData As String)
    'Make note if the info on this row is different than what's in the list
    XMLProcedures.DifferingInfoCheck infoType, sheetInfo, listInfo, DifferingInfoDict, dictKey, sheetName, dataNode, dataNode.DataType, IdentifyingData
End Sub

Private Sub FlagAsNotInList(nodeName As String)
    'Add attribute to each node saying it's not in this sheet
    XMLProcedures.FlagNodeTypeInList nodeName, sheetName, "False"
End Sub

Private Function SearchWrapper(value As String) As String
    'Create a search string with appropriate single or double quotes
    SearchWrapper = XMLProcedures.XPathExpression(value)
End Function
