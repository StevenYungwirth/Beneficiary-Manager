VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsRTContactList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_datasheet As clsDataSheet
Private Const worksheetName As String = "RT Contacts"
Private Const idHeader As String = "ID"
Private Const typeHeader As String = "Type"
Private Const householdNameHeader As String = "Family Name"
Private Const firstNameHeader As String = "First Name"
Private Const middleNameHeader As String = "Middle Name"
Private Const lastNameHeader As String = "Last Name"
Private Const suffixHeader As String = "Suffix"
Private Const nicknameHeader As String = "Nickname"
Private Const statusHeader As String = "Status"
Private Const categoryHeader As String = "Category"
Private Const dateOfDeathHeader As String = "Date Of Death"
Private Const allowedMemberType As String = "Individual"
Private Const newHouseholdDictName As String = "New Household"
Private Const newMemberDictName As String = "New Member"
Private Const dupeHouseholdDictName As String = "Duplicate Household"
Private Const dupeMemberDictName As String = "Duplicate Member"
Private Const diffFirstNameDictName As String = "Different First Name"
Private Const diffLastNameDictName As String = "Different Last Name"
Private Const diffStatusDictName As String = "Different Status"
Private Const diffDeathDateDictName As String = "Different Date of Death"
Private Const diffHouseholdNameDictName As String = "Different Household Name"

Public Property Get DataSize() As Integer
    DataSize = m_datasheet.DataSize
End Property

Private Property Get ErrorDictKeys() As String()
    Dim keyArray(0 To 8) As String
    keyArray(0) = newHouseholdDictName
    keyArray(1) = newMemberDictName
    keyArray(2) = dupeHouseholdDictName
    keyArray(3) = dupeMemberDictName
    keyArray(4) = diffFirstNameDictName
    keyArray(5) = diffLastNameDictName
    keyArray(6) = diffStatusDictName
    keyArray(7) = diffDeathDateDictName
    keyArray(8) = diffHouseholdNameDictName
    ErrorDictKeys = keyArray
End Property

Private Property Get NewHouseholdDictMessage(householdName As String) As String
    NewHouseholdDictMessage = "Household Added - " & householdName & " with its members"
End Property

Private Property Get NewMemberDictMessage(lastName As String, firstName As String, householdName As String) As String
    NewMemberDictMessage = "Member Added - " & lastName & ", " & firstName & " | Household - " & householdName
End Property

Private Property Get DupeHouseholdDictMessage(householdName As String) As String
    DupeHouseholdDictMessage = "Household Duplicate - " & householdName & " - Multiple instances found in XML"
End Property

Private Property Get DupeMemberDictMessage(memberName As String) As String
    DupeMemberDictMessage = "Member Duplicate - " & memberName & " - Multiple instances found in XML"
End Property

Private Property Get sheetName() As String
    sheetName = ProjectGlobals.m_rtContactName
End Property

Private Property Get XMLClientList() As DOMDocument60
    Set XMLClientList = ProjectGlobals.ClientListFile
End Property

Public Sub ClassBuilder(path As String)
    'Set the required headers
    Dim requiredHeaders() As String
    requiredHeaders = Split(idHeader & "," & householdNameHeader & "," & firstNameHeader & "," & middleNameHeader & "," & lastNameHeader & "," _
                      & suffixHeader & "," & nicknameHeader & "," & statusHeader & "," & categoryHeader & "," & dateOfDeathHeader, ",")
    
    'Set the datasheet in order to get the data
    Set m_datasheet = ClassConstructor.NewDataSheet(filePath:=path, wkstName:=worksheetName, reqHeaders:=requiredHeaders)
End Sub

Public Function ImportToXML(Optional testingLimit As Integer) As Dictionary
    'Contains master Redtail Contact ID
    'Contains household name, member name, member active status, date of death
    
    'Set up dictionaries for the possible errors to return
    Dim errorDict As Dictionary
    Set errorDict = SetUpDictionaries
    
    'Flag every household and member as not in the list
    FlagAsNotInList "Household"
    FlagAsNotInList "Member"
    
    'Iterate through each row of the RT contact list
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
        UpdateMemberNodes rw, householdNodeList, errorDict
        
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

Private Function UpdateHouseholdNodes(rw As Integer, errorDict As Dictionary, householdsUpdated As Dictionary) As IXMLDOMNodeList
    'Get the household from this row
    Dim rtHousehold As clsHousehold
    Set rtHousehold = GetHouseholdFromRow(rw)
    
    'Set the key for the error dictionaries
    Dim dictKey As String
    dictKey = rtHousehold.NameOfHousehold
    
    'Attempt to find the household in the XML
    Dim householdNodeList As IXMLDOMNodeList
    Set householdNodeList = XMLRead.FindHouseholds(rtHousehold.NameOfHousehold)
    
    Dim memberType As String
    memberType = rtHousehold.members.Items(0).TypeOfMember
    If householdNodeList.Length = 0 And memberType = allowedMemberType Then
        'The household wasn't found and the member type is allowed. Add it with the member to the XML
        AddHouseholdToClientList rtHousehold, XMLClientList.SelectSingleNode("//Client_List"), errorDict(newHouseholdDictName), dictKey
    ElseIf householdNodeList.Length > 0 Then
        'At least one household was found, update the info for each one if it hasn't yet been updated
        Dim householdFound As Integer
        For householdFound = 0 To householdNodeList.Length - 1
            If Not householdsUpdated.Exists(dictKey) Then
                'The household hasn't been updated yet. Update it
                UpdateHouseholdNode rtHousehold, householdNodeList(householdFound), errorDict, dictKey
                
                If householdNodeList.Length > 1 And Not errorDict(dupeHouseholdDictName).Exists(dictKey) Then
                    'Make note that multiple households were found
                    errorDict(dupeHouseholdDictName).Add dictKey, DupeHouseholdDictMessage(rtHousehold.NameOfHousehold)
                End If
                
                'Add the household to the updated households dictionary
                householdsUpdated.Add dictKey, dictKey
            End If
        Next householdFound
    End If
    
    Set UpdateHouseholdNodes = XMLRead.FindHouseholds(householdName:=rtHousehold.NameOfHousehold, alsoReturnAUA:=True)
End Function

Private Sub UpdateMemberNodes(rw As Integer, householdNodeList As IXMLDOMNodeList, errorDict As Dictionary)
    'Exit the sub if there are no household nodes
    If householdNodeList.Length = 0 Then Exit Sub

    'Get the member from this row
    Dim rtMember As clsMember
    Set rtMember = GetMemberFromRow(rw)
    
    'Set the key for the error dictionaries
    Dim dictKey As String
    dictKey = rtMember.redtailID

    'Add/update the member for each household node
    Dim householdFound As Integer
    For householdFound = 0 To householdNodeList.Length - 1
        'Attempt to find the member in the XML
        Dim rtMemberList As IXMLDOMNodeList
        Set rtMemberList = XMLRead.FindMembersInHousehold(householdNode:=householdNodeList(householdFound), fName:=rtMember.fName, lName:=rtMember.lName, _
                                                          nickname:=rtMember.nickname, redtailID:=rtMember.redtailID)
                                                      
        'Add the member if it wasn't found, update it if it was
        If rtMemberList.Length = 0 And rtMember.TypeOfMember = allowedMemberType Then
            'The member wasn't found and the member's type is allowed. Add it to the household in the XML if the household name isn't blank
            If householdNodeList(householdFound).SelectSingleNode("./Name").Text <> vbNullString Then
                AddMemberToHousehold rtMember, householdNodeList(householdFound), errorDict(newMemberDictName), dictKey
            End If
        ElseIf rtMemberList.Length > 0 Then
            'At least one member was found. Update the info for each one
            Dim memberFound As Integer
            For memberFound = 0 To rtMemberList.Length - 1
                UpdateMemberNode rtMember, rtMemberList(memberFound), errorDict, dictKey
            Next memberFound
        End If
        
        If rtMemberList.Length > 1 And Not errorDict.Item(dupeMemberDictName).Exists(dictKey) Then
            'Make note that there are multiple member nodes with the same
            errorDict.Item(dupeMemberDictName).Add dictKey, DupeMemberDictMessage(rtMember.NameOfMember)
        End If
    Next householdFound
End Sub

Private Function GetHouseholdFromRow(rw As Integer) As clsHousehold
    'Create a household from the name
    Dim returnHousehold As clsHousehold
    Set returnHousehold = ClassConstructor.NewHousehold(householdName:=m_datasheet.GetData(householdNameHeader, rw))
    
    'Get the member from this row
    Dim returnMember As clsMember
    Set returnMember = GetMemberFromRow(rw)
    
    'Add the member to the household
    returnHousehold.AddMember returnMember
    
    'Return the household
    Set GetHouseholdFromRow = returnHousehold
End Function

Private Function GetMemberFromRow(rw As Integer) As clsMember
    'Get the member's attributes
    Dim memberFirstName As String, memberLastName As String, memberFullName As String, memberDateOfDeath As String
    memberFirstName = GetData(firstNameHeader, rw)
    memberLastName = GetData(lastNameHeader, rw)
    memberFullName = CombineNameParts(memberFirstName, GetData(middleNameHeader, rw), memberLastName, GetData(suffixHeader, rw), GetData(nicknameHeader, rw))
    memberDateOfDeath = Me.GetData(dateOfDeathHeader, rw)
    
    'Create the member
    Dim memberToReturn As clsMember
    Set memberToReturn = ClassConstructor.NewMember(memberFullName:=memberFullName, memberFirstName:=memberFirstName, memberLastName:=memberLastName, _
                                                    memberType:=Me.GetData(typeHeader, rw), memberStatus:=Me.GetData(statusHeader, rw), memberDateOfDeath:=memberDateOfDeath, _
                                                    contactID:=Me.GetData("ID", rw))
    
    'Return the member
    Set GetMemberFromRow = memberToReturn
End Function

Private Sub AddHouseholdToClientList(rtHousehold As clsHousehold, clientListNode As IXMLDOMNode, newHouseholdDict As Dictionary, dictKey As String)
    'Add the household to the client list node
    XMLWrite.AddHouseholdToNode rtHousehold, clientListNode, sheetName
    
    'Make note that the household with the member was added
    If Not newHouseholdDict.Exists(rtHousehold.NameOfHousehold) Then
        newHouseholdDict.Add dictKey, NewHouseholdDictMessage(rtHousehold.NameOfHousehold)
    End If
End Sub

Private Sub AddMemberToHousehold(rtMember As clsMember, householdNode As IXMLDOMNode, newMemberDict As Dictionary, dictKey As String)
    'Add the member to the household node
    XMLWrite.AddMemberToNode rtMember, householdNode, sheetName

    'Make note that it's being added, unless the household was already added in the import process
    Dim householdName As String
    householdName = householdNode.SelectSingleNode("./Name").Text
    If Not newMemberDict.Exists(dictKey & householdName) And XMLProcedures.GetAddDate(householdNode) < ProjectGlobals.ImportTime Then
        newMemberDict.Add dictKey & householdName, NewMemberDictMessage(rtMember.lName, rtMember.fName, householdName)
    End If
End Sub

Private Sub UpdateHouseholdNode(rtHousehold As clsHousehold, householdNode As IXMLDOMNode, errorDict As Dictionary, dictKey As String)
    With rtHousehold
        'Update the information in the household node
        Dim listHouseholdName As String
        XMLUpdate.UpdateHouseholdFromRTContactList householdNode, .NameOfHousehold, listHouseholdName
        
        'Check for differing information
        CheckForDifferingInfo "Name", .NameOfHousehold, listHouseholdName, errorDict(diffHouseholdNameDictName), dictKey, householdNode
    End With
End Sub

Private Sub UpdateMemberNode(rtMember As clsMember, memberNode As IXMLDOMNode, errorDict As Dictionary, dictKey As String)
    With rtMember
        'Update the information in the member node
        Dim listFirstName As String, listLastName As String, listStatus As String
        Dim listDeathDate As String
        XMLUpdate.UpdateMemberFromRTContactList memberNode, .redtailID, .fName, .lName, .Status, .dateOfDeath, listFirstName, _
                                                            listLastName, listStatus, listDeathDate
        
        'Check for differing information
        CheckForDifferingInfo "First Name", .fName, listFirstName, errorDict(diffFirstNameDictName), dictKey, memberNode, .NameOfMember
        CheckForDifferingInfo "Last Name", .lName, listLastName, errorDict(diffLastNameDictName), dictKey, memberNode, .NameOfMember
        CheckForDifferingInfo "Status", .Status, listStatus, errorDict(diffStatusDictName), dictKey, memberNode, .NameOfMember
        CheckForDifferingInfo "Date_of_Death", .dateOfDeath, listDeathDate, errorDict(diffDeathDateDictName), dictKey, memberNode, .NameOfMember
    End With
End Sub

Private Function CombineNameParts(firstName As String, middleName As String, lastName As String, suffix As String, nickname As String) As String
    'Put the row's values into an array
    Dim rowArray(4) As String
    rowArray(0) = firstName
    rowArray(1) = middleName
    rowArray(2) = lastName
    rowArray(3) = suffix
    rowArray(4) = nickname
    
    'rowArray contains: First name, middle name, last name, suffix, nickname
    'Build the string as "[last name] [suffix], [first name] [middle name] ([nickname])"
    Dim returnName As String
    returnName = rowArray(2)
    returnName = AddToString(addThisString:=rowArray(3), toThisString:=returnName, separatedBy:=" ")
    returnName = AddToString(addThisString:=rowArray(0), toThisString:=returnName, separatedBy:=", ")
    returnName = AddToString(addThisString:=rowArray(1), toThisString:=returnName, separatedBy:=" ")
    returnName = AddToString(addThisString:=rowArray(4), toThisString:=returnName, separatedBy:=" (")
    If Len(rowArray(4)) > 0 Then
        returnName = returnName & ")"
    End If
    
    'Return the name
    CombineNameParts = returnName
End Function

Private Function AddToString(addThisString As String, toThisString As String, separatedBy As String) As String
    If Len(addThisString) > 0 Then
        AddToString = toThisString & separatedBy & addThisString
    Else
        AddToString = toThisString
    End If
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
