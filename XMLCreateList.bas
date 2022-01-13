Attribute VB_Name = "XMLCreateList"
Option Explicit
Private Const ClientListFile As String = "Z:\FPIS - Operations\Beneficiary Project\households.xml"
Private Const AssociatedFileLocation As String = "Z:\YungwirthSteve\Beneficiary Report\Assets\associated accounts.txt"

Public Sub CreateHouseholdsXMLFile(households As Dictionary)
    'Create the households file
    Dim xmlFile As DOMDocument60
    Set xmlFile = New DOMDocument60
    With xmlFile
        .setProperty "SelectionLanguage", "XPath"
        .setProperty "SelectionNamespaces", "xmlns:xsl='http://www.w3.org/1999/XSL/Transform'"
        
        'Append an xml processing instruction
        .appendChild xmlFile.createProcessingInstruction("xml", "version='1.0' encoding='UTF-8'")
    
        'Add the client list node with all households
        Dim clientListNode As IXMLDOMNode
        Set clientListNode = CreateClientListNode(households, xmlFile)
        
        'Add the root node to the file
        .appendChild clientListNode
        
        'Save the XML file
        .Save ClientListFile
    End With
End Sub

Private Function CreateClientListNode(households As Dictionary, xmlFile As DOMDocument60) As IXMLDOMNode
    'Create the ClientList node
    Dim clientListNode As IXMLDOMNode
    Set clientListNode = xmlFile.createNode(1, "Client_List", "")
    
    'Add each household to the node
    AddHouseholdsToNode households.Items, clientListNode, xmlFile
    
    'Return the client list node
    Set CreateClientListNode = clientListNode
End Function

Private Sub AddHouseholdsToNode(clientListHouseholds As Variant, targetNode As IXMLDOMNode, xmlFile As DOMDocument60)
    Dim clientListHousehold As Variant
    For Each clientListHousehold In clientListHouseholds
        Dim household As clsHousehold
        Set household = clientListHousehold
        AddHouseholdToNode household, targetNode, xmlFile
    Next clientListHousehold
End Sub

Public Sub AddHouseholdToNode(clientListHousehold As clsHousehold, targetNode As IXMLDOMNode, xmlFile As DOMDocument60)
    'Create a houeshold node
    Dim householdNode As IXMLDOMElement
    Set householdNode = xmlFile.createNode(1, "Household", "")
    
    'Set properties for the household
    householdNode.setAttribute "Name", clientListHousehold.NameOfHousehold
    
    'Add each member to the household
    AddMembersToNode clientListHousehold.Members.Items, householdNode, xmlFile
    
    'Add household to node
    targetNode.appendChild householdNode
End Sub

Public Sub AddMembersToNode(householdMembers As Variant, targetNode As IXMLDOMNode, xmlFile As DOMDocument60)
    Dim householdMember As Variant
    For Each householdMember In householdMembers
        Dim member As clsMember
        Set member = householdMember
        AddMemberToNode member, targetNode, xmlFile
    Next householdMember
End Sub

Public Sub AddMemberToNode(householdMember As clsMember, targetNode As IXMLDOMNode, xmlFile As DOMDocument60)
    'Create a member node
    Dim memberNode As IXMLDOMElement
    Set memberNode = xmlFile.createNode(1, "Member", "")
    
    'Set properties for the member
    With memberNode
        .setAttribute "First_Name", householdMember.FName
        .setAttribute "Last_Name", householdMember.LName
        .setAttribute "Active", CStr(householdMember.Active)
        .setAttribute "Deceased", CStr(householdMember.Deceased)
    End With
    
    'Add each account to the member
    AddAccountsToNode householdMember.accounts.Items, memberNode, xmlFile
    
    'Add member to node
    targetNode.appendChild memberNode
End Sub

Public Sub AddAccountsToNode(memberAccounts As Variant, targetNode As IXMLDOMNode, xmlFile As DOMDocument60)
    Dim memberAccount As Variant
    For Each memberAccount In memberAccounts
        AddAccountToNode memberAccount, targetNode, xmlFile
    Next memberAccount
End Sub

Public Sub AddAccountToNode(memberAccount As Variant, targetNode As IXMLDOMNode, xmlFile As DOMDocument60)
    'Load the list of Associated Bank account names
    Dim associatedAccounts() As String
    associatedAccounts = LoadAssociatedAccounts

    'Create an account node
    Dim accountNode As IXMLDOMElement
    Set accountNode = xmlFile.createNode(1, "Account", "")
    
    'Set properties for the account
    With accountNode
        .setAttribute "Redtail_ID", memberAccount.ID
        .setAttribute "Name", memberAccount.NameOfAccount
        .setAttribute "Number", memberAccount.Number
        .setAttribute "Type", memberAccount.TypeOfAccount
        .setAttribute "Custodian", memberAccount.custodian
        .setAttribute "Owner", memberAccount.Owner.NameOfMember
        .setAttribute "Active", CStr(memberAccount.Active)
        
        'Create a Tag node
        Dim tagNode As IXMLDOMNode
        Set tagNode = xmlFile.createNode(1, "Tag", "")
        
        'Have tag's text be "WEC" or "Associated" if it's easily identifiable from the account name or in the list of known Associated Bank account names
        tagNode.Text = AutoTag(memberAccount.NameOfAccount, associatedAccounts)
        
        'Add the Tag node
        .appendChild tagNode
    End With

    'Add each beneficiary to the account
    AddBeneficiariesToNode memberAccount.Benes, accountNode, xmlFile
    
    'Add the node to the member
    targetNode.appendChild accountNode
End Sub

Private Function LoadAssociatedAccounts() As String()
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

Private Sub AddBeneficiariesToNode(accountBenes As Variant, targetNode As IXMLDOMNode, xmlFile As DOMDocument60)
    Dim accountBene As Variant
    For Each accountBene In accountBenes
        Dim bene As clsBeneficiary
        Set bene = accountBene
        AddBeneficiaryToNode bene, targetNode, xmlFile
    Next accountBene
End Sub

Public Sub AddBeneficiaryToNode(accountBene As clsBeneficiary, targetNode As IXMLDOMNode, xmlFile As DOMDocument60)
    'Create a beneficiary node
    Dim beneNode As IXMLDOMElement
    Set beneNode = xmlFile.createNode(1, "Beneficiary", "")
    
    'Set properties for the account
    With beneNode
        .setAttribute "ID", GetNextBeneID(xmlFile)
        .setAttribute "Name", accountBene.NameOfBeneficiary
        .setAttribute "Relationship", accountBene.Relation
        .setAttribute "Level", accountBene.Level
        .setAttribute "Percent", accountBene.Percent
        .setAttribute "Added_On", accountBene.AddDate
        .setAttribute "Added_By", "Import"
        .setAttribute "Last_Updated", accountBene.AddDate
        .setAttribute "Updated_By", "Import"
    End With

    'Add the node to the account
    targetNode.appendChild beneNode
End Sub

Private Function GetNextBeneID(xmlFile As DOMDocument60) As Integer
    'Get the client list node
    Dim clientListNode As IXMLDOMElement
    Set clientListNode = xmlFile.SelectSingleNode("Client_List")
    
    'Increment and return the max beneficiary ID attribute
    GetNextBeneID = clientListNode.getAttribute("Max_Beneficiary_ID") + 1
    clientListNode.setAttribute "Max_Beneficiary_ID", clientListNode.getAttribute("Max_Beneficiary_ID") + 1
End Function
