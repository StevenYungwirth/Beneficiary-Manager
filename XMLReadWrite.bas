Attribute VB_Name = "XMLReadWrite"
Option Explicit
Public Const ClientListFolder As String = "Z:\FPIS - Operations\Beneficiary Project\"
'Public Const ClientListFile As String = "Z:\FPIS - Operations\Beneficiary Project\households.xml"
Public Const ClientListFile As String = "Z:\YungwirthSteve\Beneficiary Report\Assets\Sample Households.xml"
Public Const SampleClientListFile As String = "Z:\YungwirthSteve\Beneficiary Report\Assets\Sample Households.xml"

Public Function ReadHouseholdsFromFile() As Dictionary
    'Load the XML file
    Dim xmlFile As DOMDocument60
    Set xmlFile = LoadClientList
    
    'Set up the household dictionary to return
    Dim households As Dictionary
    Set households = New Dictionary
    
    'Convert each household node to a clsHousehold, and add it to the dictionary
    Dim xmlHouseholds As IXMLDOMNodeList
    Set xmlHouseholds = xmlFile.getElementsByTagName("Household")
    Dim householdNode As Integer
    For householdNode = 0 To xmlHouseholds.Length - 1
        'Add the household to the array
        Dim householdToAdd As clsHousehold
        Set householdToAdd = ReadHouseholdFromNode(xmlHouseholds(householdNode))
        households.Add householdToAdd.NameOfHousehold, householdToAdd
    Next householdNode
    
    'Return the households
    Set ReadHouseholdsFromFile = households
    
    'Unload the file
    Set xmlFile = Nothing
End Function

Public Function AddBenesToXML(NewBene As clsBeneficiary) As IXMLDOMElement
    'Load the XML file
    Dim xmlFile As DOMDocument60
    Set xmlFile = LoadClientList
    
    'Add the new bene to the file and return the created node
    Set AddBenesToXML = WriteBeneToXML(NewBene, xmlFile)
    
    'Save the file
    xmlFile.Save ClientListFile
    
    'Unload the file
    Set xmlFile = Nothing
End Function

Public Function UpdateRemoveAccounts(Update_1_Remove_2 As Integer, affectedAccounts As Dictionary) As Dictionary
    'Load the XML file
    Dim xmlFile As DOMDocument60
    Set xmlFile = LoadClientList
    
    Dim accountToUpdate As Variant
    For Each accountToUpdate In affectedAccounts.Items
        'Recast the account as a clsAccount
        Dim tempAccount As clsAccount
        Set tempAccount = accountToUpdate
        
        'Update the XML file
        If Not UpdateRemoveAccount(xmlFile, Update_1_Remove_2, tempAccount) Then
            'The account wasn't updated. Return it
            UpdateRemoveAccounts.Add tempAccount.Number, tempAccount
        End If
    Next accountToUpdate
End Function

Public Function UpdateRemoveBene(Update_1_Remove_2 As Integer, beneToUpdateOrRemove As clsBeneficiary, Optional updatedBene As clsBeneficiary) As Boolean
    'Load the XML file
    Dim xmlFile As DOMDocument60
    Set xmlFile = LoadClientList
    
    'Find the beneficiary in the file
    Dim accountNode As IXMLDOMNode
    Set accountNode = FindAccountByBene(beneToUpdateOrRemove, xmlFile)
    Dim beneNode As IXMLDOMNode
    Set beneNode = FindBeneNodeInAccountNode(beneToUpdateOrRemove, accountNode)
    
    'Update/Remove the beneficiary node
    If Update_1_Remove_2 = 2 And Not beneNode Is Nothing Then
        UpdateRemoveBene = RemoveBeneficiary(accountNode, beneNode)
    ElseIf Update_1_Remove_2 = 1 And Not updatedBene Is Nothing And Not beneNode Is Nothing _
    And (beneToUpdateOrRemove.NameOfBeneficiary <> updatedBene.NameOfBeneficiary _
    Or beneToUpdateOrRemove.Level <> updatedBene.Level _
    Or beneToUpdateOrRemove.Percent <> updatedBene.Percent) Then
        'Both benes exist and don't equal each other. Update the beneficiary node
        UpdateRemoveBene = UpdateBeneficiary(beneNode, updatedBene)
    End If
    
    'Save the file
    If UpdateRemoveBene Then
        xmlFile.Save ClientListFile
    End If
    
    'Unload the file
    Set xmlFile = Nothing
End Function

Public Function FindAccountByBene(bene As clsBeneficiary, xmlFile As DOMDocument60) As IXMLDOMNode
    'Find the account nodes with the given name
    Dim accountNodes As IXMLDOMNodeList
    Set accountNodes = xmlFile.SelectNodes("//Account[@Name=""" & bene.account.NameOfAccount & """]")
    
    'If there are multiple accounts with the same name, find the one with the matching account number
    If accountNodes.Length > 1 Then
        Dim accountNode As Variant
        For Each accountNode In accountNodes
            Dim node As IXMLDOMElement
            Set node = accountNode
            If node.getAttribute("Number") = CStr(bene.account.Number) Then
                Set FindAccountByBene = accountNode
            End If
        Next accountNode
    Else
        'There's only one account with that name
        Set FindAccountByBene = accountNodes.Item(0)
    End If
End Function

Public Function FindBeneNodeInAccountNode(bene As clsBeneficiary, accountNode As IXMLDOMNode) As IXMLDOMNode
    Dim beneNode As Integer
    For beneNode = 0 To accountNode.ChildNodes.Length - 1
        Dim beneAtNode As clsBeneficiary
        Set beneAtNode = ReadBeneficiaryFromNode(accountNode.ChildNodes(beneNode))
        If beneAtNode.NameOfBeneficiary = bene.NameOfBeneficiary And beneAtNode.Level = bene.Level And beneAtNode.Percent = bene.Percent _
        And beneAtNode.account.Number = bene.account.Number Then 'And beneAtNode.AcctName = bene.AcctName Then
            'The beneficiary in this node matches the one we're looking for. Return the node
            Set FindBeneNodeInAccountNode = accountNode.ChildNodes(beneNode)
            Exit Function
        End If
    Next beneNode
End Function

Public Function LoadClientList(Optional isSample As Boolean) As DOMDocument60
    'See if the households file is available
    Dim fso As FileSystemObject
    Set fso = New FileSystemObject
    
    'Set the filepath for the list
    Dim filePath As String
    If isSample Then
        filePath = SampleClientListFile
    Else
        filePath = ClientListFile
    End If
    
    If fso.FileExists(filePath) Then
        Set LoadClientList = New DOMDocument60
        LoadClientList.Load filePath
    Else
        MsgBox "Client List not found in default location."
        End
    End If
End Function

Public Function ReadHouseholdFromNode(node As IXMLDOMElement) As clsHousehold
    'Declare a temporary household to return
    Dim householdToReturn As clsHousehold
    Set householdToReturn = New clsHousehold

    'Set the name and address
    householdToReturn.NameOfHousehold = node.getAttribute("Name")
    
    'Get the household members
    Dim xmlMembers As IXMLDOMNodeList
    Set xmlMembers = node.SelectNodes("./Member")
    Dim member As Integer
    For member = 0 To xmlMembers.Length - 1
        'Get the member from the node
        Dim mmbr As clsMember
        Set mmbr = ReadMemberFromNode(xmlMembers(member))
        mmbr.ContainingHousehold = householdToReturn
        
        'Add the member to the household
        householdToReturn.AddMember mmbr
    Next member
    
    'Return the household
    Set ReadHouseholdFromNode = householdToReturn
End Function

Public Function ReadMemberFromNode(node As IXMLDOMElement) As clsMember
    'Declare a temporary member to return
    Dim memberToReturn As clsMember
    Set memberToReturn = New clsMember

    'Set the names and if they're deceased
    With memberToReturn
        .FName = node.getAttribute("First_Name")
        .LName = node.getAttribute("Last_Name")
        .Active = node.getAttribute("Active")
        .Deceased = node.getAttribute("Deceased")
    End With
            
    'Check if there's an override node
    Dim overrideNode As IXMLDOMElement
    Set overrideNode = node.SelectSingleNode("Override")
    If Not overrideNode Is Nothing Then
        'Get the overridden properties
        ReadMemberOverride memberToReturn, overrideNode
    End If

    'Get the member's accounts
    Dim xmlAccounts As IXMLDOMNodeList
    Set xmlAccounts = node.SelectNodes("Account")
    Dim account As Integer
    For account = 0 To xmlAccounts.Length - 1
        'Get the account from the node
        Dim acct As clsAccount
        Set acct = ReadAccountFromNode(xmlAccounts(account))
        acct.Owner = memberToReturn
        
        'Add the account to the member
        memberToReturn.AddAccount acct
    Next account
    
    'Return the member
    Set ReadMemberFromNode = memberToReturn
End Function

Private Sub ReadMemberOverride(member As clsMember, overrideNode As IXMLDOMElement)
    With member
        .NameOfMember = ReadOverride("Name", overrideNode, .NameOfMember)
        .FName = ReadOverride("First_Name", overrideNode, .FName)
        .LName = ReadOverride("Last_Name", overrideNode, .LName)
        .FullName = ReadOverride("Full_Name", overrideNode, .FullName)
        .Deceased = ReadOverride("Deceased", overrideNode, .Deceased)
        .Active = ReadOverride("Active", overrideNode, .Active)
    End With
End Sub

Public Function ReadAccountFromNode(node As IXMLDOMElement, Optional skipBeneficiaries As Boolean) As clsAccount
    'Declare a temporary account to return
    Dim accountToReturn As clsAccount
    Set accountToReturn = New clsAccount
    
    'Set the account ID, name, number, type, custodian, if it's active, and tag
    With accountToReturn
        .ID = node.getAttribute("Redtail_ID")
        .NameOfAccount = node.getAttribute("Name")
        .Number = node.getAttribute("Number")
        .TypeOfAccount = node.getAttribute("Type")
        .custodian = node.getAttribute("Custodian")
        .Active = node.getAttribute("Active")
        
        If Not IsNull(node.getAttribute("Balance")) Then
            .Balance = node.getAttribute("Balance")
        End If
        If Not node.SelectSingleNode("Tag") Is Nothing Then
            .Tag = node.SelectSingleNode("Tag").Text
        End If
    End With
            
    'Check if there's an override node
    Dim overrideNode As IXMLDOMElement
    Set overrideNode = node.SelectSingleNode("Override")
    If Not overrideNode Is Nothing Then
        'Get the overridden properties
        ReadAccountOverride accountToReturn, overrideNode
    End If
        
    If Not skipBeneficiaries Then
        'Get the account's beneficiaries
        Dim xmlBeneficiaries As IXMLDOMNodeList
        Set xmlBeneficiaries = node.SelectNodes("Beneficiary")
        Dim bene As Integer
        For bene = 0 To xmlBeneficiaries.Length - 1
            'Add the bene to the account
            accountToReturn.AddBene ReadBeneficiaryFromNode(xmlBeneficiaries(bene))
        Next bene
    End If
    
    'Return the account
    Set ReadAccountFromNode = accountToReturn
End Function

Private Sub ReadAccountOverride(account As clsAccount, overrideNode As IXMLDOMElement)
    With account
        .ID = ReadOverride("ID", overrideNode, .ID)
        .NameOfAccount = ReadOverride("Name", overrideNode, .NameOfAccount)
        .Number = ReadOverride("Number", overrideNode, .Number)
        .Balance = ReadOverride("Balance", overrideNode, .Balance)
        .TypeOfAccount = ReadOverride("Type", overrideNode, .TypeOfAccount)
        .custodian = ReadOverride("Custodian", overrideNode, .custodian)
        .Active = ReadOverride("Active", overrideNode, .Active)
        .CloseDate = ReadOverride("Close_Date", overrideNode, .CloseDate)
        .Tag = ReadOverride("Tag", overrideNode, .Tag)
    End With
End Sub

Public Function ReadBeneficiaryFromNode(node As IXMLDOMElement) As clsBeneficiary
    'Declare a temporary beneficiary to return
    Dim beneToReturn As clsBeneficiary
    Set beneToReturn = New clsBeneficiary
    
    'Check if needed text nodes are there
    If Not IsNull(node.getAttribute("Name")) And Not IsNull(node.getAttribute("Level")) And Not IsNull(node.getAttribute("Percent")) Then
        'All children are present
        With beneToReturn
            'Set the ID, name, level, and percent
            .ID = node.getAttribute("ID")
            .NameOfBeneficiary = node.getAttribute("Name")
            .Level = node.getAttribute("Level")
            .Percent = node.getAttribute("Percent")
            
            'Set the relationship, if it's there
            If Not IsNull(node.getAttribute("Relationship")) Then
                .Relation = node.getAttribute("Relationship")
            End If
            
            'Set the updated date, if it's there
            If Not IsNull(node.getAttribute("Last_Updated")) Then
                .UpdatedDate = node.getAttribute("Last_Updated")
            End If
            
            'Get parent account's name and number
            Dim parentAccount As IXMLDOMElement
            Set parentAccount = node.parentNode
            .account = ReadAccountFromNode(parentAccount, True)
        End With
    End If
    
    'Return the beneficiary
    Set ReadBeneficiaryFromNode = beneToReturn
End Function

Private Function ReadOverride(attributeName As String, node As IXMLDOMElement, initialValue As Variant) As Variant
    If Not IsNull(node.getAttribute(attributeName)) Then
        ReadOverride = node.getAttribute(attributeName)
    Else
        ReadOverride = initialValue
    End If
End Function

Private Function WriteBeneToXML(bene As clsBeneficiary, xmlFile As DOMDocument60) As IXMLDOMElement
    'Find the account in the client list file
    Dim selectedAccountNode As IXMLDOMNode
    Set selectedAccountNode = FindAccountByBene(bene, xmlFile)
    
    'Create a beneficiary node
    Dim beneNode As IXMLDOMElement
    Set beneNode = xmlFile.createNode(1, "Beneficiary", "")
    
    'Set properties for the account
    With beneNode
        .setAttribute "ID", bene.ID
        .setAttribute "Name", bene.NameOfBeneficiary
        .setAttribute "Relationship", bene.Relation
        .setAttribute "Level", bene.Level
        .setAttribute "Percent", bene.Percent
        .setAttribute "Added_On", bene.AddDate
        .setAttribute "Added_By", bene.AddedBy
        .setAttribute "Last_Updated", bene.AddDate
        .setAttribute "Updated_By", bene.AddedBy
    End With
        
    'Add the node to the account
    selectedAccountNode.appendChild beneNode
    
    'Increment the bene ID
    SetNextBeneID xmlFile
    
    'Return the added node
    Set WriteBeneToXML = beneNode
End Function

Private Sub SetNextBeneID(xmlFile As DOMDocument60)
    'Get the client list node
    Dim clientListNode As IXMLDOMElement
    Set clientListNode = xmlFile.SelectSingleNode("Client_List")
    
    'Increment and return the max beneficiary ID attribute
    clientListNode.setAttribute "Max_Beneficiary_ID", clientListNode.getAttribute("Max_Beneficiary_ID") + 1
End Sub

Private Function UpdateRemoveAccount(xmlFile As DOMDocument60, Update_1_Remove_2 As Integer, updatedAccount As clsAccount) As Boolean
    'Find the account in the file
    Dim accountNodeToUpdate As IXMLDOMElement
    Set accountNodeToUpdate = FindAccountByNumber(xmlFile, updatedAccount.Number)
    
    If Not accountNodeToUpdate Is Nothing Then
        If Update_1_Remove_2 = 1 Then
            'Update the account
            With accountNodeToUpdate
                .setAttribute "Type", updatedAccount.TypeOfAccount
                If updatedAccount.CloseDate > CDate(0) Then
                    'Account is closed
                    .setAttribute "Active", False
                End If
            End With
            UpdateRemoveAccount = True
        ElseIf Update_1_Remove_2 = 2 Then
            'Remove the account
            accountNodeToUpdate.parentNode.RemoveChild accountNodeToUpdate
            UpdateRemoveAccount = True
        End If
        
        If UpdateRemoveAccount Then
            'Save the XML file
            xmlFile.Save ClientListFile
        End If
    End If
    
    'Close the XML file
    Set xmlFile = Nothing
End Function

Private Function FindAccountByNumber(xmlFile As DOMDocument60, accountNumber As String) As IXMLDOMElement
    'Find the account nodes with the given number
    Dim accountNodes As IXMLDOMNodeList
    Set accountNodes = xmlFile.SelectNodes("Account[@Number=""" & accountNumber & """]")
    
    If accountNodes.Length = 1 Then
        'There's only one account with that number
        Set FindAccountByNumber = accountNodes.Item(0)
    Else
        'There are multiple accounts with that number
    End If
End Function

Private Function UpdateBeneficiary(beneToUpdate As IXMLDOMElement, NewBene As clsBeneficiary) As Boolean
    'Update the beneficiary
    With beneToUpdate
        .setAttribute "Name", NewBene.NameOfBeneficiary
        .setAttribute "Level", NewBene.Level
        .setAttribute "Percent", NewBene.Percent
        .setAttribute "Last_Updated", NewBene.AddDate
        .setAttribute "Updated_By", NewBene.AddedBy
    End With
    UpdateBeneficiary = True
End Function

Private Function RemoveBeneficiary(accountNode As IXMLDOMNode, beneToRemove As IXMLDOMNode) As Boolean
    'Remove the beneficiary
    accountNode.RemoveChild beneToRemove
    RemoveBeneficiary = True
End Function
