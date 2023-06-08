Attribute VB_Name = "XMLWrite"
Option Explicit

Private Property Get MSExportName() As String
    MSExportName = ProjectGlobals.m_msExportName
End Property

Private Property Get MSAccountName() As String
    MSAccountName = ProjectGlobals.m_msAccountName
End Property

Private Property Get RTAccountName() As String
    RTAccountName = ProjectGlobals.m_rtAccountName
End Property

Private Property Get RTContactName() As String
    RTContactName = ProjectGlobals.m_rtContactName
End Property

Private Property Get BeneListName() As String
    BeneListName = ProjectGlobals.m_beneListName
End Property

Private Property Get ManualBeneListName() As String
    ManualBeneListName = ProjectGlobals.m_manualBeneListName
End Property

Private Property Get XMLClientList() As DOMDocument60
    Set XMLClientList = ProjectGlobals.ClientListFile
End Property

Public Function AddHouseholdToNode(clientListHousehold As clsHousehold, targetNode As IXMLDOMNode, Optional sheetName As String) As IXMLDOMNode
    'Create a houeshold node
    Dim householdNode As IXMLDOMElement
    Set householdNode = CreateAndAppendNode(targetNode, "Household", ProjectGlobals.HouseholdNodeProperties(clientListHousehold), sheetName)
    SetAddedAttributes householdNode, sheetName
    
    'Add each member to the household
    AddMembersToNode clientListHousehold.members.Items, householdNode, sheetName
    
    'Return the added node
    Set AddHouseholdToNode = householdNode
End Function

Public Function AddMemberToNode(householdMember As clsMember, targetNode As IXMLDOMNode, Optional sheetName As String) As IXMLDOMNode
    'Create the member node and append it to the target node
    Dim memberNode As IXMLDOMElement
    Set memberNode = CreateAndAppendNode(targetNode, "Member", ProjectGlobals.MemberNodeProperties(householdMember), sheetName)
    SetAddedAttributes memberNode, sheetName
    
    'Add each account to the member
    AddAccountsToNode householdMember.Accounts.Items, memberNode, sheetName
    
    'Return the added node
    Set AddMemberToNode = memberNode
End Function

Public Function AddAccountToNode(memberAccount As clsAccount, targetNode As IXMLDOMNode, Optional sheetName As String) As IXMLDOMNode
    'Create the account node and append it to the target node
    Dim accountNode As IXMLDOMElement
    Set accountNode = CreateAndAppendNode(targetNode, "Account", ProjectGlobals.AccountNodeProperties(memberAccount), sheetName)
    SetAddedAttributes accountNode, sheetName

    'Add each beneficiary to the account
    AddBeneficiariesToNode memberAccount.Benes.Items, accountNode, sheetName
    
    'Return the added node
    Set AddAccountToNode = accountNode
End Function

Public Function AddBeneficiaryToNodes() As IXMLDOMElement
    'Put the account nodes the beneficiary is being added to into an array
    Dim accountNodesToAddBeneTo(1) As IXMLDOMNode
    If selectedAccountNode Is Nothing Then
        'Find the account nodes and convert the list to an array
        Dim accountsFound As IXMLDOMNodeList
        Set accountsFound = XMLRead.FindAccountByBene(bene)
        If accountsFound.Length = 0 Then
            'The account couldn't be found. Return nothing
            Set AddBeneficiaryToNode = Nothing
            Exit Function
        ElseIf accountsFound.Length = 1 Then
            Set accountNodesToAddBeneTo(0) = accountsFound(0)
        Else
            ReDim accountNodesToAddBeneTo(0 To accountsFound.Length - 1) As IXMLDOMNode
            Dim accountFound As Integer
            For accountFound = 0 To accountsFound.Length - 1
                Set accountNodesToAddBeneTo(accountFound) = accountsFound(accountFound)
            Next accountFound
        End If
    Else
        Set accountNodesToAddBeneTo(0) = selectedAccountNode
    End If
End Function

Public Function AddBeneficiaryToNode(bene As clsBeneficiary, Optional selectedAccountNode As IXMLDOMNode, Optional sheetName As String) As IXMLDOMElement
    'Find the account in the client list file
    If selectedAccountNode Is Nothing Then
        Set selectedAccountNode = XMLRead.FindAccountByBene(bene)
        If selectedAccountNode Is Nothing Then
            'The account couldn't be found. Return nothing
            Set AddBeneficiaryToNode = Nothing
            Exit Function
        End If
    End If
    
    'Create a beneficiary node
    Dim benenode As IXMLDOMElement
    Set benenode = CreateAndAppendNode(selectedAccountNode, "Beneficiary", ProjectGlobals.BeneficiaryNodeProperties(bene), sheetName)
    SetAddedAttributes benenode, sheetName
    
    'Return the added node
    Set AddBeneficiaryToNode = benenode
End Function

Public Function UpdateRemoveBene(Update_1_Remove_2 As Integer, beneToUpdateOrRemove As clsBeneficiary, Optional updatedBene As clsBeneficiary) As Boolean
    'Find the account
    Dim accountNodeList As IXMLDOMNodeList
    Set accountNodeList = XMLRead.FindAccounts(accountNumber:=beneToUpdateOrRemove.account.Number, accountName:=beneToUpdateOrRemove.account.NameOfAccount)

    'Exit the function if the account node wasn't found
    If accountNodeList.Length = 0 Then Exit Function
    
    'Find the beneficiary in the file
    Dim accountNode As IXMLDOMNode: Set accountNode = accountNodeList(0)
    Dim beneNodeList As IXMLDOMNodeList
    With beneToUpdateOrRemove
        Set beneNodeList = XMLRead.FindBenesInAccount(accountNode, .NameOfBeneficiary, .Level, .Percent)
    End With
    
    If beneNodeList.Length > 0 Then
        'Update/Remove the beneficiary node
        Dim benenode As IXMLDOMNode
        Set benenode = beneNodeList(0)
        If Update_1_Remove_2 = 2 Then
            'Remove the beneficiary node from the account node
            UpdateRemoveBene = RemoveBeneficiary(accountNode, benenode)
        ElseIf Update_1_Remove_2 = 1 And Not updatedBene Is Nothing _
        And (beneToUpdateOrRemove.NameOfBeneficiary <> updatedBene.NameOfBeneficiary _
        Or beneToUpdateOrRemove.Level <> updatedBene.Level _
        Or beneToUpdateOrRemove.Percent <> updatedBene.Percent) Then
            'Both benes exist and don't equal each other. Update the beneficiary node
            UpdateRemoveBene = XMLUpdate.UpdateBeneficiaryFromForm(benenode, updatedBene)
        End If
        
        'Save the file
        If UpdateRemoveBene Then
            XMLProcedures.FormatAndSaveXML
        End If
    End If
End Function

Private Function CreateClientListNode(households As Dictionary) As IXMLDOMNode
    'Create the ClientList node
    Dim clientListNode As IXMLDOMNode
    Set clientListNode = XMLClientList.createNode(1, "Client_List", "")
    
    'Add each household to the node
    AddHouseholdsToNode households.Items, clientListNode
    
    'Return the client list node
    Set CreateClientListNode = clientListNode
End Function

Private Sub AddHouseholdsToNode(clientListHouseholds As Variant, targetNode As IXMLDOMNode, Optional sheetName As String)
    Dim clientListHousehold As Variant
    For Each clientListHousehold In clientListHouseholds
        Dim household As clsHousehold
        Set household = clientListHousehold
        AddHouseholdToNode household, targetNode, sheetName
    Next clientListHousehold
End Sub

Private Sub AddMembersToNode(householdMembers As Variant, targetNode As IXMLDOMNode, sheetName As String)
    Dim householdMember As Variant
    For Each householdMember In householdMembers
        Dim member As clsMember
        Set member = householdMember
        AddMemberToNode member, targetNode, sheetName
    Next householdMember
End Sub

Private Sub AddAccountsToNode(memberAccounts As Variant, targetNode As IXMLDOMNode, sheetName As String)
    Dim memberAccountItem As Variant
    For Each memberAccountItem In memberAccounts
        Dim memberAccount As clsAccount
        Set memberAccount = memberAccountItem
        AddAccountToNode memberAccount, targetNode, sheetName
    Next memberAccountItem
End Sub

Private Sub AddBeneficiariesToNode(accountBenes As Variant, targetNode As IXMLDOMNode, sheetName As String)
    Dim accountBene As Variant
    For Each accountBene In accountBenes
        Dim bene As clsBeneficiary
        Set bene = accountBene
        AddBeneficiaryToNode bene, targetNode, sheetName
    Next accountBene
End Sub

Private Function RemoveBeneficiary(accountNode As IXMLDOMNode, beneToRemove As IXMLDOMNode) As Boolean
    'Remove the beneficiary
    accountNode.RemoveChild beneToRemove
    RemoveBeneficiary = True
End Function

Private Function CreateAndAppendNode(parentNode As IXMLDOMNode, childName As String, nodeProperties As Dictionary, sheetName As String) As IXMLDOMElement
    'Create the child node to append
    Dim childNode As IXMLDOMElement
    Set childNode = parentNode.OwnerDocument.createNode(1, childName, "")
    
    'Add properties to the child node
    AddPropertiesToNode childNode, nodeProperties, sheetName
    
    'Add child to parent
    If Not parentNode Is Nothing Then
        parentNode.appendChild childNode
    End If
    
    'Mark the child as being in the sheet
    XMLProcedures.FlagNodeInList childNode, sheetName, True
    
    'Return the child node
    Set CreateAndAppendNode = childNode
End Function

Private Sub AddPropertiesToNode(parentNode As IXMLDOMNode, nodeProperties As Dictionary, sheetName As String)
    Dim prop As Integer
    For prop = 0 To nodeProperties.count - 1
        'TODO see if leaving this commented increases performance
'        Debug.Print nodeProperties.Keys(prop) & " " & nodeProperties.Items(prop)
        
        'Create and append the property's node
        Dim childNode As IXMLDOMNode
        Set childNode = XMLClientList.createNode(1, nodeProperties.Keys(prop), "")
        childNode.Text = nodeProperties.Items(prop)
        SetAddedAttributes childNode, sheetName
        parentNode.appendChild childNode
    Next prop
    
    'Set Added On date
    Dim parentEle As IXMLDOMElement
    Set parentEle = parentNode
End Sub

Private Sub SetAddedAttributes(attributeNode As IXMLDOMNode, adderName As String)
    If adderName <> vbNullString Then
        Dim attributeElement As IXMLDOMElement: Set attributeElement = attributeNode
        attributeElement.setAttribute "Added_By", adderName
        attributeElement.setAttribute "Added_On", Format(ProjectGlobals.ImportTime, "m/d/yyyy h:mm;@")
    End If
End Sub
