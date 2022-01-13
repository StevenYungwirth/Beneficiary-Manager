Attribute VB_Name = "ImportDataSheets"
Option Explicit
Private Const backupFolder As String = "Z:\YungwirthSteve\Beneficiary Report\Backup\Households\"
Public Sub ImportData()
    'Load the XML file
    Dim xmlFile As DOMDocument60
    Set xmlFile = XMLReadWrite.LoadClientList
    
    'Get the date the file was created
    Dim clientList As IXMLDOMElement
    Set clientList = xmlFile.SelectSingleNode("Client_List")
    Dim createDate As String
    createDate = clientList.getAttribute("Create_Date")
    
    'Create a backup copy of the XML file
    xmlFile.Save backupFolder & "Households " & Replace(createDate, "/", "-") & ".xml"
    
    'Reset the create date
    clientList.setAttribute "Create_Date", Format(Now(), "m/d/yyyy")
    
    'Declare each sheet to get from the form
    Dim tdaBeneList As clsTDABeneList
    Dim msAccounts As clsMSAccountList
    Dim rtAccounts As clsRTAccountList
    Dim rtContacts As clsRTContactList
    
    'Get the households from any importing sheets, using the current sheets for any that aren't being imported
    Dim newHouseholds As Dictionary
    Set newHouseholds = LoadNewHouseholds(tdaBeneList, msAccounts, rtAccounts, rtContacts)
    
    'Add/update/remove info in the XML
    CompareAndProcessHouseholds newHouseholds, xmlFile
    
    'Turn off screen updating
    UpdateScreen "Off"
    
    'Put the new sheets' data into the workbook
    tdaBeneList.FillWorksheet ThisWorkbook.Worksheets("TDA Bene List")
    msAccounts.FillWorksheet ThisWorkbook.Worksheets("MS Accounts")
    rtAccounts.FillWorksheet ThisWorkbook.Worksheets("RT Accounts")
    rtContacts.FillWorksheet ThisWorkbook.Worksheets("RT Contacts")
    
    'Turn screen updating back on
    UpdateScreen "On"
    
    'Save and close the XML
    xmlFile.Save XMLReadWrite.ClientListFile
    Set xmlFile = Nothing
    
    'Save this workbook
    ThisWorkbook.Save
    
    'Show confirmation
    MsgBox "Worksheets have been successfully imported."
End Sub

Private Function LoadNewHouseholds(tdaBeneList As clsTDABeneList, msAccounts As clsMSAccountList, rtAccounts As clsRTAccountList, rtContacts As clsRTContactList) As Dictionary
    'Open the form
    Load frmImport
    frmImport.Show
    
    'Get each sheet from the form
    With frmImport
        Set tdaBeneList = ClassConstructor.NewTDABeneList(.TDABeneFile)
        Set msAccounts = ClassConstructor.NewMSAccountList(.MSAccountsFile, True)
        Set rtAccounts = ClassConstructor.NewRTAccountList(.RTAccountsFile)
        Set rtContacts = ClassConstructor.NewRTContactList(.RTContactsFile, True)
    End With
    
    'Unload the form
    Unload frmImport
    
    'Get the list of households from the sheets
    Set LoadNewHouseholds = LoadHouseholds.GetHouseholds(tdaBeneList, msAccounts, rtAccounts, rtContacts)
End Function

Private Sub CompareAndProcessHouseholds(households As Dictionary, xmlFile As DOMDocument60)
    'Add "Delete" attributes to every XML tag
    AddDeleteAttribute xmlFile
    
    'Declare a list of elements that will be added to the XML
    Dim addingHouseholds As Collection, addingMembers As Collection, addingAccounts As Collection, addingBenes As Collection
    Set addingHouseholds = New Collection
    Set addingMembers = New Collection
    Set addingAccounts = New Collection
    Set addingBenes = New Collection
    
    Dim addingElements As Dictionary
    Set addingElements = New Dictionary
    With addingElements
        .Add "Households", addingHouseholds
        .Add "Members", addingMembers
        .Add "Accounts", addingAccounts
        .Add "Beneficiaries", addingBenes
    End With
    
    'Add/update info in the XML, removing "Delete" attribute from every element found
    ReconcileHouseholds households, xmlFile, addingElements
    
    'Show nodes to be added and deleted
    If (addingElements("Households").count > 0 Or addingElements("Members").count > 0 Or addingElements("Accounts").count > 0 Or addingElements("Beneficiaries").count > 0) _
        Or GetDeletedNodes(xmlFile).Length > 0 Then
        'Store the XML changes in a CSV file
        CreateChangeLog.CreateSheet addingElements, GetDeletedNodes(xmlFile)
        
        'Show the form with the XML changes
        Load frmShowChanges
        With frmShowChanges
            .ShowAddedNodes addingElements
            .ShowDeletedNodes GetDeletedNodes(xmlFile)
            .Show
        End With
        
        'Unload the form
        Unload frmShowChanges
    End If
    
    'Remove the elements that still have the "Delete" attribute
    DeleteLeftoverElements xmlFile
    
    '(maybe do this) If accounts are being removed and added (The member changed; probably a Morningstar issue), bring their benes with
    'Could do the same with members, but a problem comes up with common names (i.e. Tom Nelson leaves, but we gain a different Tom Nelson)
End Sub

Private Sub AddDeleteAttribute(xmlFile As DOMDocument60)
    'Get the list of every node except the root
    Dim allNodes As IXMLDOMNodeList
    Set allNodes = xmlFile.SelectNodes("Client_List//*")
    
    'For each node, set an attribute "Delete" equal to true
    Dim node As Variant
    For Each node In allNodes
        Dim SelectedNode As IXMLDOMElement
        Set SelectedNode = node
        SelectedNode.setAttribute "Delete", "True"
    Next node
End Sub

Private Sub ReconcileHouseholds(households As Dictionary, xmlFile As DOMDocument60, addingElements As Dictionary)
    'Declare a list of households to add to the XML
    Dim householdsToAdd As Collection
    Set householdsToAdd = New Collection

    'Look for each household in the XML
    Dim household As Variant
    For Each household In households.Items
        'Get the current household
        Dim currentHousehold As clsHousehold
        Set currentHousehold = household
    
        'Get the client list node
        Dim clientListNode As IXMLDOMNode
        Set clientListNode = xmlFile.SelectSingleNode("Client_List")
        
        Dim foundHousehold As IXMLDOMElement
        Set foundHousehold = FindHouseholdNode(currentHousehold.NameOfHousehold, clientListNode)
        If Not foundHousehold Is Nothing Then
            'The household was found, remove the "Delete" attribute from it
            foundHousehold.removeAttribute "Delete"
            
            'Look for each member within the household
            ReconcileMembers currentHousehold.Members, foundHousehold, xmlFile, addingElements
        Else
            'The household wasn't found. Add it to the XML
            XMLCreateList.AddHouseholdToNode currentHousehold, clientListNode, xmlFile
            householdsToAdd.Add currentHousehold
        End If
    Next household
    
    'Return the list of added households
    If householdsToAdd.count > 0 Then
        Dim householdToAdd As Variant
        For Each householdToAdd In householdsToAdd
            addingElements("Households").Add householdToAdd
        Next householdToAdd
    End If
End Sub

Private Function FindHouseholdNode(householdName As String, clientListNode As IXMLDOMNode) As IXMLDOMElement
    Dim households As IXMLDOMNodeList
    Set households = clientListNode.SelectNodes("Household[@Name=""" & householdName & """]")
    
    If households.Length = 1 Then
        'There's only one household, return it
        Set FindHouseholdNode = households.Item(0)
    ElseIf households.Length > 0 Then
        'Multiple households were found
        Stop
    Else
        'The household wasn't found, return nothing
        Set FindHouseholdNode = Nothing
    End If
End Function

Private Sub ReconcileMembers(Members As Dictionary, householdNode As IXMLDOMNode, xmlFile As DOMDocument60, addingElements As Dictionary)
    'Declare a list of members to add to the XML
    Dim membersToAdd As Collection
    Set membersToAdd = New Collection
    
    'Look for each member in the household
    Dim member As Variant
    For Each member In Members.Items
        Dim currentMember As clsMember
        Set currentMember = member
        
        Dim foundMember As IXMLDOMElement
        Set foundMember = FindMemberNode(currentMember.FName, currentMember.LName, householdNode)
        If Not foundMember Is Nothing Then
            'The member was found, remove the "Delete" attribute from it
            foundMember.removeAttribute "Delete"
            
            'Set the member's non-key properties to be the incoming properties
            foundMember.setAttribute "Active", CStr(currentMember.Active)
            foundMember.setAttribute "Deceased", CStr(currentMember.Deceased)
            
            'Look for each account within the member
            ReconcileAccounts currentMember.accounts, foundMember, xmlFile, addingElements
        Else
            'The member wasn't found. Add it to the XML
            XMLCreateList.AddMemberToNode currentMember, householdNode, xmlFile
            membersToAdd.Add currentMember
        End If
    Next member
    
    'Return the list of added members
    If membersToAdd.count > 0 Then
        Dim memberToAdd As Variant
        For Each memberToAdd In membersToAdd
            addingElements("Members").Add memberToAdd
        Next memberToAdd
    End If
End Sub

Private Function FindMemberNode(firstName As String, lastName As String, householdNode As IXMLDOMNode) As IXMLDOMElement
    Dim searchString As String
    searchString = "Member[@First_Name=""" & firstName & """ and @Last_Name=""" & lastName & """]"
    Dim Members As IXMLDOMNodeList
    Set Members = householdNode.SelectNodes(searchString)
    
    If Members.Length = 1 Then
        'There's only one member, return it
        Set FindMemberNode = Members.Item(0)
    ElseIf Members.Length > 0 Then
        'Multiple members were found
        Stop
    Else
        'The member wasn't found, return nothing
        Set FindMemberNode = Nothing
    End If
End Function

Private Sub ReconcileAccounts(accounts As Dictionary, memberNode As IXMLDOMNode, xmlFile As DOMDocument60, addingElements As Dictionary)
    'Declare a list of accounts to add to the XML
    Dim accountsToAdd As Collection
    Set accountsToAdd = New Collection
    
    'Look for each account in the member
    Dim account As Variant
    For Each account In accounts.Items
        Dim currentAccount As clsAccount
        Set currentAccount = account
        
        Dim foundAccount As IXMLDOMElement
        Set foundAccount = FindAccountNode(currentAccount.NameOfAccount, currentAccount.Number, memberNode)
        If Not foundAccount Is Nothing Then
            'The account was found, remove the "Delete" attribute from it and its tag
            foundAccount.removeAttribute "Delete"
            Dim tagNode As IXMLDOMElement
            Set tagNode = foundAccount.SelectSingleNode("Tag")
            tagNode.removeAttribute "Delete"
            
            'Set the account's non-key properties to be the incoming properties
            With foundAccount
                .setAttribute "Redtail_ID", currentAccount.ID
                .setAttribute "Type", currentAccount.TypeOfAccount
                .setAttribute "Custodian", currentAccount.custodian
                .setAttribute "Owner", currentAccount.Owner.NameOfMember
                .setAttribute "Active", CStr(currentAccount.Active)
                .setAttribute "Balance", currentAccount.Balance
            End With
            
            'Look for each beneficiary within the account
            ReconcileBenes currentAccount.Benes, foundAccount, xmlFile, addingElements
        Else
            'The account wasn't found. Add it to the XML
            XMLCreateList.AddAccountToNode currentAccount, memberNode, xmlFile
            accountsToAdd.Add currentAccount
        End If
    Next account
    
    'Return the list of added accounts
    If accountsToAdd.count > 0 Then
        Dim accountToAdd As Variant
        For Each accountToAdd In accountsToAdd
            addingElements("Accounts").Add accountToAdd
        Next accountToAdd
    End If
End Sub

'Private Sub TestReconcile(eleDict As Dictionary, parentNode As IXMLDOMNode, xmlFile As DOMDocument60, addingElements As Dictionary)
'    'Declare a list of components to add to the XML
'    Dim elesToAdd As Collection
'    Set elesToAdd = New Collection
'
'    'Look for each component in the parent
'    Dim ele As Variant
'    For Each ele In eleDict.Items
'        Dim currentEle As Object
'        Set currentEle = ele
'
'        Dim foundEle As IXMLDOMElement
'        Set foundEle = FindChildNode() 'Parameters differ per component
'        If Not foundEle Is Nothing Then
'            'The component was found, remove the "Delete" attribute from it
'            foundEle.removeAttribute "Delete"
'
'            'Set the component's non-key properties to be the incoming properties
'            foundEle.setAttribute "", "" 'Differs per component
'
'            'Loop for each child within the component
'            TestReconcile 'Bene has no children
'        Else
'            'The component wasn't found. Add it to the XML
'            XMLCreateList 'Call a different function depending on the component
'            elesToAdd.Add currentEle
'        End If
'    Next ele
'
'    'Return the list of added components
'    If elesToAdd.count > 0 Then
'        Dim eleToAdd As Variant
'        For Each eleToAdd In elesToAdd
'            addingElements("").Add eleToAdd 'Differs per component
'        Next eleToAdd
'End Sub

Private Function FindAccountNode(accountName As String, accountNumber As String, memberNode As IXMLDOMNode) As IXMLDOMElement
    Dim searchString As String
    searchString = "Account[@Name=""" & accountName & """ and @Number=""" & accountNumber & """]"
    Dim accounts As IXMLDOMNodeList
    Set accounts = memberNode.SelectNodes(searchString)
    
    If accounts.Length = 1 Then
        'There's only one account, return it
        Set FindAccountNode = accounts.Item(0)
    ElseIf accounts.Length > 0 Then
        'Multiple accounts were found
        Stop
    Else
        'The account wasn't found, return nothing
        Set FindAccountNode = Nothing
    End If
End Function

Private Sub ReconcileBenes(Benes As Collection, accountNode As IXMLDOMNode, xmlFile As DOMDocument60, addingElements As Dictionary)
    'Declare a list of beneficiaries to add to the XML
    Dim benesToAdd As Collection
    Set benesToAdd = New Collection
    
    'Look for each bene in the account
    Dim bene As Variant
    For Each bene In Benes
        Dim currentBene As clsBeneficiary
        Set currentBene = bene
        
        Dim foundBene As IXMLDOMElement
        Set foundBene = FindBeneNode(currentBene, accountNode)
        If Not foundBene Is Nothing Then
            'The beneficiary was found, remove the "Delete" attribute from it
            foundBene.removeAttribute "Delete"
            
            'Set the beneficiary's non-key properties to be the incoming properties
            With foundBene
                If .getAttribute("Level") <> currentBene.Level Or .getAttribute("Percent") <> currentBene.Percent Then
                    'At least one of the level or percent values are different. Update them and add the update information
                    .setAttribute "Level", currentBene.Level
                    .setAttribute "Percent", currentBene.Percent
                    .setAttribute "Last_Updated", currentBene.AddDate
                    .setAttribute "Updated_By", "Import"
                End If
            End With
        Else
            'The beneficiary wasn't found. Add it to the XML
            XMLCreateList.AddBeneficiaryToNode currentBene, accountNode, xmlFile
            benesToAdd.Add currentBene
        End If
    Next bene
    
    'Return the list of added beneficiaries
    If benesToAdd.count > 0 Then
        Dim beneToAdd As Variant
        For Each beneToAdd In benesToAdd
            addingElements("Beneficiaries").Add beneToAdd
        Next beneToAdd
    End If
End Sub

Private Function FindBeneNode(bene As clsBeneficiary, accountNode As IXMLDOMNode) As IXMLDOMElement
    Dim searchString As String
    searchString = "Beneficiary[@Name=""" & Replace(bene.NameOfBeneficiary, """", "&quot;") & """ and @Level=""" & bene.Level & """ and @Percent=""" & bene.Percent & """]"
    Dim Benes As IXMLDOMNodeList
    Set Benes = accountNode.SelectNodes(searchString)
    
    If Benes.Length = 1 Then
        'There's only one beneficiary, return it
        Set FindBeneNode = Benes.Item(0)
    ElseIf Benes.Length > 0 Then
        'Multiple beneficiaries were found
        'Gene Musgrove IRA #939926870 and Roth IRA #945979350 has John D Musgrove (Labeled as son in TD) as a 100% primary bene and John D Musgrove (Brother) as a 100% Contingent
        'Carol Ready TOD 2 #939571779 has two identical beneficiaries - "St Vincent De Paul, P, 15"
        'In her Rollover IRA #962995180, the beneficiaries show one should be "St Vincent De Paul - Gladstone, MI" and the other "St Vincent De Paul - Escanaba, MI"
'        Stop
    Else
        'The beneficiary wasn't found, return nothing
        Set FindBeneNode = Nothing
    End If
End Function

Private Sub DeleteLeftoverElements(xmlFile As DOMDocument60)
    'Get the list of every node with the "Delete" attribute
    Dim nodesToDelete As IXMLDOMNodeList
    Set nodesToDelete = GetDeletedNodes(xmlFile)
    
    'For each node, remove it from its parent
    Dim node As Variant
    For Each node In nodesToDelete
        Dim SelectedNode As IXMLDOMNode
        Set SelectedNode = node
        SelectedNode.parentNode.RemoveChild SelectedNode
    Next node
End Sub

Private Sub UpdateScreen(OnOrOff As String)
    With Application
        If OnOrOff = "Off" Then
            .ScreenUpdating = False
            .EnableEvents = False
            .DisplayStatusBar = False
            .Calculation = xlCalculationManual
        ElseIf OnOrOff = "On" Then
            .ScreenUpdating = True
            .EnableEvents = True
            .DisplayStatusBar = True
            .Calculation = xlCalculationAutomatic
            Dim Reset As Long
            Reset = ActiveSheet.UsedRange.Rows.count
        End If
    End With
End Sub

Private Function GetDeletedNodes(xmlFile As DOMDocument60) As IXMLDOMNodeList
    Set GetDeletedNodes = xmlFile.SelectNodes("//*[@Delete='True']")
End Function
