Attribute VB_Name = "XMLRead"
Option Explicit

Private Property Get XMLClientList() As DOMDocument60
    Set XMLClientList = ProjectGlobals.ClientListFile
End Property

Public Function HouseholdsFromFile() As Dictionary
    'Set up the household dictionary to return
    Dim households As Dictionary
    Set households = New Dictionary
    
    'Convert each household node to a clsHousehold, and add it to the dictionary
    Dim xmlHouseholds As IXMLDOMNodeList
    Set xmlHouseholds = XMLClientList.SelectNodes("//Household")
    Dim householdNode As Integer
    For householdNode = 0 To xmlHouseholds.Length - 1
        'Add the household to the array
        Dim householdToAdd As clsHousehold
        Set householdToAdd = HouseholdFromNode(xmlHouseholds(householdNode), True)
        
        households.Add householdToAdd.NameOfHousehold, householdToAdd
    Next householdNode
    
    'Return the households
    Set HouseholdsFromFile = households
End Function

Public Function FindHouseholds(householdName As String, Optional morningstarID As String, Optional alsoReturnAUA As Boolean) As IXMLDOMNodeList
    'Attempt to find the household
    Dim msHouseholdList As IXMLDOMNodeList
    Set msHouseholdList = XMLClientList.SelectNodes("not //*")
    Dim attempt As Integer: attempt = 0
    Do
        attempt = attempt + 1
        Select Case attempt
        Case 1
            'Attempt to find the household by its ID
            If morningstarID <> vbNullString Then
                Set msHouseholdList = XMLClientList.SelectNodes("//Household[./Morningstar_ID[text()='" & morningstarID & "']]")
            End If
        Case 2
            'Attempt to find the household by its name
            Dim searchString As String
            searchString = "//Household[./Name[text()=" & SearchWrapper(householdName) & "]]"
            If alsoReturnAUA Then
                searchString = searchString & " | //Household[./Name[text()=" & SearchWrapper(householdName & " AUA") & "]]"
            End If
            Set msHouseholdList = XMLClientList.SelectNodes(searchString)
        End Select
    Loop While msHouseholdList.Length = 0 And attempt <= 1
    
    'Return the nodes found
    Set FindHouseholds = msHouseholdList
End Function

Public Function FindMembersInHousehold(householdNode As IXMLDOMNode, fName As String, lName As String, Optional fullName As String, _
                                       Optional nickname As String, Optional redtailID As Long) As IXMLDOMNodeList
    'Attempt to find the member by its name
    Dim memberNodeList As IXMLDOMNodeList
    Set memberNodeList = householdNode.SelectNodes("not //*")
    Dim attempt As Integer: attempt = 0
    Do
        attempt = attempt + 1
        Select Case attempt
        Case 1
            'Attempt to find the member by Redtail ID
            If redtailID <> 0 Then
                Set memberNodeList = householdNode.SelectNodes("Member[./Redtail_ID[text()='" & redtailID & "']]")
            End If
        Case 2
            'Attempt to find the member by its first and last name
            Set memberNodeList = householdNode.SelectNodes("Member[./Last_Name[text()=" & SearchWrapper(lName) & "] and (./First_Name[text()=" _
                                 & SearchWrapper(fName) & "] or ./First_Name[text()=" & SearchWrapper(nickname) & "])]")
        Case 3
            'Attempt to find the member by its full name
            If fullName <> vbNullString Then
                Set memberNodeList = householdNode.SelectNodes("./Member[./Full_Name[text()=" & SearchWrapper(fullName) & "]]")
            End If
        End Select
    Loop While memberNodeList.Length = 0 And attempt <= 1
    
    'Return the nodes found
    Set FindMembersInHousehold = memberNodeList
End Function

Public Function FindAccounts(accountNumber As String, Optional accountName As String, Optional morningstarID As String, Optional householdNode As IXMLDOMElement) As IXMLDOMNodeList
    'Get which node to search from
    Dim nodeToSearchFrom As IXMLDOMNode
    Dim searchPrefix As String
    If householdNode Is Nothing Then
        Set nodeToSearchFrom = XMLClientList.SelectSingleNode("Client_List")
        searchPrefix = "//"
    Else
        Set nodeToSearchFrom = householdNode
        searchPrefix = vbNullString
    End If
    
    'Declare a list to return
    Dim accountNodeList As IXMLDOMNodeList
    Set accountNodeList = nodeToSearchFrom.SelectNodes("not //*")
    
    'Attempt to find the account
    Dim attempt As Integer: attempt = 0
    Do
        attempt = attempt + 1
        Select Case attempt
        Case 1
            'Attempt to find the account by ID
            If morningstarID <> vbNullString Then
                Set accountNodeList = nodeToSearchFrom.SelectNodes(searchPrefix & "Account[./Morningstar_ID[text()='" & morningstarID & "']]")
            End If
        Case 2
            'Attempt to find the account by name and number
            If accountNumber = vbNullString And accountName <> vbNullString Then
                Set accountNodeList = nodeToSearchFrom.SelectNodes(searchPrefix & "Account[./Name[text()=" & SearchWrapper(accountName) & "]]")
            ElseIf accountNumber <> vbNullString And accountName <> vbNullString Then
                Set accountNodeList = nodeToSearchFrom.SelectNodes(searchPrefix & "Account[./Name[text()=" & SearchWrapper(accountName) & "] and ./Number[text()='" & accountNumber & "']]")
            End If
        Case 3
            'Attempt to find the account by number
            If accountNumber <> vbNullString Then
                '(account number not being blank prevents heldaway accounts with no account number from being pulled in)
                Set accountNodeList = nodeToSearchFrom.SelectNodes(searchPrefix & "Account[./Number[text()='" & accountNumber & "']]")
            End If
        End Select
    Loop While accountNodeList.Length = 0 And attempt <= 2
    
    'Return the nodes found
    Set FindAccounts = accountNodeList
End Function

Public Function FindAccountInMember(memberNode As IXMLDOMNode, accountNumber As String, Optional accountName As String, Optional accountMSID As String) As IXMLDOMNode
    'Attempt to find the account
    Dim accountsFound As IXMLDOMNodeList
    Set accountsFound = memberNode.SelectNodes("not //*")
    Dim attempt As Integer: attempt = 0
    Do
        attempt = attempt + 1
        Select Case attempt
        Case 1
            'Attempt to find the account by Morningstar ID
            If accountMSID <> vbNullString Then
                Set accountsFound = memberNode.SelectNodes("./Account[./Morningstar_ID[text()='" & accountMSID & "']]")
            End If
        Case 2
            'Attempt to find the account by number (or name, if the account number is blank)
            If accountNumber <> vbNullString Then
                Set accountsFound = memberNode.SelectNodes("./Account[./Number[text() = '" & accountNumber & "']]")
            Else
                Set accountsFound = memberNode.SelectNodes("./Account[./Name[text() = " & SearchWrapper(accountName) & "]]")
            End If
        End Select
    Loop While accountsFound.Length = 0 And attempt <= 1
    
    'Return the first node found
    If accountsFound.Length > 0 Then Set FindAccountInMember = accountsFound(0)
End Function

'TODO need to change/delete "if there are multiple accounts" section. Changed this to initially look for account number
Public Function FindAccountByBene(bene As clsBeneficiary) As IXMLDOMNode
    'Find the account nodes with the given name
    Dim accountNodes As IXMLDOMNodeList
    Set accountNodes = XMLClientList.SelectNodes("//Account[./Number[text()=" & SearchWrapper(bene.account.Number) & "]]")
    
    'If there are multiple accounts with the same name, find the one with the matching account number
    If accountNodes.Length > 1 Then
        Dim accountNode As Variant
        For Each accountNode In accountNodes
            Dim Node As IXMLDOMElement
            Set Node = accountNode
            If Node.SelectSingleNode("Number") = bene.account.Number Then
                Set FindAccountByBene = accountNode
            End If
        Next accountNode
    Else
        'There's only one account with that name
        Set FindAccountByBene = accountNodes.Item(0)
    End If
End Function

Public Function FindBenesInAccount(accountNode As IXMLDOMNode, beneName As String, beneLevel As String, benePercent As Double, Optional beneID As Integer) As IXMLDOMNodeList
    'Attempt to find the beneficiary
    Dim beneNodeList As IXMLDOMNodeList
    Set beneNodeList = accountNode.SelectNodes("not //*")
    Dim attempt As Integer: attempt = 0
    Do
        attempt = attempt + 1
        Select Case attempt
        Case 1
            'Attempt to find the beneficiary by ID
            If beneID <> 0 Then
                Set beneNodeList = accountNode.SelectNodes("Beneficiary[./ID[text()='" & beneID & "']]")
            End If
        Case 2
            'Attempt to find the beneficiary by name, level, and percent
            Set beneNodeList = accountNode.SelectNodes("Beneficiary[./Name[text()=" & SearchWrapper(beneName) & "] and ./Level[text()='" _
                               & beneLevel & "'] and ./Percent[text()='" & benePercent & "']]")
        Case 3
            'Attempt to find the beneficiary by only the name
            Set beneNodeList = accountNode.SelectNodes("Beneficiary[./Name[text()=" & SearchWrapper(beneName) & "]]")
        End Select
    Loop While beneNodeList.Length = 0 And attempt <= 2
    
    'Return the nodes found
    Set FindBenesInAccount = beneNodeList
End Function

Public Function HouseholdFromNode(sourceNode As IXMLDOMElement, readChildren As Boolean) As clsHousehold
    'Declare a temporary household to return
    Dim householdToReturn As clsHousehold: Set householdToReturn = New clsHousehold

    'Set the name and address
    householdToReturn.NameOfHousehold = sourceNode.SelectSingleNode("Name").Text
    householdToReturn.Active = sourceNode.SelectSingleNode("Active").Text
    
    If readChildren Then
        'Get the household members
        Dim xmlMembers As IXMLDOMNodeList
        Set xmlMembers = sourceNode.SelectNodes("Member")
        Dim member As Integer
        For member = 0 To xmlMembers.Length - 1
            'Get the member from the node
            Dim mmbr As clsMember
            Set mmbr = MemberFromNode(xmlMembers(member), True)
            mmbr.ContainingHousehold = householdToReturn
            
            'Add the member to the household
            householdToReturn.AddMember mmbr
        Next member
    End If
    
    'Return the household
    Set HouseholdFromNode = householdToReturn
End Function

Public Function MemberFromNode(sourceNode As IXMLDOMElement, readChildren As Boolean) As clsMember
    'Declare a temporary member to return
    Dim memberToReturn As clsMember
    Set memberToReturn = New clsMember

    'Set the names and if they're deceased
    With memberToReturn
        .fName = sourceNode.SelectSingleNode("First_Name").Text
        .lName = sourceNode.SelectSingleNode("Last_Name").Text
        .Status = sourceNode.SelectSingleNode("Status").Text
        .Active = sourceNode.SelectSingleNode("Active").Text
        .dateOfDeath = sourceNode.SelectSingleNode("Date_of_Death").Text
        
        'Get the parent household
        Dim parentHousehold As IXMLDOMElement
        Set parentHousehold = sourceNode.parentNode
        .ContainingHousehold = HouseholdFromNode(sourceNode.parentNode, False)
    End With
            
    'Check if there's an override node
    Dim overrideNode As IXMLDOMElement
    Set overrideNode = sourceNode.SelectSingleNode("Override")
    If Not overrideNode Is Nothing Then
        'Get the overridden properties
        MemberOverride memberToReturn, overrideNode
    End If

    If readChildren Then
        'Get the member's accounts
        Dim xmlAccounts As IXMLDOMNodeList
        Set xmlAccounts = sourceNode.SelectNodes("Account")
        Dim account As Integer
        For account = 0 To xmlAccounts.Length - 1
            'Get the account from the node
            Dim acct As clsAccount
            Set acct = AccountFromNode(xmlAccounts(account), True)
            acct.owner = memberToReturn
            
            'Add the account to the member
            memberToReturn.AddAccount acct
        Next account
    End If
    
    'Return the member
    Set MemberFromNode = memberToReturn
End Function

Public Function MemberNameFromNode(memberNode As IXMLDOMNode) As String
    'Return the full name if present, or last name[comma] first name
    If memberNode.SelectSingleNode("Full_Name") Is Nothing Then
        MemberNameFromNode = memberNode.SelectSingleNode("Last_Name").Text & ", " & memberNode.SelectSingleNode("First_Name").Text
    Else
        MemberNameFromNode = memberNode.SelectSingleNode("Full_Name").Text
    End If
End Function

Public Function AccountFromNode(sourceNode As IXMLDOMElement, readChildren As Boolean) As clsAccount
    'Declare a temporary account to return
    Dim accountToReturn As clsAccount
    Set accountToReturn = New clsAccount
    
    'Set the account ID, name, number, type, custodian, if it's active, and tag
    With accountToReturn
        .morningstarID = sourceNode.SelectSingleNode("Morningstar_ID").Text
        .NameOfAccount = sourceNode.SelectSingleNode("Name").Text
        .Number = sourceNode.SelectSingleNode("Number").Text
        .TypeOfAccount = sourceNode.SelectSingleNode("Type").Text
        .custodian = sourceNode.SelectSingleNode("Custodian").Text
        .Active = sourceNode.SelectSingleNode("Active").Text
        
        
        'Get the parent member
        Dim parentMember As IXMLDOMElement
        Set parentMember = sourceNode.parentNode
        .owner = MemberFromNode(parentMember, False)
        
        If Not sourceNode.SelectSingleNode("Balance") Is Nothing Then
            .Balance = sourceNode.SelectSingleNode("Balance").Text
        End If
        If Not sourceNode.SelectSingleNode("Tag") Is Nothing Then
            .Tag = sourceNode.SelectSingleNode("Tag").Text
        End If
    End With
            
    'Check if there's an override node
    Dim overrideNode As IXMLDOMElement
    Set overrideNode = sourceNode.SelectSingleNode("Override")
    If Not overrideNode Is Nothing Then
        'Get the overridden properties
        accountOverride accountToReturn, overrideNode
    End If
        
    If readChildren Then
        'Get the account's beneficiaries
        Dim xmlBeneficiaries As IXMLDOMNodeList
        Set xmlBeneficiaries = sourceNode.SelectNodes("Beneficiary")
        Dim bene As Integer
        For bene = 0 To xmlBeneficiaries.Length - 1
            'Add the bene to the account
            accountToReturn.AddBene BeneficiaryFromNode(xmlBeneficiaries(bene)), False
        Next bene
    End If
    
    'Return the account
    Set AccountFromNode = accountToReturn
End Function

Public Function BeneficiaryFromNode(sourceNode As IXMLDOMElement) As clsBeneficiary
    'Declare a temporary beneficiary to return
    Dim beneToReturn As clsBeneficiary
    Set beneToReturn = New clsBeneficiary
    
    'Check if needed text nodes are there
    If Not sourceNode.SelectSingleNode("Name") Is Nothing _
    And Not sourceNode.SelectSingleNode("Level") Is Nothing _
    And Not sourceNode.SelectSingleNode("Percent") Is Nothing Then
        'All children are present
        With beneToReturn
            'Set the ID, name, level, and percent
            If Not sourceNode.SelectSingleNode("ID") Is Nothing Then
                If sourceNode.SelectSingleNode("ID").Text <> vbNullString Then
                    .id = sourceNode.SelectSingleNode("ID").Text
                End If
            End If
            .NameOfBeneficiary = sourceNode.SelectSingleNode("Name").Text
            .Level = sourceNode.SelectSingleNode("Level").Text
            .Percent = sourceNode.SelectSingleNode("Percent").Text
            
            'Set the relationship, if it's there
            If Not sourceNode.SelectSingleNode("Relationship") Is Nothing Then
                .Relation = sourceNode.SelectSingleNode("Relationship").Text
            End If
            
'            If .id = 994 Then Stop
            
            'Set the updated date, if it's there
            If Not sourceNode.SelectSingleNode("Last_Updated") Is Nothing Then
                .UpdatedDate = sourceNode.SelectSingleNode("Last_Updated").Text
            End If
            
            'Get the parent account
            Dim parentAccount As IXMLDOMElement
            Set parentAccount = sourceNode.parentNode
            .account = AccountFromNode(parentAccount, False)
        End With
    End If
    
    'Return the beneficiary
    Set BeneficiaryFromNode = beneToReturn
End Function

Private Function searchString(bene As clsBeneficiary) As String
    'Search for the beneficiary using the account and bene information
    Dim accountPortion As String, benePortion As String
    If bene.account.custodian = ProjectGlobals.DefaultCustodian Then
        accountPortion = "Account[./Number[text()='" & bene.account.Number & "']]"
    Else
        accountPortion = "account[./Name[text()=" & SearchWrapper(bene.account.NameOfAccount) & "] and ./Number[text()='" _
                         & bene.account.Number & "']]"
    End If
    benePortion = "Beneficiary[./Name[text()=" & SearchWrapper(bene.NameOfBeneficiary) & "] and ./Level[text()='" & bene.Level _
                  & "'] and ./Percent[text()='" & bene.Percent & "']]"
    searchString = "//" & accountPortion & "/" & benePortion
End Function

Private Sub MemberOverride(member As clsMember, overrideNode As IXMLDOMElement)
    With member
        .NameOfMember = ReadOverride("Name", overrideNode, .NameOfMember)
        .fName = ReadOverride("First_Name", overrideNode, .fName)
        .lName = ReadOverride("Last_Name", overrideNode, .lName)
        .dateOfDeath = ReadOverride("Date_of_Death", overrideNode, .dateOfDeath)
        .Status = ReadOverride("Status", overrideNode, .Status)
    End With
End Sub

Private Sub accountOverride(account As clsAccount, overrideNode As IXMLDOMElement)
    With account
        .redtailID = ReadOverride("ID", overrideNode, .redtailID)
        .NameOfAccount = ReadOverride("Name", overrideNode, .NameOfAccount)
        .Number = ReadOverride("Number", overrideNode, .Number)
        .Balance = ReadOverride("Balance", overrideNode, .Balance)
        .TypeOfAccount = ReadOverride("Type", overrideNode, .TypeOfAccount)
        .custodian = ReadOverride("Custodian", overrideNode, .custodian)
        .Active = ReadOverride("Active", overrideNode, .Active)
        .closeDate = ReadOverride("Close_Date", overrideNode, .closeDate)
        .Tag = ReadOverride("Tag", overrideNode, .Tag)
    End With
End Sub

Private Function ReadOverride(attributeName As String, sourceNode As IXMLDOMElement, initialValue As Variant) As Variant
    If Not sourceNode.SelectSingleNode(attributeName) Is Nothing Then
        ReadOverride = sourceNode.SelectSingleNode(attributeName).Text
    Else
        ReadOverride = initialValue
    End If
End Function

Private Function SearchWrapper(value As String) As String
    'Create a search string with appropriate single or double quotes
    SearchWrapper = XMLProcedures.XPathExpression(value)
End Function
