VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTDABeneList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_datasheet As clsDataSheet
Private Const accountNumberHeader As String = "Account#"
Private Const accountTypeHeader As String = "AcctDescription"
Private Const openDateHeader As String = "OpenDate"
Private Const closeDateHeader As String = "DateClosed"
Private Const name1Header As String = "NameLine1"
Private Const name2Header As String = "NameLine2"
Private Const name3Header As String = "NameLine3"
Private Const beneNameHeader As String = "Name"
Private Const beneRelationshipHeader As String = "Relationship"
Private Const beneLevelHeader As String = "BeneLevel"
Private Const benePercentHeader As String = "Percentage"
Private Const newAccountDictName As String = "New Account"
Private Const newBeneDictName As String = "New Beneficiary"
Private Const dupeAccountDictName As String = "Duplicate Account"
Private Const dupeBeneDictName As String = "Duplicate Beneficiary"
Private Const diffCloseDateDictName As String = "Different Close Date"
Private Const diffAccountTypeDictName As String = "Different Account Type"
Private Const diffLevelDictName As String = "Different Beneficiary Level"
Private Const diffPercentDictName As String = "Different Beneficiary Percent"

Private Property Get ErrorDictKeys() As String()
    Dim keyArray(0 To 7) As String
    keyArray(0) = newAccountDictName
    keyArray(1) = newBeneDictName
    keyArray(2) = dupeAccountDictName
    keyArray(3) = dupeBeneDictName
    keyArray(4) = diffCloseDateDictName
    keyArray(5) = diffAccountTypeDictName
    keyArray(6) = diffLevelDictName
    keyArray(7) = diffPercentDictName
    ErrorDictKeys = keyArray
End Property

Private Property Get NewAccountDictMessage(accountNumber As String) As String
    NewAccountDictMessage = "Account Added - " & accountNumber & " without a household"
End Property

Private Property Get NewBeneDictMessage(beneName As String, accountNumber As String, accountName As String) As String
    NewBeneDictMessage = "Beneficiary Added - " & beneName & " | Account - " & accountNumber
    
    If accountName <> vbNullString Then
        NewBeneDictMessage = NewBeneDictMessage & " - " & accountName
    End If
End Property

Private Property Get DupeAccountDictMessage(accountNumber As String) As String
    DupeAccountDictMessage = "Account Duplicate - " & accountNumber & " - Multiple instances found in XML"
End Property

Private Property Get DupeBeneDictMessage(beneName As String, accountNumber As String) As String
    DupeBeneDictMessage = "Beneficiary Duplicate - " & beneName & " - Multiple instances found in account " & accountNumber
End Property

Public Property Get DataSize() As Integer
    DataSize = m_datasheet.DataSize
End Property

Private Property Get sheetName() As String
    sheetName = ProjectGlobals.m_beneListName
End Property

Private Property Get XMLClientList() As DOMDocument60
    Set XMLClientList = ProjectGlobals.ClientListFile
End Property

Public Sub ClassBuilder(path As String)
    'Set the required headers
    Dim requiredHeaders() As String
    requiredHeaders = Split(accountNumberHeader & "," & accountTypeHeader & "," & openDateHeader & "," & closeDateHeader _
                            & "," & name1Header & "," & name2Header & "," & name3Header _
                            & "," & beneNameHeader & "," & beneRelationshipHeader & "," & beneLevelHeader & "," & benePercentHeader, ",")
    
    'Set the datasheet in order to get the data
    Set m_datasheet = ClassConstructor.NewDataSheet(filePath:=path, wkstName:=sheetName, reqHeaders:=requiredHeaders)
End Sub

Public Function GetData(headerName As String, datapoint As Integer) As String
    GetData = m_datasheet.GetData(headerName, datapoint)
End Function

Public Sub FillWorksheet(sht As Worksheet)
    m_datasheet.FillWorksheet sht
End Sub

Private Function GetBeneFromRow(rw As Integer) As clsBeneficiary
    'Get the beneficiary on this row
    Dim tempBene As clsBeneficiary
    Set tempBene = ClassConstructor.NewBene(beneName:=GetData(beneNameHeader, rw), beneLevel:=GetData(beneLevelHeader, rw), _
                                            benePercent:=Val(GetData(benePercentHeader, rw)), beneRelation:=GetData(beneRelationshipHeader, rw))
    tempBene.account.Number = GetData(accountNumberHeader, rw)
                                            
    'Return the beneficiary, if there's one on this row
    If Len(tempBene.NameOfBeneficiary) > 0 Then
        Set GetBeneFromRow = tempBene
    End If
End Function

Public Sub AddBenesFromTD(accountDict As Dictionary)
    'Set an array for missing accounts to log to error file
    Dim missingAccounts() As String
    ReDim missingAccounts(0) As String
    
    'For each row, find the account and add the beneficiaries
    Dim rw As Integer
    For rw = 1 To UBound(m_datasheet.Data, 1) - 1
        Dim accountNumber As String
        accountNumber = GetData(accountNumberHeader, rw)
        If accountDict.Exists(accountNumber) Then
            'The account is in the dictionary. Add the bene if it's there and mark it as active
            If Len(GetData(beneNameHeader, rw)) > 0 Then
                Dim beneToAdd As clsBeneficiary
                Set beneToAdd = ClassConstructor.NewBene(beneName:=GetData(beneNameHeader, rw), beneLevel:=GetData(beneLevelHeader, rw), _
                                                         benePercent:=CDbl(GetData(benePercentHeader, rw)), beneRelation:=GetData(beneRelationshipHeader, rw))
                
                accountDict.Item(accountNumber).AddBene beneToAdd, False
            End If
            
            'Check if the account is closed
            If Len(GetData(closeDateHeader, rw)) > 0 Then
                'The account is closed. Mark it as inactive
                accountDict.Item(accountNumber).Active = False
            Else
                'The account is open. Mark it as active
                accountDict.Item(accountNumber).Active = True
            End If
        Else
            'The account isn't in the dictionary. Add it to the missing accounts
            Dim missingAccount As String
            missingAccount = accountNumber & " - " & GetData(name1Header, rw) & " " & GetData(name2Header, rw) & " " & GetData(name3Header, rw)
            If Not IsInArray(missingAccount, missingAccounts) Then
                If Len(missingAccounts(0)) = 0 Then
                    missingAccounts(0) = missingAccount
                Else
                    ReDim Preserve missingAccounts(0 To UBound(missingAccounts) + 1)
                    missingAccounts(UBound(missingAccounts)) = missingAccount
                End If
            End If
        End If
    Next rw
    
    'Log all errors to the error file
    ErrorHandling.LogErrorToFile ErrorHandling.AccountsNotInMSError(missingAccounts)
End Sub

Public Function ImportToXML(Optional testingLimit As Integer) As Dictionary
    'Contains master Redtail Contact ID
    'Contains household name, member name, member active status, date of death
    
    'Set up dictionaries for the possible errors to return
    Dim errorDict As Dictionary
    Set errorDict = SetUpDictionaries
    
    'Flag every account and beneficiary as not in the list
    FlagAsNotInList "Account"
    FlagAsNotInList "Beneficiary"
    
    'Iterate through each row of the TD bene list
    Dim accountsUpdated As Dictionary: Set accountsUpdated = New Dictionary
    Dim rw As Integer
    For rw = 1 To m_datasheet.DataSize - 1
        If rw Mod 100 = 0 Then
            DoEvents
            Debug.Print sheetName & " " & rw & " / " & m_datasheet.DataSize - 1
            Application.StatusBar = sheetName & " " & rw & " / " & m_datasheet.DataSize - 1
        End If
         
        'Update or add the account
        Dim accountNodeList As IXMLDOMNodeList
        Set accountNodeList = UpdateAccountNodes(rw, errorDict, accountsUpdated)
        
        'Update or add the beneficiary
        UpdateBeneNodes rw, accountNodeList, errorDict
        
        'Exit the loop if the testing limit has been reached
        If rw = testingLimit Then: rw = m_datasheet.DataSize - 1
    Next rw
    
    'Reset the status bar
    Application.StatusBar = False
    
    Set ImportToXML = errorDict
End Function

Public Function GetAccountFromRow(rw As Integer) As clsAccount
    'Initialize the account
    Dim tdAccount As clsAccount
    Set tdAccount = ClassConstructor.NewAccount(accountName:=vbNullString, accountNumber:=m_datasheet.GetData(accountNumberHeader, rw), _
                                                accountType:=m_datasheet.GetData(accountTypeHeader, rw), accountCustodian:=ProjectGlobals.DefaultCustodian, _
                                                accountTag:=vbNullString)
                                                
    'Add the open/close dates
    Dim openDate As String, closeDate As String
    tdAccount.openDate = m_datasheet.GetData(openDateHeader, rw)
    tdAccount.closeDate = m_datasheet.GetData(closeDateHeader, rw)
    
    'Convert TD account types to Morningstar
    tdAccount.TypeOfAccount = ConvertAccountType(tdAccount.TypeOfAccount)
    
    'Add the beneficiary on this row
    tdAccount.AddBene GetBeneFromRow(rw), False
    
    'Return the account
    Set GetAccountFromRow = tdAccount
End Function

Private Function UpdateAccountNodes(rw As Integer, errorDict As Dictionary, accountsUpdated As Dictionary) As IXMLDOMNodeList
    'Get the account from this row
    Dim tdAccount As clsAccount
    Set tdAccount = GetAccountFromRow(rw)
    
    'Set the key for the error dictionaries
    Dim dictKey As String
    dictKey = tdAccount.Number
    
    'Attempt to find the account(s) in the XML by its number
    Dim tdAccountList As IXMLDOMNodeList
    Set tdAccountList = XMLRead.FindAccounts(tdAccount.Number)
    
    'Add the account if it wasn't found, update it if it was
    If tdAccountList.Length = 0 Then
        'The account wasn't found. Add it to the XML
        AddAccountToClientList tdAccount, XMLClientList.SelectSingleNode("Client_List"), errorDict(newAccountDictName), dictKey
    Else
        'At least one account node was found, update the info for each one if it hasn't yet been updated
        If Not accountsUpdated.Exists(dictKey) Then
            'The account hasn't been updated yet. Update it
            Dim nodeFound As Integer
            For nodeFound = 0 To tdAccountList.Length - 1
                'Update the account
                UpdateAccountNode tdAccount, tdAccountList(nodeFound), errorDict, dictKey
            Next nodeFound
        
            'Check for duplicate accounts
            If tdAccountList.Length > 1 And Not errorDict.Item(dupeAccountDictName).Exists(tdAccount.Number) Then
                'Make note that there are multiple nodes with the same account number
                errorDict.Item(dupeAccountDictName).Add tdAccount.Number, tdAccount.Number & " - Multiple instances found in XML"
            End If
            
            'Add the household to the updated households dictionary
            accountsUpdated.Add dictKey, dictKey
        End If
    End If
    
    'Return the account nodes
    Set UpdateAccountNodes = tdAccountList
End Function

Private Sub UpdateBeneNodes(rw As Integer, accountNodeList As IXMLDOMNodeList, errorDict As Dictionary)
    'Exit the sub if no account was found
    If accountNodeList.Length = 0 Then Exit Sub
    
    'Get the beneficiary from this row
    Dim beneOnRow As clsBeneficiary
    Set beneOnRow = GetBeneFromRow(rw)
    
    'Exit the sub if no beneficiary is on this row
    If beneOnRow Is Nothing Then Exit Sub
    
    'Set the key for the error dictionaries
    Dim dictKey As String
    With beneOnRow
        dictKey = .NameOfBeneficiary & .Level & .Percent & .account.Number
    End With
    
    'Add/update the beneficiary in each account found
    Dim accountFound As Integer
    For accountFound = 0 To accountNodeList.Length - 1
        'Attempt to find the beneficiary(s) in the account by its components
        Dim tdBeneList As IXMLDOMNodeList
        With beneOnRow
            Set tdBeneList = XMLRead.FindBenesInAccount(accountNodeList(accountFound), .NameOfBeneficiary, .Level, .Percent)
        End With
        
        If tdBeneList.Length = 0 And XMLProcedures.GetAddDate(accountNodeList(accountFound)) < ProjectGlobals.ImportTime Then
            'The beneficiary wasn't found. Add it to the account unless the account was already added in the process
            AddBeneToAccount beneOnRow, accountNodeList(accountFound), errorDict(newBeneDictName), dictKey
        Else
            'At least one beneficiary was found. Update the info for each one
            Dim nodeFound As Integer
            For nodeFound = 0 To tdBeneList.Length - 1
                UpdateBeneNode beneOnRow, tdBeneList(nodeFound), errorDict, dictKey
            Next nodeFound
    
            'Check for duplicate accounts
            If tdBeneList.Length > 1 And Not errorDict.Item(dupeBeneDictName).Exists(dictKey) Then
                'Make note that there are multiple account nodes with the same name/number
                errorDict.Item(dupeBeneDictName).Add dictKey, DupeBeneDictMessage(beneOnRow.NameOfBeneficiary, beneOnRow.account.Number)
            End If
        End If
    Next accountFound
End Sub

Private Sub AddAccountToClientList(tdAccount As clsAccount, clientListNode As IXMLDOMNode, newAccountDict As Dictionary, dictKey As String)
    'Add the account as its own node to the client list
    XMLWrite.AddAccountToNode tdAccount, clientListNode, sheetName
    
    'Make note that the account was added
    If Not newAccountDict.Exists(dictKey) Then
        newAccountDict.Add dictKey, NewAccountDictMessage(tdAccount.Number)
    End If
End Sub

Private Sub AddBeneToAccount(bene As clsBeneficiary, accountNode As IXMLDOMNode, newBeneDict As Dictionary, dictKey As String)
    'Add the beneficiary as its own node to the client list
    XMLWrite.AddBeneficiaryToNode bene, accountNode, sheetName
    
    'Make note that the beneficiary was added unless the account was already added in the process
    If Not newBeneDict.Exists(dictKey) Then
        'See if the account name is available
        Dim accountName As String
        If Not accountNode.SelectSingleNode("Name") Is Nothing Then
            accountName = accountNode.SelectSingleNode("Name").Text
        End If
        
        newBeneDict.Add dictKey, NewBeneDictMessage(bene.NameOfBeneficiary, bene.account.Number, accountName)
    End If
End Sub

Private Sub UpdateAccountNode(tdAccount As clsAccount, accountNode As IXMLDOMNode, errorDict As Dictionary, dictKey As String)
    With tdAccount
        'Update the information in the account node
        Dim listAccountType As String, listCloseDate As String
        XMLUpdate.UpdateAccountFromBeneList accountNode, .custodian, .openDate, .closeDate, .TypeOfAccount, listCloseDate, listAccountType
                          
        'Check for differing info
        Dim IdentifyingData As String
        If Not accountNode.SelectSingleNode("Name") Is Nothing Then
            IdentifyingData = accountNode.SelectSingleNode("Name").Text
        Else
            IdentifyingData = .Number
        End If
        CheckForDifferingInfo "Close_Date", .closeDate, listCloseDate, errorDict(diffCloseDateDictName), dictKey, accountNode, IdentifyingData
        CheckForDifferingInfo "Type", .TypeOfAccount, listAccountType, errorDict(diffAccountTypeDictName), dictKey, accountNode, IdentifyingData
    End With
End Sub

Private Sub UpdateBeneNode(bene As clsBeneficiary, benenode As IXMLDOMNode, errorDict As Dictionary, dictKey As String)
    With bene
        'Update the information in the beneficiary node
        Dim listLevel As String
        Dim listPercent As Double
        XMLUpdate.UpdateBeneFromBeneList benenode, .Level, .Percent, listLevel, listPercent
        
        'Check for differing info
        Dim IdentifyingData As String
        IdentifyingData = .account.Number & " - " & .NameOfBeneficiary
        CheckForDifferingInfo "Level", .Level, listLevel, errorDict(diffLevelDictName), dictKey, benenode, IdentifyingData
        CheckForDifferingInfo "Percent", .Percent, listPercent, errorDict(diffLevelDictName), dictKey, benenode, IdentifyingData
    End With
End Sub

Private Function ConvertAccountType(accountType As String) As String
    'Get the TD account types and the respective Morningstar types
    Dim accountTypes As Dictionary
    Set accountTypes = New Dictionary
    With accountTypes
        .Add "403b7-TDA", "403(b)"
        .Add "CORPORATION", "Corporate"
        .Add "Coverdell ESA", "Coverdell ESA"
        .Add "ESTATE", "Estate"
        .Add "INDIVIDUAL", "Individual"
        .Add "Individual 401k-TDA", "401(k)"
        .Add "Individual Roth 401k-TDA", "Roth 401(k)"
        .Add "IRA", "Traditional IRA"
        .Add "JOINT TENANTS WROS", "Joint (JTWROS)"
        .Add "Minor Roth IRA", "Roth IRA"
        .Add "NON-PROFIT ORGANIZATION", "Non-Profit Organization-ex"
        .Add "ROLLOVER IRA", "Traditional Rollover IRA"
        .Add "Roth Bene IRA", "IRA - Roth Beneficiary"
        .Add "ROTH IRA", "Roth IRA"
        .Add "SEP IRA", "SEP IRA"
        .Add "SIMPLE IRA", "SIMPLE IRA"
        .Add "TENANTS BY THE ENTIRETY", "Tenants in Entirety"
        .Add "Traditional Bene IRA", "Beneficiary IRA"
        .Add "TRUST", "Trust"
        .Add "UGMA/UTMA", "UGMA/UTMA"
    End With
    
    'Convert the account type, if possible
    If accountTypes.Exists(accountType) Then
        ConvertAccountType = accountTypes(accountType)
    Else
        ConvertAccountType = accountType
    End If
End Function

Private Function IsInArray(searchTerm As String, arr() As String) As Boolean
    'Check the array for the search term
    Dim filteredList() As String
    filteredList = Filter(arr, searchTerm)
    Dim name As Integer
    Do While name <= UBound(filteredList) And Not IsInArray
        If filteredList(name) = searchTerm Then
            'The search term was found in the array
            IsInArray = True
        End If
        name = name + 1
    Loop
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
