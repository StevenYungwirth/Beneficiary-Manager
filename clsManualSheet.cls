VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsManualSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_datasheet As clsDataSheet
Private Const worksheetName As String = "Manual Beneficiaries"
Private Const accountNameHeader As String = "Account Name/ID"
Private Const accountNumberHeader As String = "Account#"
Private Const morningstarIDHeader As String = "Morningstar ID"
Private Const beneIDHeader As String = "Bene ID"
Private Const beneNameHeader As String = "Name"
Private Const beneLevelHeader As String = "BeneLevel"
Private Const benePercentHeader As String = "Percentage"
Private Const actionHeader As String = "Action"
Private Const addDateHeader As String = "Added"
Private Const addedByHeader As String = "By"
Private Const newBeneDictName As String = "New Beneficiary"
Private Const dupeBeneDictName As String = "Duplicate Beneficiary"
Private Const diffBeneNameDictName As String = "Different Bene Name"
Private Const diffLevelDictName As String = "Different Bene Level"
Private Const diffPercentDictName As String = "Different Bene Percent"
Private Const invalidActionDictName As String = "Invalid Action"

Private Property Get sheetName() As String
    sheetName = ProjectGlobals.m_manualBeneListName
End Property

Private Property Get ErrorDictKeys() As String()
    Dim keyArray(0 To 5) As String
    keyArray(0) = newBeneDictName
    keyArray(1) = dupeBeneDictName
    keyArray(2) = diffBeneNameDictName
    keyArray(3) = diffLevelDictName
    keyArray(4) = diffPercentDictName
    keyArray(5) = invalidActionDictName
    ErrorDictKeys = keyArray
End Property

Private Property Get InvalidActionDictMessage(rw As Integer, action As String) As String
    InvalidActionDictMessage = "Row " & rw + 1 & " has an invalid action: " & action
End Property

Private Property Get NewBeneDictMessage(beneName As String, accountNumber As String, accountName As String) As String
    If accountNumber = vbNullString Then
        NewBeneDictMessage = "Beneficiary Added - " & beneName & " | Account - " & accountName
    Else
        NewBeneDictMessage = "Beneficiary Added - " & beneName & " | Account - " & accountNumber & " - " & accountName
    End If
End Property

Private Property Get DupeBeneDictMessage(accountName As String, beneName As String) As String
    DupeBeneDictMessage = "Beneficiary Duplicate - " & beneName & " - Multiple instances found in account " & accountName
End Property

Private Property Get XMLClientList() As DOMDocument60
    Set XMLClientList = ProjectGlobals.ClientListFile
End Property

Public Sub ClassBuilder()
    'Set the required headers
    Dim requiredHeaders() As String
    requiredHeaders = Split(accountNameHeader & "," & accountNumberHeader & "," & morningstarIDHeader & "," _
                          & beneIDHeader & "," & beneNameHeader & "," & beneLevelHeader & "," & benePercentHeader & "," _
                          & actionHeader & "," & addDateHeader & "," & addedByHeader, ",")
    
    'Set the datasheet in order to get the data
    Set m_datasheet = ClassConstructor.NewDataSheet(filePath:=vbNullString, wkstName:=worksheetName, reqHeaders:=requiredHeaders)
End Sub

Public Function ImportToXML(Optional testingLimit As Integer) As Dictionary
    'Contains master beneficiary ID for heldaway benes
    'Contains account name, account number, Redtail account ID, beneficiary info, added time/by
    
    'Set up dictionaries for the possible errors to return
    Dim errorDict As Dictionary
    Set errorDict = SetUpDictionaries

    'Flag every account and beneficiary as not in the list
    FlagAsNotInList "Account"
    FlagAsNotInList "Beneficiary"
    
    'Get a dictionary of the beneficiaries
    Dim beneDict As Dictionary
    Set beneDict = GetBeneficiaries(errorDict(invalidActionDictName))
    
    'Iterate through each beneficiary in the dictionary
    Dim beneItem As Integer
    For beneItem = 0 To beneDict.count - 1
        If beneItem Mod 100 = 0 Then
            DoEvents
            Debug.Print sheetName & " " & beneItem + 1 & " / " & beneDict.count
            Application.StatusBar = sheetName & " " & beneItem + 1 & " / " & beneDict.count
        End If
        
        'Update or add the beneficiary
        Dim bene As clsBeneficiary
        Set bene = beneDict.Items(beneItem)
        UpdateBeneficiaryNodes bene, errorDict
        
        'Exit the loop if the testing limit has been reached
        If beneItem = testingLimit - 1 Then: beneItem = beneDict.count - 1
    Next beneItem
    
    'Reset the status bar
    Application.StatusBar = False
    
    Set ImportToXML = errorDict
End Function

Public Function GetData(headerName As String, datapoint As Integer) As String
    GetData = m_datasheet.GetData(headerName, datapoint)
End Function

Private Function GetBeneficiaries(invalidActionDict As Dictionary) As Dictionary
    'Declare a dictionary to return
    Dim beneDict As Dictionary: Set beneDict = New Dictionary

    'Iterate through each row
    Dim rw As Integer
    For rw = 1 To m_datasheet.DataSize - 1
        If rw Mod 100 = 0 Then
            DoEvents
            Debug.Print "Reading " & sheetName & " row " & rw & " / " & m_datasheet.DataSize
        End If
        
        'Get the beneficiary from the row
        Dim rowBene As clsBeneficiary
        Set rowBene = GetBeneFromRow(rw)
        
        'If it's already in the dictionary (bene id), do the corresponding action to it
        If beneDict.Exists(rowBene.id) Then
            'Get the action from the row
            Dim action As String
            action = GetData(actionHeader, rw)
            
            'Perform the corresponding action
            Dim actionPerformed As Boolean
            actionPerformed = False
            If action = "Update" Then
                actionPerformed = UpdateBene(rowBene, beneDict(rowBene.id))
            ElseIf action = "Delete" Then
                actionPerformed = DeleteBene(rowBene, beneDict)
            ElseIf action = "Confirm" Then
                actionPerformed = ConfirmBene(beneDict(rowBene.id), rowBene.UpdatedDate, rowBene.updatedBy)
            End If
            
            If Not actionPerformed Then
                'Make a note that the action on the row isn't a valid option
                '(This should theoretically only happen if someone messes with the spreadsheet)
                invalidActionDict.Add rowBene.id, InvalidActionDictMessage(rw, action)
            End If
        Else
            'Otherwise, add it to the dictionary
            beneDict.Add rowBene.id, rowBene
        End If
    Next rw
    
    'Return the dictionary
    Set GetBeneficiaries = beneDict
End Function

Private Sub UpdateBeneficiaryNodes(bene As clsBeneficiary, errorDict As Dictionary)
    'Attempt to find the account(s) in the XML
    Dim accountNodeList As IXMLDOMNodeList
    With bene.account
        Set accountNodeList = XMLRead.FindAccounts(accountNumber:=.Number, accountName:=.NameOfAccount, morningstarID:=.morningstarID)
    End With
    
    'Set the key for the dictionaries
    Dim dictKey As String
    dictKey = bene.id
    
    'Iterate through each account
    Dim accountNodeFound As Variant
    For Each accountNodeFound In accountNodeList
        Dim accountNode As IXMLDOMNode: Set accountNode = accountNodeFound
        
        'Attempt to find the beneficiary(ies) in the account
        Dim beneNodeList As IXMLDOMNodeList
        Set beneNodeList = XMLRead.FindBenesInAccount(accountNode, bene.NameOfBeneficiary, bene.Level, bene.Percent, bene.id)
    
        If beneNodeList.Length = 0 Then
            'The beneficiary wasn't found. Add it to the account
            AddBeneToAccount bene, accountNode, errorDict(newBeneDictName), dictKey
        Else
            'At least one beneficiary node was found. Update the info for each one
            Dim nodeFound As Integer
            For nodeFound = 0 To beneNodeList.Length - 1
                UpdateBeneNode bene, beneNodeList(nodeFound), errorDict, dictKey
            Next nodeFound
    
            'Check for duplicate beneficiaries
            If beneNodeList.Length > 1 And Not errorDict.Item(dupeBeneDictName).Exists(dictKey) Then
                'Make note that there are multiple account nodes with the same name/number
                errorDict.Item(dupeBeneDictName).Add dictKey, DupeBeneDictMessage(bene.account.NameOfAccount, bene.NameOfBeneficiary)
            End If
        End If
    Next accountNodeFound
End Sub

Private Sub AddBeneToAccount(bene As clsBeneficiary, accountNode As IXMLDOMNode, newBeneDict As Dictionary, dictKey As String)
    'Add the beneficiary as its own node to the client list
    XMLWrite.AddBeneficiaryToNode bene, accountNode, sheetName
    
    'Make note that the beneficiary was added unless the account was already added in the process
    If Not newBeneDict.Exists(dictKey) And XMLProcedures.GetAddDate(accountNode) < ProjectGlobals.ImportTime Then
        'See if the account name is available
        Dim accountName As String
        If Not accountNode.SelectSingleNode("Name") Is Nothing Then
            accountName = accountNode.SelectSingleNode("Name").Text
        End If
        
        newBeneDict.Add dictKey, NewBeneDictMessage(bene.NameOfBeneficiary, bene.account.Number, accountName)
    End If
End Sub

Private Sub UpdateBeneNode(bene As clsBeneficiary, benenode As IXMLDOMNode, errorDict As Dictionary, dictKey As String)
    'Update the information in the beneficiary node
    Dim listName As String, listLevel As String
    Dim listPercent As Double
    With bene
        XMLUpdate.UpdateBeneFromManualSheet benenode, .NameOfBeneficiary, .Level, .Percent, .UpdatedDate, .updatedBy, listName, listLevel, listPercent
    
        'Check for differing info
        Dim IdentifyingData As String
        IdentifyingData = .account.NameOfAccount & " - " & .NameOfBeneficiary
        CheckForDifferingInfo "Name", .NameOfBeneficiary, listName, errorDict(diffBeneNameDictName), dictKey, benenode, IdentifyingData
        CheckForDifferingInfo "Level", .Level, listLevel, errorDict(diffLevelDictName), dictKey, benenode, IdentifyingData
        CheckForDifferingInfo "Percent", .Percent, listPercent, errorDict(diffLevelDictName), dictKey, benenode, IdentifyingData
    End With
End Sub

Private Function UpdateBene(beneWithUpdatedInfo As clsBeneficiary, beneWithOldInfo As clsBeneficiary) As Boolean
    With beneWithOldInfo
        'Update the information
        .account.NameOfAccount = beneWithUpdatedInfo.account.NameOfAccount
        .NameOfBeneficiary = beneWithUpdatedInfo.NameOfBeneficiary
        .Level = beneWithUpdatedInfo.Level
        .Percent = beneWithUpdatedInfo.Percent
        .UpdatedDate = beneWithUpdatedInfo.UpdatedDate
        .updatedBy = beneWithUpdatedInfo.updatedBy
        UpdateBene = True
    End With
End Function

Private Function DeleteBene(beneToDelete As clsBeneficiary, beneDict As Dictionary) As Boolean
    If beneDict.Exists(beneToDelete.id) Then
        'Remove the beneficiary from the dictionary
        beneDict.Remove beneToDelete.id
        DeleteBene = True
    End If
End Function

Private Function ConfirmBene(beneToConfirm As clsBeneficiary, updateDate As String, updateBy As String) As Boolean
    With beneToConfirm
        If updateDate > .UpdatedDate Then
            .UpdatedDate = updateDate
            .updatedBy = updateBy
            ConfirmBene = True
        End If
    End With
End Function

Private Function GetAccountFromRow(rw As Integer, accountDict As Dictionary) As clsAccount
    'Get the account name and number
    Dim accountName As String, accountNumber As String
    accountName = GetData(accountNameHeader, rw)
    accountNumber = GetData(accountNumberHeader, rw)
    
    'Find the account in the dictionary
    If accountDict.Exists(accountNumber) Then
        Set GetAccountFromRow = accountDict(accountNumber)
    ElseIf accountDict.Exists(accountName & accountNumber) Then
        Set GetAccountFromRow = accountDict(accountName & accountNumber)
    End If
End Function

Private Function GetBeneFromRow(rw As Integer) As clsBeneficiary
    'Initialize the beneficiary using the row's data
    Set GetBeneFromRow = ClassConstructor.NewBene(GetData(beneNameHeader, rw), GetData(beneLevelHeader, rw), GetData(benePercentHeader, rw))
    
    'Get the beneficiary ID, account name and number
    With GetBeneFromRow
        .id = GetData(beneIDHeader, rw)
        .addDate = GetData(addDateHeader, rw)
        .UpdatedDate = GetData(addDateHeader, rw)
        .AddedBy = GetData(addedByHeader, rw)
        .updatedBy = GetData(addedByHeader, rw)
        With .account
            .NameOfAccount = GetData(accountNameHeader, rw)
            .Number = GetData(accountNumberHeader, rw)
            .morningstarID = GetData(morningstarIDHeader, rw)
        End With
    End With
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
    'If the account/bene name has single or double quotes, wrap the search string in the opposite
    SearchWrapper = XMLProcedures.XPathExpression(value)
End Function
