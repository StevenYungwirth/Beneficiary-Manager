Attribute VB_Name = "ImportIndividualSheets"
Option Explicit
Private Const removedNodesDictName As String = "Removed Elements"
Private Const msAccountErrorDictName As String = "MS Account"
Private Const msExportErrorDictName As String = "MS Export"
Private Const rtAccountErrorDictName As String = "RT Account"
Private Const rtContactErrorDictName As String = "RT Contact"
Private Const beneListErrorDictName As String = "TD Bene"
Private Const manualBeneErrorDictName As String = "Manual Sheet"
Private Const notesFileName As String = "Z:\FPIS - Operations\Beneficiary Project\Archive\Logs\Log"

Private Property Get SubDictList() As String()
    Dim returnArr(0 To 6) As String
    returnArr(0) = removedNodesDictName
    returnArr(1) = msExportErrorDictName
    returnArr(2) = msAccountErrorDictName
    returnArr(3) = rtAccountErrorDictName
    returnArr(4) = rtContactErrorDictName
    returnArr(5) = beneListErrorDictName
    returnArr(6) = manualBeneErrorDictName
    SubDictList = returnArr
End Property

Private Property Get InSheetAttributeNames() As String()
    Dim returnArr(0 To 5) As String
    returnArr(0) = "In_MS_Export"
    returnArr(1) = "In_MS_Accounts"
    returnArr(2) = "In_RT_Accounts"
    returnArr(3) = "In_RT_Contacts"
    returnArr(4) = "In_Bene_List"
    returnArr(5) = "In_Manual_Sheet"
    InSheetAttributeNames = returnArr
End Property

Private Property Get XMLClientList() As DOMDocument60
    Set XMLClientList = ProjectGlobals.ClientListFile
End Property

Public Sub ImportNewSheets()
    'Get the client list node
    Dim clientList As IXMLDOMElement
    Set clientList = XMLClientList.SelectSingleNode("Client_List")
    
    'Check for existence
    If clientList Is Nothing Then
        MsgBox """Client_List"" node not found in client list XML."
        Exit Sub
    End If
    
    'Save this client list in the archive folder
    Dim createDate As String
    createDate = clientList.getAttribute("Create_Date")
    If createDate <> vbNullString Then
        'Uncomment for actual file
        XMLClientList.Save ArchiveFolder & "Households " & Replace(createDate, "/", "-") & ".xml"
    End If
    
    'Set the create date
    clientList.setAttribute "Create_Date", Format(ProjectGlobals.ImportTime, "yyyy/mm/dd")

    'Load each sheet that could be imported (doesn't need to be every sheet)
    Dim msExport As clsMSHouseholdExport
    Dim msAccounts As clsMSAccountList
    Dim rtAccounts As clsRTAccountList
    Dim rtContacts As clsRTContactList
    Dim tdaBeneList As clsTDABeneList
    Dim manualSheet As clsManualSheet
    LoadSheets msExport, msAccounts, rtAccounts, rtContacts, tdaBeneList, manualSheet

    'Set up dictionary for error messages that come up while importing sheets
    Dim errorDict As Dictionary
    Set errorDict = SetupErrorDictionary

    'Start the timer
    Timer.TimeStart

    'Import each individual sheet
    If Not msExport Is Nothing Then Set errorDict(msExportErrorDictName) = msExport.ImportToXML
    If Not msAccounts Is Nothing Then Set errorDict(msAccountErrorDictName) = msAccounts.ImportToXML
    If Not rtContacts Is Nothing Then Set errorDict(rtContactErrorDictName) = rtContacts.ImportToXML
'    If Not rtAccounts Is Nothing Then Set errordict(rtAccountErrorDictName) = rtAccounts.ImportToXML
    If Not tdaBeneList Is Nothing Then Set errorDict(beneListErrorDictName) = tdaBeneList.ImportToXML
    If Not manualSheet Is Nothing Then Set errorDict(manualBeneErrorDictName) = manualSheet.ImportToXML

    'Log the time
    Timer.TimeEnd

    'Reconcile any nodes that were added to the client list
    'Can the parent nodes be found? (The order of the import functions should make this unneeded)

    'Update household/member/account active statuses
    UpdateStatuses

    'Delete any nodes that aren't in any list
    DeleteLeftoverElements errorDict(removedNodesDictName)

    'Alert if an account doesn't have a household
    'Alert if MS account type is TOD (Should be individual or joint)
    'Alert if qualified account doesn't have beneficiaries
    'Alert if account name has TOD but there's no beneficiaries
    'Alert about components if they're equal and in the same parent - ask to delete?
    '(possible if account is created through RT Accounts and moved with MS Accounts?)
    'Alert if owner name doesn't seem right
    'Alert if account's owner name is "Multiple"
    'Alert if RT contact doesn't have a household
    'Alert if TD converted account type doesn't match MS account type
    'TODO alphabatize all comboboxes

    'Report errors found
    'TODO have a log for each sheet instead of all in one?
    ReportLogs errorDict
    
    'Format, save, and close the XML file
    XMLProcedures.FormatAndSaveXML
    CloseXMLClientList
    
    'Log in the workbook the file paths of the importing sheets

    'Open the error log
    'TODO prompt to open the log(s)
    CreateObject("Shell.Application").Open (notesFileName & " " & Format(ProjectGlobals.ImportTime, "yyyy-mm-dd") & ".txt")
    
    'Show confirmation of import
    MsgBox "Worksheets have been successfully imported."
    Debug.Print "Import complete"
End Sub

Private Sub CloseXMLClientList()
    ProjectGlobals.CloseClientFile
End Sub

Private Sub UpdateStatuses()
    'Update account statuses
    Debug.Print "Updating account statuses"
    Dim accountNodes As IXMLDOMNodeList
    Set accountNodes = XMLClientList.SelectNodes("//Account")
    Dim accountNode As Integer
    For accountNode = 0 To accountNodes.Length - 1
        UpdateAccountNodeActive accountNodes(accountNode)
    Next accountNode
    
    'Update member statuses
    Debug.Print "Updating member statuses"
    Dim memberNodes As IXMLDOMNodeList
    Set memberNodes = XMLClientList.SelectNodes("//Member")
    Dim memberNode As Variant
    For memberNode = 0 To memberNodes.Length - 1
        UpdateMemberNodeActive memberNodes(memberNode)
    Next memberNode
    
    'Update household statuses
    Debug.Print "Updating household statuses"
    Dim householdNodes As IXMLDOMNodeList
    Set householdNodes = XMLClientList.SelectNodes("//Household")
    Dim householdNode As Integer
    For householdNode = 0 To householdNodes.Length - 1
        UpdateHouseholdNodeActive householdNodes(householdNode)
    Next householdNode
    
'    XMLProcedures.FormatAndSaveXML
End Sub

Private Sub UpdateAccountNodeActive(accountNode As IXMLDOMElement)
    'Account is active if it has a non-zero balance (default custodian) or balance greater than 1 (held-away)
    '1 is arbitrary; it can't be 0 because an account can close but closing transactions may still leave fractional shares in the positions
    Dim accountBalance As Double
    Dim accountCustodian As String
    accountBalance = accountNode.SelectSingleNode("Balance").Text
    accountCustodian = accountNode.SelectSingleNode("Custodian").Text
    accountNode.SelectSingleNode("Active").Text _
    = ((accountBalance > 0 And accountCustodian = ProjectGlobals.DefaultCustodian) _
    Or (accountBalance > 1 And accountCustodian <> ProjectGlobals.DefaultCustodian))
End Sub

Private Sub UpdateMemberNodeActive(memberNode As IXMLDOMElement)
    'Check if member has a date of death and is deceased
    Dim isMemberDeceased As Boolean
    isMemberDeceased = (memberNode.SelectSingleNode("Date_of_Death").Text <> vbNullString)
    
    'Check that the member's status is a possible Redtail active status
    Dim activeStatusOptions() As String
    activeStatusOptions = ProjectGlobals.RedtailActiveStatuses
    Dim isMemberActiveInRT As Boolean: isMemberActiveInRT = False
    Dim statusOption As Integer
    Do While Not isMemberActiveInRT And statusOption <= UBound(activeStatusOptions)
        isMemberActiveInRT = (memberNode.SelectNodes("Status[text()='" & activeStatusOptions(statusOption) & "']").Length > 0)
        statusOption = statusOption + 1
    Loop
    
    'Check each account to see if at least one is active
    Dim hasActiveAccount As Boolean: hasActiveAccount = False
    Dim activeAccounts As IXMLDOMNodeList
    Set activeAccounts = memberNode.SelectNodes("Account/Active[text()='True']")
    hasActiveAccount = (activeAccounts.Length > 0)
    
    'Set the member node's active property. Member is active if they're not deceased, have at least one active account,
    'and they have one of the possible Redtail active statuses (or the status is blank)
    memberNode.SelectSingleNode("Active").Text = Not isMemberDeceased And hasActiveAccount _
    And (isMemberActiveInRT Or memberNode.SelectSingleNode("Status").Text = vbNullString)
End Sub

Private Sub UpdateHouseholdNodeActive(householdNode As IXMLDOMElement)
    'Household is active if there is at least one active member
    Dim activeMembers As IXMLDOMNodeList
    Set activeMembers = householdNode.SelectNodes("./Member/Active[text()='True']")
    householdNode.SelectSingleNode("Active").Text = (activeMembers.Length > 0)
End Sub

Private Function SetupErrorDictionary() As Dictionary
    'Initialize dictionaries for each sheet and add them to the dictionary
    Dim errorDict As Dictionary: Set errorDict = New Dictionary
    Dim subDict As Variant
    For Each subDict In SubDictList
        If Not errorDict.Exists(subDict) Then
            errorDict.Add subDict, New Dictionary
        End If
    Next subDict
    
    'Return the dictionary
    Set SetupErrorDictionary = errorDict
End Function

Private Sub LoadSheets(msExport As clsMSHouseholdExport, msAccounts As clsMSAccountList, rtAccounts As clsRTAccountList, rtContacts As clsRTContactList, _
                       tdaBeneList As clsTDABeneList, manualSheet As clsManualSheet)
    'Open the form
    Load frmImport
    frmImport.Show
    
    'Get each sheet from the form
    With frmImport
        Set tdaBeneList = ClassConstructor.NewTDABeneList(.TDABeneFile)
        Set msAccounts = ClassConstructor.NewMSAccountList(.MSAccountsFile)
        Set rtAccounts = ClassConstructor.NewRTAccountList(.RTAccountsFile)
        Set rtContacts = ClassConstructor.NewRTContactList(.RTContactsFile)
        Set msExport = ClassConstructor.NewMSHouseholdExport(.MSHouseholdExportFile)
    End With
    
    'Unload the form
    Unload frmImport
    
    'Get the sheet with manual beneficiaries
    Set manualSheet = ClassConstructor.NewManualSheet
End Sub

Private Sub ReportLogs(errorDict As Dictionary)
    Debug.Print "Logging errors to file"

    'Get the error file
    Dim fsobj As FileSystemObject: Set fsobj = New FileSystemObject
    Dim logFile As TextStream
    Set logFile = fsobj.OpenTextFile(notesFileName & " " & Format(ProjectGlobals.ImportTime, "yyyy-mm-dd") & ".txt", ForWriting, True)

    'TODO put all errors into an array and then print all at once for better performance
    'Log everything in the error dictionary
    'Iterate through each sheet
    Dim dictKey As Variant
    For Each dictKey In errorDict.Keys
        'Log the sheet name
        logFile.WriteLine dictKey
        Debug.Print "Logging " & dictKey
        
        'Log each sheet's error dictionaries
        Dim sheetDict As Dictionary: Set sheetDict = errorDict(dictKey)
        Dim sheetDictKey As Variant
        For Each sheetDictKey In sheetDict.Keys
            If TypeOf sheetDict(sheetDictKey) Is Dictionary Then
                'Log each error dictionary's items
                Dim logDict As Dictionary: Set logDict = sheetDict(sheetDictKey)
                Dim logItem As Variant
                For Each logItem In logDict.Items
                    'Write the log string
                    logFile.WriteLine logItem
                Next logItem
            Else
                'Write the log string
                logFile.WriteLine sheetDict(sheetDictKey)
            End If
        Next sheetDictKey
        
        'Add a space between each sheet's logs
        logFile.WriteBlankLines 1
    Next dictKey
    
    'Close the log file
    logFile.Close
End Sub

Private Sub DeleteLeftoverElements(errorDict As Dictionary)
    'Delete the nodes not present in any importing sheet
    RemoveNodes GetDeletedNodes("Household"), errorDict
    RemoveNodes GetDeletedNodes("Member"), errorDict
    RemoveNodes GetDeletedNodes("Account"), errorDict
    RemoveNodes GetDeletedNodes("Beneficiary"), errorDict
End Sub

Private Sub RemoveNodes(nodeList As IXMLDOMNodeList, deletedNodes As Dictionary)
    'For each node, remove it from its parent
    Dim Node As Variant
    For Each Node In nodeList
        Dim selectedNode As IXMLDOMNode
        If Not Node Is Nothing Then
            Set selectedNode = Node
            If Not selectedNode.parentNode Is Nothing Then
                'Log that the node and its children (if applicable) were deleted
                Dim logNote As String: logNote = vbNullString
                If selectedNode.BaseName = "Household" Then
                    logNote = DeletedHouseholdNote(selectedNode)
                ElseIf selectedNode.BaseName = "Member" Then
                    logNote = DeletedMemberNote(selectedNode)
                ElseIf selectedNode.BaseName = "Account" Then
                    logNote = DeletedAccountNote(selectedNode)
                ElseIf selectedNode.BaseName = "Beneficiary" Then
                    logNote = DeletedBeneNote(selectedNode)
                End If
                deletedNodes.Add deletedNodes.count, logNote
                
                'Remove the node from its parent
                selectedNode.parentNode.RemoveChild selectedNode
            End If
        End If
    Next Node
End Sub

Private Function DeletedHouseholdNote(householdNode As IXMLDOMNode) As String
    DeletedHouseholdNote = "Deleted Household: " & householdNode.SelectSingleNode("Name").Text & " - with all its members, accounts, and beneficiaries"
End Function

Private Function DeletedMemberNote(memberNode As IXMLDOMNode) As String
    Dim returnNote As String
    returnNote = "Deleted Member: " & XMLRead.MemberNameFromNode(memberNode)
    returnNote = returnNote & " within " & memberNode.SelectSingleNode("../Name").Text & " - with all its accounts and beneficiaries"
    DeletedMemberNote = returnNote
End Function

Private Function DeletedAccountNote(accountNode As IXMLDOMNode) As String
    Dim returnNote As String
    returnNote = "Deleted Account: " & accountNode.SelectSingleNode("Name").Text & " | " & accountNode.SelectSingleNode("Number").Text & " within "
    
    Dim householdName As String, memberName As String
    If accountNode.parentNode.parentNode Is Nothing Then
        'Account is by itself and isn't attached to a member or household
        householdName = "No household"
        memberName = "No member"
    ElseIf accountNode.parentNode.BaseName = "Member" Then
        householdName = accountNode.SelectSingleNode("../../Name").Text
        memberName = XMLRead.MemberNameFromNode(accountNode.parentNode)
    End If
    
    returnNote = returnNote & householdName & "/" & memberName & " - with all its beneficiaries"
    DeletedAccountNote = returnNote
End Function

Private Function DeletedBeneNote(benenode As IXMLDOMNode) As String
    Dim returnNote As String
    returnNote = "Deleted Beneficiary: " & benenode.SelectSingleNode("Name").Text & " | " & benenode.SelectSingleNode("Level").Text & " | " & benenode.SelectSingleNode("Percent").Text & " within "
    
    Dim householdName As String, memberName As String
    If benenode.parentNode.parentNode.parentNode Is Nothing Then
        'Bene's account is by itself and isn't attached to a member or household
        householdName = "No household"
        memberName = "No member"
    ElseIf benenode.parentNode.parentNode.BaseName = "Member" Then
        householdName = benenode.SelectSingleNode("../../../Name").Text
        memberName = XMLRead.MemberNameFromNode(benenode.parentNode.parentNode)
    End If
    
    returnNote = returnNote & householdName & "/" & memberName & "/" & benenode.SelectSingleNode("../Name").Text & " | " & benenode.SelectSingleNode("../Number").Text
    
    DeletedBeneNote = returnNote
End Function

Private Function GetDeletedNodes(elementName As String) As IXMLDOMNodeList
    'Build the query to look for elements with In_[sheet] attributes being false or missing
    Dim attributeNames() As String
    attributeNames = InSheetAttributeNames
    Dim deleteString As String
    deleteString = "["
    Dim inSheetName As Integer
    For inSheetName = 0 To UBound(attributeNames)
        If inSheetName <> 0 Then
            deleteString = deleteString & " and "
        End If
        deleteString = deleteString & "(@" & attributeNames(inSheetName) & "='False' or not(@" & attributeNames(inSheetName) & "))"
    Next inSheetName
    deleteString = deleteString & "]"

    'Add in empty members that don't have accounts anymore
    deleteString = deleteString & " | //Member[./Full_Name[text()='EmptyMember'] and count(./Account) = 0]"

    'Get the elements not in any sheet
    Dim elementsToDelete As IXMLDOMNodeList
    Set elementsToDelete = XMLClientList.SelectNodes("//" & elementName & deleteString)
                           
    'Return the elements
    Set GetDeletedNodes = elementsToDelete
End Function
