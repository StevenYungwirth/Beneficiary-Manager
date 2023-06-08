Attribute VB_Name = "XMLUpdate"
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

Public Sub UpdateHouseholdFromMSExport(householdNode As IXMLDOMElement, householdName As String, householdID As String, ByRef returnIDInList As String, _
                                       ByRef returnNameInList As String)
    'Update the ID and household name
    returnIDInList = ReturnOldAndSetNewComponentAttribute(householdNode, "Morningstar_ID", householdID, MSExportName)
    returnNameInList = ReturnOldAndSetNewComponentAttribute(householdNode, "Name", householdName, MSExportName)
    
    'Mark household as being in MS export list
    XMLProcedures.FlagNodeInList householdNode, MSExportName, True
End Sub

Public Sub UpdateAccountFromMSExport(accountNode As IXMLDOMNode, accountID As String, accountName As String, accountNumber As String, _
                                     accountType As String, accountCustodian As String, ByRef returnIDInList As String, ByRef returnNameInList As String, _
                                     ByRef returnNumberInList As String, ByRef returnTypeInList As String, ByRef returnCustodianInList As String)
    'Update the ID, account name, and type
    returnIDInList = ReturnOldAndSetNewComponentAttribute(accountNode, "Morningstar_ID", accountID, MSExportName)
    returnNameInList = ReturnOldAndSetNewComponentAttribute(accountNode, "Name", accountName, MSExportName)
    returnTypeInList = ReturnOldAndSetNewComponentAttribute(accountNode, "Type", accountType, MSExportName)
    
    'Declare the higher priorty sheet array
    Dim higherPrioritySheets(0) As String
    higherPrioritySheets(0) = BeneListName
    
    'Update the number and custodian
    returnNumberInList = ReturnOldAndSetNewComponentAttribute(accountNode, "Number", accountNumber, MSExportName, higherPrioritySheets)
    returnCustodianInList = ReturnOldAndSetNewComponentAttribute(accountNode, "Custodian", accountCustodian, MSExportName, higherPrioritySheets)
    
    'Mark account as being in MS export list
    XMLProcedures.FlagNodeInList accountNode, MSExportName, True
End Sub

Public Sub UpdateHouseholdFromMSAccounts(householdNode As IXMLDOMElement, householdName As String, ByRef returnNameInList As String)
    'Update the household name
    returnNameInList = ReturnOldAndSetNewComponentAttribute(householdNode, "Name", householdName, MSAccountName)
    
    'Mark household as being in MS account list
    XMLProcedures.FlagNodeInList householdNode, MSAccountName, True
End Sub

Public Sub UpdateMemberFromMSAccounts(memberNode As IXMLDOMElement, firstName As String, lastName As String, fullName As String, ByRef returnFirstNameInList As String, _
                                      ByRef returnLastNameInList As String)
    'Declare the higher priorty sheet array
    Dim higherPrioritySheets(0) As String
    higherPrioritySheets(0) = RTContactName
    
    'Update the first, last, and full names
    returnFirstNameInList = ReturnOldAndSetNewComponentAttribute(memberNode, "First_Name", firstName, MSAccountName, higherPrioritySheets)
    returnLastNameInList = ReturnOldAndSetNewComponentAttribute(memberNode, "Last_Name", lastName, MSAccountName, higherPrioritySheets)
    ReturnOldAndSetNewComponentAttribute memberNode, "Full_Name", fullName, MSAccountName, higherPrioritySheets
    
    'Mark member as being in MS account list
    XMLProcedures.FlagNodeInList memberNode, MSAccountName, True
End Sub

Public Sub UpdateAccountFromMSAccountList(accountNode As IXMLDOMElement, accountName As String, accountType As String, accountBalance As Double, _
                                          ownerName As String, accountDiscretionary As Boolean, accountCustodian As String, ByRef returnAccountNameInList As String, _
                                          ByRef returnAccountTypeInList As String, ByRef returnOwnerNameInList As String, ByRef returnDiscretionaryInList As Boolean, _
                                          ByRef returnCustodianInList As String)
    'Update the name, type, owner name, discretionary, and balance
    returnAccountNameInList = ReturnOldAndSetNewComponentAttribute(accountNode, "Name", accountName, MSAccountName)
    returnAccountTypeInList = ReturnOldAndSetNewComponentAttribute(accountNode, "Type", accountType, MSAccountName)
    returnOwnerNameInList = ReturnOldAndSetNewComponentAttribute(accountNode, "Owner", ownerName, MSAccountName)
        returnDiscretionaryInList = CBool(ReturnOldAndSetNewComponentAttribute(accountNode, "Discretionary", accountDiscretionary, MSAccountName))
    ReturnOldAndSetNewComponentAttribute accountNode, "Balance", accountBalance, MSAccountName
    
    'Declare the higher priorty sheet array
    Dim higherPrioritySheets(0) As String
    higherPrioritySheets(0) = BeneListName
    
    'Update the custodian
    returnCustodianInList = ReturnOldAndSetNewComponentAttribute(accountNode, "Custodian", accountCustodian, MSAccountName, higherPrioritySheets)
    
    'Mark account as being in MS account list
    XMLProcedures.FlagNodeInList accountNode, MSAccountName, True
End Sub

Public Sub UpdateAccountFromRTAccountList(accountNode As IXMLDOMElement, redtailID As Double, accountType As String, accountCustodian As String, _
                                          accountNumber As String, ByRef returnRedtailIDInList As Double, ByRef returnAccountTypeInList As String, _
                                          ByRef returnAccountCustodianInList As String, ByRef returnAccountNumberInList As String)
    'Update the Redtail ID
    returnRedtailIDInList = ReturnOldAndSetNewComponentAttribute(accountNode, "Redtail_ID", redtailID, RTAccountName)
    
    'Declare the higher priorty sheet array
    Dim higherPrioritySheets(3) As String
    higherPrioritySheets(0) = MSAccountName
    higherPrioritySheets(1) = MSExportName
    
    'Update the account type
    returnAccountTypeInList = ReturnOldAndSetNewComponentAttribute(accountNode, "Type", accountType, RTAccountName, higherPrioritySheets)
    
    'Update the account custodian
    higherPrioritySheets(2) = BeneListName
    returnAccountCustodianInList = ReturnOldAndSetNewComponentAttribute(accountNode, "Custodian", accountCustodian, RTAccountName, higherPrioritySheets)
    
    'Update the account number
    higherPrioritySheets(3) = ManualBeneListName
    returnAccountNumberInList = ReturnOldAndSetNewComponentAttribute(accountNode, "Number", accountNumber, RTAccountName, higherPrioritySheets)
    
    'Mark account as being in RT Account list
    XMLProcedures.FlagNodeInList accountNode, RTAccountName, True
End Sub

Public Sub UpdateHouseholdFromRTContactList(householdNode As IXMLDOMElement, householdName As String, ByRef returnHouseholdNameInList As String)
    'Declare the higher priorty sheet array
    Dim higherPrioritySheets(0) As String
    higherPrioritySheets(0) = MSAccountName
    
    'Update the household name
    returnHouseholdNameInList = ReturnOldAndSetNewComponentAttribute(householdNode, "Name", householdName, RTContactName, higherPrioritySheets)
    
    'Mark household as being in RT Contact list
    XMLProcedures.FlagNodeInList householdNode, RTContactName, True
End Sub

Public Sub UpdateMemberFromRTContactList(memberNode As IXMLDOMElement, redtailID As Long, firstName As String, lastName As String, memberStatus As String, _
                                         dateOfDeath As String, ByRef returnFirstNameInList As String, ByRef returnLastNameInList As String, _
                                         ByRef returnStatusInList As String, ByRef returnDateOfDeath As String)
    'Update the ID, name (trim it first), status, and date of death
    returnFirstNameInList = ReturnOldAndSetNewComponentAttribute(memberNode, "First_Name", Trim(firstName), RTContactName)
    returnLastNameInList = ReturnOldAndSetNewComponentAttribute(memberNode, "Last_Name", Trim(lastName), RTContactName)
    returnStatusInList = ReturnOldAndSetNewComponentAttribute(memberNode, "Status", memberStatus, RTContactName)
    returnDateOfDeath = ReturnOldAndSetNewComponentAttribute(memberNode, "Date_of_Death", dateOfDeath, RTContactName)
    
    'Mark member as being in RT Contact list
    XMLProcedures.FlagNodeInList memberNode, RTContactName, True
End Sub

Public Sub UpdateAccountFromBeneList(accountNode As IXMLDOMElement, custodian As String, openDate As String, closeDate As String, accountType As String, _
                                     ByRef returnCloseDateInList As String, ByRef returnAccountTypeInList As String)
    'Update custodian, open date, close date
    ReturnOldAndSetNewComponentAttribute accountNode, "Custodian", custodian, BeneListName
    ReturnOldAndSetNewComponentAttribute accountNode, "Open_Date", openDate, BeneListName
    returnCloseDateInList = ReturnOldAndSetNewComponentAttribute(accountNode, "Close_Date", closeDate, BeneListName)
    
    'Declare the higher priorty sheet array
    Dim higherPrioritySheets(1) As String
    higherPrioritySheets(0) = MSAccountName
    higherPrioritySheets(1) = MSExportName
    
    'Update account type
    returnAccountTypeInList = ReturnOldAndSetNewComponentAttribute(accountNode, "Type", accountType, BeneListName, higherPrioritySheets)
    
    'Mark account as being in the bene list
    XMLProcedures.FlagNodeInList accountNode, BeneListName, True
End Sub

Public Sub UpdateBeneFromBeneList(benenode As IXMLDOMElement, beneLevel As String, benePercent As Double, ByRef returnBeneLevel As String, _
                                  ByRef returnBenePercent As Double)
    'Update the level and percent
    returnBeneLevel = ReturnOldAndSetNewComponentAttribute(benenode, "Level", beneLevel, BeneListName)
    returnBenePercent = ReturnOldAndSetNewComponentAttribute(benenode, "Percent", benePercent, BeneListName)

    'Mark beneficiary as being in the bene list
    XMLProcedures.FlagNodeInList benenode, BeneListName, True
End Sub

Public Sub UpdateBeneFromManualSheet(benenode As IXMLDOMElement, beneName As String, beneLevel As String, benePercent As Double, _
                                     lastUpdated As String, updatedBy As String, ByRef returnBeneName As String, ByRef returnBeneLevel As String, _
                                     ByRef returnBenePercent As Double)
    'Update the name, level, percent, and last updated/by
    returnBeneName = ReturnOldAndSetNewComponentAttribute(benenode, "Name", beneName, BeneListName)
    returnBeneLevel = ReturnOldAndSetNewComponentAttribute(benenode, "Level", beneLevel, BeneListName)
    returnBenePercent = ReturnOldAndSetNewComponentAttribute(benenode, "Percent", benePercent, BeneListName)
    ReturnOldAndSetNewComponentAttribute benenode, "Last_Updated", lastUpdated, BeneListName
    ReturnOldAndSetNewComponentAttribute benenode, "Updated_By", updatedBy, BeneListName

    'Mark beneficiary as being in the bene list
    XMLProcedures.FlagNodeInList benenode, ManualBeneListName, True
End Sub

Public Function UpdateBeneficiaryFromForm(beneToUpdate As IXMLDOMElement, NewBene As clsBeneficiary) As Boolean
    'Update the beneficiary
    With beneToUpdate
        ReturnOldAndSetNewComponentAttribute beneToUpdate, "Name", NewBene.NameOfBeneficiary, ProjectGlobals.m_manualBeneListName
        ReturnOldAndSetNewComponentAttribute beneToUpdate, "Level", NewBene.Level, ProjectGlobals.m_manualBeneListName
        ReturnOldAndSetNewComponentAttribute beneToUpdate, "Percent", NewBene.Percent, ProjectGlobals.m_manualBeneListName
        ReturnOldAndSetNewComponentAttribute beneToUpdate, "Last_Updated", NewBene.addDate, ProjectGlobals.m_manualBeneListName
        ReturnOldAndSetNewComponentAttribute beneToUpdate, "Updated_By", NewBene.AddedBy, ProjectGlobals.m_manualBeneListName
    End With
    UpdateBeneficiaryFromForm = True
End Function

Private Function ReturnOldAndSetNewComponentAttribute(componentNode As IXMLDOMElement, attributeName As String, attributeValueToSet As Variant, sheetName As String, _
                                                      Optional higherPrioritySheets As Variant) As Variant
    'Set the old value in the list
    Dim returnValue As Variant
    returnValue = ReturnComponentAttribute(componentNode, attributeName)

    'Get the sheet that added this attribute
    Dim componentChild As IXMLDOMElement
    Set componentChild = componentNode.SelectSingleNode(attributeName)
    Dim attributeAddedBy As String
    If Not componentChild Is Nothing Then
        If Not IsNull(componentChild.getAttribute("Added_By")) Then
            attributeAddedBy = componentChild.getAttribute("Added_By")
        End If
    End If
    
    'Determine if a sheet with higher priority added the attribute value
    Dim canSetAttribute As Boolean: canSetAttribute = True
    If Not IsMissing(higherPrioritySheets) Then
        Dim sht As Integer
        For sht = LBound(higherPrioritySheets) To UBound(higherPrioritySheets)
            If higherPrioritySheets(sht) = attributeAddedBy And higherPrioritySheets(sht) <> vbNullString Then
                canSetAttribute = False
            End If
        Next sht
    End If
    
    'Set the new value if possible
    If canSetAttribute Then
        'Add the node if it's not present
        Dim attributeNode As IXMLDOMNode
        If componentNode.SelectSingleNode(attributeName) Is Nothing Then
            'The property node is missing. Create and append it
            Set attributeNode = ProjectGlobals.ClientListFile.createNode(1, attributeName, vbNullString)
            componentNode.appendChild attributeNode
        End If
    
        'Set the new value in the list
        If VarType(attributeValueToSet) = vbDate Then
            componentNode.SelectSingleNode(attributeName).Text = Format(attributeValueToSet, "dd/mm/yyyy")
        Else
            componentNode.SelectSingleNode(attributeName).Text = CStr(attributeValueToSet)
        End If
        
        'If the value changed, change the last update date
        If returnValue <> attributeValueToSet Then
            SetUpdatedBy componentNode.SelectSingleNode(attributeName), sheetName
        End If
    
        'If the value is being added, set which sheet added the attribute
        If returnValue = vbNullString Or attributeAddedBy = vbNullString Then
            SetAddedBy componentNode.SelectSingleNode(attributeName), sheetName
        End If
    End If
    
    'Return the old value in the list
    If returnValue = vbNullString And VarType(attributeValueToSet) = vbBoolean Then
        'Return false if the node has no text and the attribute is a boolean
        ReturnOldAndSetNewComponentAttribute = False
    ElseIf returnValue = vbNullString And VarType(attributeValueToSet) = vbDate Then
        'Return default date if the node has no text and the attribute is a date
        ReturnOldAndSetNewComponentAttribute = CDate(0)
    Else
        'Return the text from the node
        ReturnOldAndSetNewComponentAttribute = returnValue
    End If
End Function

Private Function ReturnComponentAttribute(componentNode As IXMLDOMElement, childName As String) As Variant
    'Check that the child node exists
    Dim componentChild As IXMLDOMElement
    Set componentChild = componentNode.SelectSingleNode(childName)

    'Return the value in the list
    If Not componentChild Is Nothing Then
        'The child exists, return its text
        ReturnComponentAttribute = componentChild.Text
    End If
End Function

Private Sub SetAddedBy(attributeNode As IXMLDOMNode, adderName As String)
    Dim attributeElement As IXMLDOMElement: Set attributeElement = attributeNode
    attributeElement.setAttribute "Added_By", adderName
    attributeElement.setAttribute "Added_On", Format(ProjectGlobals.ImportTime, "m/d/yyyy h:mm;@")
End Sub

Private Sub SetUpdatedBy(attributeNode As IXMLDOMNode, updaterName As String)
    Dim attributeElement As IXMLDOMElement: Set attributeElement = attributeNode
    attributeElement.setAttribute "Updated_By", updaterName
    attributeElement.setAttribute "Last_Updated", Format(ProjectGlobals.ImportTime, "m/d/yyyy h:mm;@")
End Sub
