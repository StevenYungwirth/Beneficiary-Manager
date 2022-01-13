Attribute VB_Name = "SampleExports"
Option Explicit

Private Sub CreateSampleExports()
    'Load client list
    Dim clientList As DOMDocument60
    Set clientList = XMLReadWrite.LoadClientList(isSample:=True)
    
    'Load the export sheets
    Dim manualSheet As clsManualSheet, tdSheet As clsTDABeneList, msAccounts As clsMSAccountList, rtAccounts As clsRTAccountList, rtContacts As clsRTContactList
    Set manualSheet = ClassConstructor.NewManualSheet
    Set tdSheet = ClassConstructor.NewTDABeneList(vbNullString)
    Set msAccounts = ClassConstructor.NewMSAccountList(vbNullString)
    Set rtAccounts = ClassConstructor.NewRTAccountList(vbNullString)
    Set rtContacts = ClassConstructor.NewRTContactList(vbNullString)
    
    'Get the list of accounts
    Dim accountNodes As IXMLDOMNodeList
    Set accountNodes = clientList.SelectNodes("//Account")
    
    'Put each account onto the TDA Bene List (If TD account), Manual Beneficiary (If not), MS Account, and RT Account sheets
    Dim tdCount As Integer
    Dim accountNode As Integer
    For accountNode = 0 To accountNodes.Length - 1
        PutAccountOntoSheets accountNodes(accountNode), accountNode + 1, tdCount, manualSheet, tdSheet, msAccounts, rtAccounts
    Next accountNode
    
    'Get the list of members
    Dim memberNodes As IXMLDOMNodeList
    Set memberNodes = clientList.SelectNodes("//Member")
    
    'Put each member onto the RT Contact sheet
    Dim memberNode As Integer
    For memberNode = 0 To memberNodes.Length - 1
        AddMemberToRTSheet memberNodes(memberNode), memberNode + 1, rtContacts
    Next memberNode
    
    'Fill the sheets
    Application.ScreenUpdating = False
    manualSheet.FillWorksheet ThisWorkbook.Worksheets("Manual Beneficiaries")
    tdSheet.FillWorksheet ThisWorkbook.Worksheets("TDA Bene List")
    msAccounts.FillWorksheet ThisWorkbook.Worksheets("MS Accounts")
    rtAccounts.FillWorksheet ThisWorkbook.Worksheets("RT Accounts")
    rtContacts.FillWorksheet ThisWorkbook.Worksheets("RT Contacts")
    Application.ScreenUpdating = True
End Sub

Private Sub PutAccountOntoSheets(accountNode As IXMLDOMElement, accountNumber As Integer, ByRef tdCount As Integer, manualSheet As clsManualSheet, _
    tdaBeneList As clsTDABeneList, msSheet As clsMSAccountList, rtAccount As clsRTAccountList)
    'Get the account's custodian
    Dim custodian As String
    custodian = accountNode.getAttribute("Custodian")
    
    'Add the account to the appropriate beneficiary sheet
    If custodian = "TD Ameritrade Institutional" Then
        tdCount = tdCount + 1
        AddAccountToTDASheet accountNode, tdCount, tdaBeneList
    Else
        AddAccountToManualSheet accountNode, manualSheet
    End If
    
    'Add the account to the Morningstar and Redtail sheets
    AddAccountToMSSheet accountNode, accountNumber, msSheet
    AddAccountToRTSheet accountNode, accountNumber, rtAccount
End Sub

Private Sub AddAccountToManualSheet(accountNode As IXMLDOMElement, manualSheet As clsManualSheet)
    'Add each beneficiary to the sheet
    Dim beneficiaries As IXMLDOMNodeList
    Set beneficiaries = accountNode.SelectNodes("Beneficiary")
    Dim beneficiary As Integer
    For beneficiary = 0 To beneficiaries.Length - 1
        AddBeneToManualSheet beneficiaries(beneficiary), manualSheet
    Next beneficiary
End Sub

Private Sub AddBeneToManualSheet(beneNode As IXMLDOMElement, manualSheet As clsManualSheet)
    'Get the beneficiary's account node
    Dim parentAccount As IXMLDOMElement
    Set parentAccount = beneNode.parentNode
    
    With manualSheet
        'Add the account's attributes to the sheet
        .SetData newData:=parentAccount.getAttribute("Name"), headerName:="Account Name/ID", datapoint:=-1
        .SetData newData:=parentAccount.getAttribute("Number"), headerName:="Account#", datapoint:=-1
        .SetData newData:=parentAccount.getAttribute("Redtail_ID"), headerName:="Account ID", datapoint:=-1
        
        'Add the beneficiary's attributes to the sheet
        .SetData newData:=beneNode.getAttribute("Name"), headerName:="Name", datapoint:=-1
        .SetData newData:=beneNode.getAttribute("Level"), headerName:="BeneLevel", datapoint:=-1
        .SetData newData:=beneNode.getAttribute("Percent"), headerName:="Percentage", datapoint:=-1
        .SetData newData:="Added", headerName:="Action", datapoint:=-1
        .SetData newData:=beneNode.getAttribute("Added_On"), headerName:="Added", datapoint:=-1
        .SetData newData:=vbNullString, headerName:="By", datapoint:=-1
    End With
End Sub

Private Sub AddAccountToTDASheet(accountNode As IXMLDOMElement, accountNumber As Integer, tdSheet As clsTDABeneList)
    'Add each account to the sheet
    With tdSheet
        'Add each beneficiary to the sheet
        Dim beneficiaries As IXMLDOMNodeList
        Set beneficiaries = accountNode.SelectNodes("Beneficiary")
        
        If beneficiaries.Length = 0 Then
            'Add the account's attributes to the sheet
            .SetData newData:=accountNode.getAttribute("Number"), headerName:="Account#", datapoint:=accountNumber
            .SetData newData:=accountNode.getAttribute("Type"), headerName:="AcctDescription", datapoint:=accountNumber
            If CBool(accountNode.getAttribute("Active")) = False Then
                .SetData newData:="1/1/1990", headerName:="DateClosed", datapoint:=accountNumber
            End If
        Else
            Dim beneficiary As Integer
            For beneficiary = 0 To beneficiaries.Length - 1
                AddBeneToTDSheet beneficiaries(beneficiary), accountNumber, tdSheet
            Next beneficiary
        End If
    End With
End Sub

Private Sub AddBeneToTDSheet(beneNode As IXMLDOMElement, accountNumber As Integer, tdSheet As clsTDABeneList)
    'Get the beneficiary's account node
    Dim parentAccount As IXMLDOMElement
    Set parentAccount = beneNode.parentNode
    
    With tdSheet
        'Add the account's attributes to the sheet
        .SetData newData:=parentAccount.getAttribute("Number"), headerName:="Account#", datapoint:=accountNumber
        .SetData newData:=parentAccount.getAttribute("Type"), headerName:="AcctDescription", datapoint:=accountNumber
        If CBool(parentAccount.getAttribute("Active")) = False Then
            .SetData newData:="1/1/1990", headerName:="DateClosed", datapoint:=accountNumber
        End If
        
        'Add the beneficiary's attributes to the sheet
        .SetData newData:=beneNode.getAttribute("Name"), headerName:="Name", datapoint:=accountNumber
        .SetData newData:=beneNode.getAttribute("Level"), headerName:="BeneLevel", datapoint:=accountNumber
        .SetData newData:=beneNode.getAttribute("Percent"), headerName:="Percentage", datapoint:=accountNumber
    End With
End Sub

Private Sub AddAccountToMSSheet(accountNode As IXMLDOMElement, accountNumber As Integer, msAccounts As clsMSAccountList)
    'Add each account to the sheet
    With msAccounts
        'Add each account to the sheet
        .SetData newData:=accountNode.getAttribute("Name"), headerName:="Account Name/ID", datapoint:=accountNumber
        .SetData newData:=accountNode.getAttribute("Number"), headerName:="Account Number", datapoint:=accountNumber
        .SetData newData:=accountNode.getAttribute("Custodian"), headerName:="Current Custodian", datapoint:=accountNumber
        .SetData newData:=accountNode.getAttribute("Balance"), headerName:="Market Value" & Chr(10) & "USD", datapoint:=accountNumber
        .SetData newData:=accountNode.getAttribute("Type"), headerName:="Account Type", datapoint:=accountNumber
        
        'Add the owner name
        Dim accountOwner As IXMLDOMElement
        Set accountOwner = accountNode.parentNode
        .SetData newData:=accountOwner.getAttribute("Last_Name") & ", " & accountOwner.getAttribute("First_Name"), headerName:="Account Owner", datapoint:=accountNumber
        
        'Add the household name
        Dim accountHousehold As IXMLDOMElement
        Set accountHousehold = accountOwner.parentNode
        .SetData newData:=accountHousehold.getAttribute("Name"), headerName:="Client / Prospect Name", datapoint:=accountNumber
    End With
End Sub

Private Sub AddAccountToRTSheet(accountNode As IXMLDOMElement, accountNumber As Integer, rtAccounts As clsRTAccountList)
    'Add each account to the sheet
    With rtAccounts
        'Add the account to the sheet
        .SetData newData:=accountNode.getAttribute("Number"), headerName:="Account Number", datapoint:=accountNumber
        .SetData newData:=accountNode.getAttribute("Custodian"), headerName:="Company", datapoint:=accountNumber
        .SetData newData:=accountNode.getAttribute("Type"), headerName:="Type", datapoint:=accountNumber
        
        'Add the owner name
        Dim accountOwner As IXMLDOMElement
        Set accountOwner = accountNode.parentNode
        .SetData newData:=accountOwner.getAttribute("Last_Name") & ", " & accountOwner.getAttribute("First_Name"), headerName:="Contact Name", datapoint:=accountNumber
    End With
End Sub

Private Sub AddMemberToRTSheet(memberNode As IXMLDOMElement, memberNumber As Integer, rtContacts As clsRTContactList)
    'Add each member to the sheet
    With rtContacts
        'Add the member to the sheet
        .SetData newData:=memberNode.getAttribute("First_Name"), headerName:="First Name", datapoint:=memberNumber
        .SetData newData:=memberNode.getAttribute("Last_Name"), headerName:="Last Name", datapoint:=memberNumber
        .SetData newData:=GetStatus(memberNode), headerName:="Status", datapoint:=memberNumber
        If GetStatus(memberNode) = "Deceased" Then
            .SetData newData:="1/1/1990", headerName:="Date Of Death", datapoint:=memberNumber
        End If
        
        'Add the household name
        Dim memberHousehold As IXMLDOMElement
        Set memberHousehold = memberNode.parentNode
        .SetData newData:=memberHousehold.getAttribute("Name"), headerName:="Family Name", datapoint:=memberNumber
    End With
End Sub

Private Function GetStatus(memberNode As IXMLDOMElement) As String
    Dim isActive As Boolean
    Dim isDeceased As Boolean
    isActive = CBool(memberNode.getAttribute("Active"))
    isDeceased = CBool(memberNode.getAttribute("Deceased"))
    
    If isActive Then
        GetStatus = "Active"
    ElseIf isDeceased Then
        GetStatus = "Deceased"
    Else
        GetStatus = "InActive"
    End If
End Function
