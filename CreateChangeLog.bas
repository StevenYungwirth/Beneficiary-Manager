Attribute VB_Name = "CreateChangeLog"
Option Explicit

Public Sub CreateSheet(addedElements As Dictionary, deletedNodes As IXMLDOMNodeList)
    'Turn off screen updating
    Application.ScreenUpdating = False
    On Error GoTo BackOn

    'Create the file for logging the changes - the first worksheet is for added elements, the second is for removed elements
    Dim logBook As Workbook
    Set logBook = Application.Workbooks.Add
    Dim addSheet As Worksheet, removedSheet As Worksheet
    Set addSheet = logBook.Worksheets(1)
    addSheet.name = "Added"
    Set removedSheet = logBook.Worksheets.Add(after:=addSheet)
    removedSheet.name = "Removed"
    
    'Add the sections for each element type to each sheet
    AddSections addSheet
    AddSections removedSheet
    
    'Add the elements to their respective sheets
    AddAddedToSheet addedElements, addSheet
    AddRemovedToSheet deletedNodes, removedSheet
    
    'Sort the sections
    SortSections addSheet
    SortSections removedSheet
    
    'Save the workbook
    Dim fileName As String
    fileName = XMLReadWrite.ClientListFolder & "Change Log - " & Replace(Date, "/", "-")
    logBook.SaveAs fileName:=fileName
    
    'Close the workbook
    logBook.Close savechanges:=False
    
    'Turn screen updating back on
    Application.ScreenUpdating = True
    
    On Error GoTo 0
    Exit Sub
BackOn:
    Application.ScreenUpdating = True
End Sub

Private Sub AddSections(sht As Worksheet)
    With sht
        'Add named ranges
        With .Names
            .Add sht.name & "HouseholdStart", sht.Range("A1")
            .Add sht.name & "HouseholdEnd", sht.Range("A1")
            .Add sht.name & "MemberStart", sht.Range("A3")
            .Add sht.name & "MemberEnd", sht.Range("A3")
            .Add sht.name & "AccountStart", sht.Range("A5")
            .Add sht.name & "AccountEnd", sht.Range("A5")
            .Add sht.name & "BeneficiaryStart", sht.Range("A7")
            .Add sht.name & "BeneficiaryEnd", sht.Range("A7")
        End With
        
        'Add household header
        .Range(.name & "HouseholdStart").Value2 = "Household"
        
        'Add member headers
        With .Range(.name & "MemberStart")
            .Value2 = "Member"
            .Offset(0, 1).Value2 = "Household Name"
        End With
        
        'Add account headers
        With .Range(.name & "AccountStart")
            .Value2 = "Account"
            .Offset(0, 1).Value2 = "Number"
            .Offset(0, 2).Value2 = "Type"
            .Offset(0, 3).Value2 = "Account Owner"
            .Offset(0, 4).Value2 = "Owner Household"
        End With
        
        'Add beneficiary headers
        With .Range(.name & "BeneficiaryStart")
            .Value2 = "Beneficiary"
            .Offset(0, 1).Value2 = "Level"
            .Offset(0, 2).Value2 = "Percent"
            .Offset(0, 3).Value2 = "Account Name"
            .Offset(0, 4).Value2 = "Account Owner"
            .Offset(0, 5).Value2 = "Owner Household"
        End With
    End With
End Sub

Private Sub SortSections(sht As Worksheet)
    With sht
        Dim householdStart As Range, householdEnd As Range
        Set householdStart = .Range(sht.name & "HouseholdStart")
        Set householdEnd = .Range(sht.name & "HouseholdEnd")
        .Range(householdStart, householdEnd).sort key1:=householdStart, Header:=xlYes
        
        Dim memberStart As Range, memberEnd As Range
        Set memberStart = .Range(sht.name & "MemberStart")
        Set memberEnd = .Range(sht.name & "MemberEnd").Offset(0, 1)
        .Range(memberStart, memberEnd).sort key1:=memberStart.Offset(0, 1), key2:=memberStart, Header:=xlYes
        
        Dim accountStart As Range, accountEnd As Range
        Set accountStart = .Range(sht.name & "AccountStart")
        Set accountEnd = .Range(sht.name & "AccountEnd").Offset(0, 4)
        .Range(accountStart, accountEnd).sort key1:=accountStart.Offset(0, 4), key2:=accountStart.Offset(0, 3), key3:=accountStart, Header:=xlYes
        
        Dim beneficiaryStart As Range, beneficiaryEnd As Range
        Set beneficiaryStart = .Range(sht.name & "BeneficiaryStart")
        Set beneficiaryEnd = .Range(sht.name & "BeneficiaryEnd").Offset(0, 5)
        .Range(beneficiaryStart, beneficiaryEnd).sort key1:=beneficiaryStart.Offset(0, 3), key2:=beneficiaryStart, Header:=xlYes
    End With
End Sub

Private Sub AddAddedToSheet(addedElements As Dictionary, sht As Worksheet)
    'For each node, add it to the corresponding listbox
    Dim category As Integer
    For category = 0 To 3
        If addedElements.Items(category).count > 0 Then
            Dim elements As Collection
            Set elements = addedElements.Items(category)
            Dim element As Variant
            For Each element In elements
                If elements.count > 0 And addedElements.Keys(category) = "Households" Then
                    Dim household As clsHousehold
                    Set household = element
                    AddToHouseholdRange sht, household
                    AddMembersToSheet sht, household.Members
                ElseIf elements.count > 0 And addedElements.Keys(category) = "Members" Then
                    Dim member As clsMember
                    Set member = element
                    AddToMemberRange sht, member
                    AddAccountsToSheet sht, member.accounts
                ElseIf elements.count > 0 And addedElements.Keys(category) = "Accounts" Then
                    Dim account As clsAccount
                    Set account = element
                    AddToAccountRange sht, account
                    AddBenesToSheet sht, account.Benes
                ElseIf elements.count > 0 And addedElements.Keys(category) = "Beneficiaries" Then
                    Dim bene As clsBeneficiary
                    Set bene = element
                    AddToBeneRange sht, bene
                End If
            Next element
        End If
    Next category
    
'    'Sort the listboxes
'    SortListBox lbxHouseholdsAdded
'    SortListBox lbxMembersAdded
'    SortListBox lbxAccountsAdded
'    SortListBox lbxBenesAdded
End Sub

Private Sub AddRemovedToSheet(removedElements As IXMLDOMNodeList, sht As Worksheet)
    'For each node, add it to the corresponding listbox
    Dim node As Variant
    For Each node In removedElements
        Dim SelectedNode As IXMLDOMElement
        Set SelectedNode = node
        With SelectedNode
            If .BaseName = "Household" Then
                AddToHouseholdRange sht, XMLReadWrite.ReadHouseholdFromNode(SelectedNode)
            ElseIf .BaseName = "Member" Then
                AddToMemberRange sht, XMLReadWrite.ReadMemberFromNode(SelectedNode)
            ElseIf .BaseName = "Account" Then
                AddToAccountRange sht, XMLReadWrite.ReadAccountFromNode(SelectedNode)
            ElseIf .BaseName = "Beneficiary" Then
                AddToBeneRange sht, XMLReadWrite.ReadBeneficiaryFromNode(SelectedNode)
            End If
        End With
    Next node
    
'    'Sort the listboxes
'    SortListBox lbxHouseholdsRemoved
'    SortListBox lbxMembersRemoved
'    SortListBox lbxAccountsRemoved
'    SortListBox lbxBenesRemoved
End Sub

Private Sub AddToHouseholdRange(sht As Worksheet, household As clsHousehold)
    'Set the named range
    Dim namedRange As String
    namedRange = sht.name & "HouseholdEnd"

    'Add the household to the worksheet
    sht.Range(namedRange).Offset(1, 0).EntireRow.Insert Shift:=xlDown
    sht.Names(namedRange).RefersTo = sht.Range(namedRange).Offset(1, 0)
    sht.Range(namedRange).Value2 = household.NameOfHousehold
End Sub

Private Sub AddToMemberRange(sht As Worksheet, member As clsMember)
    'Set the named range
    Dim namedRange As String
    namedRange = sht.name & "MemberEnd"

    'Add the member to the worksheet
    sht.Range(namedRange).Offset(1, 0).EntireRow.Insert Shift:=xlDown
    sht.Names(namedRange).RefersTo = sht.Range(namedRange).Offset(1, 0)
    sht.Range(namedRange).Value2 = member.FName & " " & member.LName
    sht.Range(namedRange).Offset(0, 1).Value2 = member.ContainingHousehold.NameOfHousehold
End Sub

Private Sub AddToAccountRange(sht As Worksheet, account As clsAccount)
    'Set the named range
    Dim namedRange As String
    namedRange = sht.name & "AccountEnd"

    'Add the account to the worksheet
    sht.Range(namedRange).Offset(1, 0).EntireRow.Insert Shift:=xlDown
    sht.Names(namedRange).RefersTo = sht.Range(namedRange).Offset(1, 0)
    With sht.Range(namedRange)
        .Value2 = account.NameOfAccount
        .Offset(0, 1).Value2 = account.Number
        .Offset(0, 2).Value2 = account.TypeOfAccount
        .Offset(0, 3).Value2 = account.Owner.NameOfMember
        .Offset(0, 4).Value2 = account.Owner.ContainingHousehold.NameOfHousehold
    End With
End Sub

Private Sub AddToBeneRange(sht As Worksheet, bene As clsBeneficiary, Optional accountOwner As String, Optional householdName As String)
    'Set the named range
    Dim namedRange As String
    namedRange = sht.name & "BeneficiaryEnd"

    'Add the bene to the worksheet
    sht.Range(namedRange).Offset(1, 0).EntireRow.Insert Shift:=xlDown
    sht.Names(namedRange).RefersTo = sht.Range(namedRange).Offset(1, 0)
    With sht.Range(namedRange)
        .Value2 = bene.NameOfBeneficiary
        .Offset(0, 1).Value2 = bene.Level
        .Offset(0, 2).Value2 = bene.Percent
        .Offset(0, 3).Value2 = bene.account.NameOfAccount
        .Offset(0, 4).Value2 = bene.account.Owner.NameOfMember
        .Offset(0, 5).Value2 = bene.account.Owner.ContainingHousehold.NameOfHousehold
    End With
End Sub

Private Sub AddMembersToSheet(sht As Worksheet, Members As Dictionary)
    Dim member As Variant
    For Each member In Members.Items
        Dim memberItem As clsMember
        Set memberItem = member
        AddToMemberRange sht, memberItem
        AddAccountsToSheet sht, memberItem.accounts
    Next member
End Sub

Private Sub AddAccountsToSheet(sht As Worksheet, accounts As Dictionary)
    Dim account As Variant
    For Each account In accounts.Items
        Dim accountItem As clsAccount
        Set accountItem = account
        AddToAccountRange sht, accountItem
        AddBenesToSheet sht, accountItem.Benes
    Next account
End Sub

Private Sub AddBenesToSheet(sht As Worksheet, Benes As Collection)
    Dim bene As Variant
    For Each bene In Benes
        Dim beneItem As clsBeneficiary
        Set beneItem = bene
        AddToBeneRange sht, beneItem
    Next bene
End Sub
