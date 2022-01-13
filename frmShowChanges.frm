VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmShowChanges 
   Caption         =   "XML File Changes"
   ClientHeight    =   12480
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   14640
   OleObjectBlob   =   "frmShowChanges.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmShowChanges"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const memberColumnWidths As String = "100,164"
Private Const accountColumnWidths As String = "200, 60, 80, 100, 140"
Private Const beneColumnWidths As String = "200, 30, 40, 200, 100, 164"

Private Sub UserForm_Initialize()
    'Start the form in the middle of the screen with Excel
    Me.StartUpPosition = 0
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)
    
    'Set the listbox headers
    InitializeHouseholdBox lbxHouseholdsAdded
    InitializeHouseholdBox lbxHouseholdsRemoved
    InitializeMemberBox lbxMembersAdded
    InitializeMemberBox lbxMembersRemoved
    InitializeAccountBox lbxAccountsAdded
    InitializeAccountBox lbxAccountsRemoved
    InitializeBeneBox lbxBenesAdded
    InitializeBeneBox lbxBenesRemoved
End Sub

Private Sub btnOK_Click()
    'Hide the form
    Me.Hide
End Sub

Private Sub btnCancel_Click()
    'Hide the form and end the macro
    Me.Hide
    End
End Sub

Private Sub InitializeHouseholdBox(box As MSForms.ListBox)
    'Set the column header
    box.AddItem "Name"
End Sub

Private Sub InitializeMemberBox(box As MSForms.ListBox)
    'Set the column headers and widths
    With box
        .AddItem
        .List(0, 0) = "Name"
        .List(0, 1) = "Household Name"
        .ColumnWidths = memberColumnWidths
    End With
End Sub

Private Sub InitializeAccountBox(box As MSForms.ListBox)
    'Set the column headers and widths
    With box
        .AddItem
        .List(0, 0) = "Name"
        .List(0, 1) = "Number"
        .List(0, 2) = "Type"
        .List(0, 3) = "Owner"
        .List(0, 4) = "Household Name"
        .ColumnWidths = accountColumnWidths
    End With
End Sub

Private Sub InitializeBeneBox(box As MSForms.ListBox)
    'Set the column headers and widths
    With box
        .AddItem
        .List(0, 0) = "Name"
        .List(0, 1) = "Level"
        .List(0, 2) = "Percent"
        .List(0, 3) = "Account Name"
        .List(0, 4) = "Account Owner"
        .List(0, 5) = "Household Name"
        .ColumnWidths = beneColumnWidths
    End With
End Sub

Public Sub ShowAddedNodes(addedNodes As Dictionary)
    'For each node, add it to the corresponding listbox
    Dim category As Integer
    For category = 0 To 3
        If addedNodes.Items(category).count > 0 Then
            Dim elements As Collection
            Set elements = addedNodes.Items(category)
            Dim element As Variant
            For Each element In elements
                If elements.count > 0 And addedNodes.Keys(category) = "Households" Then
                    Dim household As clsHousehold
                    Set household = element
                    AddToHouseholdBox lbxHouseholdsAdded, household
                    AddMembersToBox household.Members
                ElseIf elements.count > 0 And addedNodes.Keys(category) = "Members" Then
                    Dim member As clsMember
                    Set member = element
                    AddToMemberBox lbxMembersAdded, member
                    AddAccountsToBox member.accounts
                ElseIf elements.count > 0 And addedNodes.Keys(category) = "Accounts" Then
                    Dim account As clsAccount
                    Set account = element
                    AddToAccountBox lbxAccountsAdded, account
                    AddBenesToBox account.Benes
                ElseIf elements.count > 0 And addedNodes.Keys(category) = "Beneficiaries" Then
                    Dim bene As clsBeneficiary
                    Set bene = element
                    AddToBeneBox lbxBenesAdded, bene
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

Public Sub ShowDeletedNodes(deletedNodes As IXMLDOMNodeList)
    'For each node, add it to the corresponding listbox
    Dim node As Variant
    For Each node In deletedNodes
        Dim SelectedNode As IXMLDOMElement
        Set SelectedNode = node
        With SelectedNode
            If .BaseName = "Household" Then
                AddToHouseholdBox lbxHouseholdsRemoved, XMLReadWrite.ReadHouseholdFromNode(SelectedNode)
            ElseIf .BaseName = "Member" Then
                AddToMemberBox lbxMembersRemoved, XMLReadWrite.ReadMemberFromNode(SelectedNode)
            ElseIf .BaseName = "Account" Then
                AddToAccountBox lbxAccountsRemoved, XMLReadWrite.ReadAccountFromNode(SelectedNode)
            ElseIf .BaseName = "Beneficiary" Then
                AddToBeneBox lbxBenesRemoved, XMLReadWrite.ReadBeneficiaryFromNode(SelectedNode)
            End If
        End With
    Next node
    
'    'Sort the listboxes
'    SortListBox lbxHouseholdsRemoved
'    SortListBox lbxMembersRemoved
'    SortListBox lbxAccountsRemoved
'    SortListBox lbxBenesRemoved
End Sub

Private Sub AddMembersToBox(Members As Dictionary)
    Dim member As Variant
    For Each member In Members.Items
        Dim memberItem As clsMember
        Set memberItem = member
        AddToMemberBox lbxMembersAdded, memberItem
        AddAccountsToBox memberItem.accounts
    Next member
End Sub

Private Sub AddAccountsToBox(accounts As Dictionary)
    Dim account As Variant
    For Each account In accounts.Items
        Dim accountItem As clsAccount
        Set accountItem = account
        AddToAccountBox lbxAccountsAdded, accountItem
        AddBenesToBox accountItem.Benes
    Next account
End Sub

Private Sub AddBenesToBox(Benes As Collection)
    Dim bene As Variant
    For Each bene In Benes
        Dim beneItem As clsBeneficiary
        Set beneItem = bene
        AddToBeneBox lbxBenesAdded, beneItem
    Next bene
End Sub

Private Sub AddToHouseholdBox(householdBox As MSForms.ListBox, household As clsHousehold)
    'Add the household to the listbox
    householdBox.AddItem household.NameOfHousehold
End Sub

Private Sub AddToMemberBox(memberBox As MSForms.ListBox, member As clsMember)
    'Add the member to the listbox
    With memberBox
        Dim topOfList As Integer
        topOfList = .ListCount
        .AddItem
        .List(topOfList, 0) = member.FName & " " & member.LName
        .List(topOfList, 1) = member.ContainingHousehold.NameOfHousehold
    End With
End Sub

Private Sub AddToAccountBox(accountBox As MSForms.ListBox, account As clsAccount)
    'Add the account to the listbox
    With accountBox
        Dim topOfList As Integer
        topOfList = .ListCount
        .AddItem
        .List(topOfList, 0) = account.NameOfAccount
        .List(topOfList, 1) = account.Number
        .List(topOfList, 2) = account.TypeOfAccount
        .List(topOfList, 3) = account.Owner.NameOfMember
        .List(topOfList, 4) = account.Owner.ContainingHousehold.NameOfHousehold
    End With
End Sub

Private Sub AddToBeneBox(beneBox As MSForms.ListBox, bene As clsBeneficiary)
    'Add the bene to the listbox
    With beneBox
        Dim topOfList As Integer
        topOfList = .ListCount
        .AddItem
        .List(topOfList, 0) = bene.NameOfBeneficiary
        .List(topOfList, 1) = bene.Level
        .List(topOfList, 2) = bene.Percent
        .List(topOfList, 3) = bene.account.NameOfAccount
        .List(topOfList, 4) = bene.account.Owner.NameOfMember
        .List(topOfList, 5) = bene.account.Owner.ContainingHousehold.NameOfHousehold
    End With
End Sub

Private Sub SortListBox(lbx As MSForms.ListBox)
    'Store the list in an array for sorting
    Dim listBoxList As Variant
    listBoxList = lbx.List
    
    'Bubble sort the array on the first value
    Dim i As Long, j As Long
    Dim tempStrings() As String
    ReDim tempStrings(1 To lbx.ColumnCount) As String
    For i = LBound(listBoxList, 1) To UBound(listBoxList, 1) - 1
        For j = i + 1 To UBound(listBoxList, 1)
            If listBoxList(i, 0) > listBoxList(j, 0) Then
                Dim listBoxColumn As Integer
                For listBoxColumn = 1 To lbx.ColumnCount
                    'Swap the first value
                    tempStrings(listBoxColumn) = listBoxList(i, listBoxColumn - 1)
                    listBoxList(i, listBoxColumn - 1) = listBoxList(j, listBoxColumn - 1)
                    listBoxList(j, listBoxColumn - 1) = tempStrings(listBoxColumn)
                Next listBoxColumn
            End If
        Next j
    Next i
    
    'Remove the contents of the listbox
    lbx.Clear
    
    'Repopulate with the sorted list
    lbx.List = listBoxList
End Sub
