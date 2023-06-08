VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmOverride 
   Caption         =   "Modify Override"
   ClientHeight    =   3480
   ClientLeft      =   48
   ClientTop       =   396
   ClientWidth     =   5784
   OleObjectBlob   =   "frmOverride.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmOverride"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_overrides As IXMLDOMNodeList

Private Property Get selectedNode() As IXMLDOMElement
    If lbxOverrides.ListIndex <> -1 Then
        Set selectedNode = m_overrides(lbxOverrides.ListIndex)
    Else
        Set selectedNode = Nothing
    End If
End Property

Private Property Get XMLClientList() As DOMDocument60
    Set XMLClientList = ProjectGlobals.ClientListFile
End Property

Private Sub UserForm_Initialize()
    'Start the form in the middle of the screen with Excel
    Me.StartUpPosition = 0
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)
    
    'Hide the controls for adding/updating/removing overrides
    lbxOverrides.Visible = False
    fraAdd.Visible = False
    btnOK.Enabled = False
    
    'Load the overrides
    FillListBox
                
    'Set the listbox column widths
    lbxOverrides.ColumnWidths = "100,30,60"
End Sub

Private Sub rdoAdd_Click()
    'Show the frame with the override options and hide the listbox
    fraAdd.Visible = True
    lbxOverrides.Visible = False
    
    'Reset the radios
    rdoActive.value = False
    rdoTag.value = False
    
    'Hide the controls for the override values
    fraActive.Visible = False
    txtTag.Visible = False
    
    'Change the ok button to say "Add"
    UpdateOKBtn
End Sub

Private Sub rdoUpdate_Click()
    ShowUpdateRemoveControls
End Sub

Private Sub rdoRemove_Click()
    ShowUpdateRemoveControls
End Sub

Private Sub rdoActive_Click()
    'Show the controls for the account active override
    fraActive.Visible = True
    
    'Hide the control for the account tag override
    txtTag.Visible = False
    
    'Reset the radios
    rdoTrue.value = False
    rdoFalse.value = False
    
    'Disable the OK button
    btnOK.Enabled = False
End Sub

Private Sub rdoTag_Click()
    'Show the controls for the account tag override
    txtTag.Visible = True
    
    'Hide the control for the account active override
    fraActive.Visible = False
    
    'Reset the text box
    txtTag.value = vbNullString
    
    'Disable the OK button
    btnOK.Enabled = False
End Sub

Private Sub rdoTrue_Click()
    'Enable the OK button
    btnOK.Enabled = True
End Sub

Private Sub rdoFalse_Click()
    'Enable the OK button
    btnOK.Enabled = True
End Sub

Private Sub txtTag_Change()
    'Enable the OK button if the text box isn't empty
    btnOK.Enabled = (txtTag.value <> vbNullString)
End Sub

Private Sub lbxOverrides_Click()
    'Enable the OK button if an item is selected
    btnOK.Enabled = (lbxOverrides.ListIndex <> -1)
End Sub

Private Sub btnOK_Click()
    'Track if changes were made to the XML
    Dim changesMade As Boolean

    'Run the proper method
    If rdoAdd.value Then
        changesMade = AddOverride
    ElseIf rdoUpdate.value Then
        changesMade = UpdateOverride
    ElseIf rdoRemove.value Then
        changesMade = RemoveOverride
    End If
    
    If changesMade Then
        'Save the XML file
        XMLClientList.Save ProjectGlobals.ClientListFilePath
    
        'Hide the form
        Me.Hide
    End If
End Sub

Private Sub btnCancel_Click()
    'Hide the form
    Me.Hide
End Sub

Private Sub FillListBox()
    'Get the override nodes
    Set m_overrides = XMLClientList.SelectNodes("//Override")
        
    'For each node, show the account name, override type, override value in the listbox
    Dim Node As Variant
    For Each Node In m_overrides
        'Set the current node
        Dim overrideNode As IXMLDOMElement
        Set overrideNode = Node
        
        'Get the parent node (Account)
        Dim overrideAccount As IXMLDOMElement
        Set overrideAccount = overrideNode.parentNode
        Dim accountName As String
        accountName = GetNodeAttribute("Name", overrideAccount)
        
        'Override nodes are set up to only have one attribute. Add it to the list if it has one
        If overrideNode.Attributes.Length > 0 Then
            With lbxOverrides
                'Get the number of items already in the list box
                Dim overrideCount As Integer
                overrideCount = .ListCount
                
                'Add the override to the list
                .AddItem
                .List(overrideCount, 0) = accountName
                .List(overrideCount, 1) = overrideNode.Attributes(0).BaseName
                .List(overrideCount, 2) = overrideNode.Attributes(0).Text
            End With
        End If
    Next Node
End Sub

Private Function GetNodeAttribute(attributeName As String, Node As IXMLDOMElement) As Variant
    If Not IsNull(Node.getAttribute(attributeName)) Then
        GetNodeAttribute = Node.getAttribute(attributeName)
    End If
End Function

Private Sub ShowUpdateRemoveControls()
    'Hide the controls for adding overrides and show the listbox
    fraAdd.Visible = False
    lbxOverrides.Visible = True
    
    'Reset the listbox
    lbxOverrides.ListIndex = -1
    
    'Change the ok button to say "Update"
    UpdateOKBtn
    
    'Disable the OK button
    btnOK.Enabled = False
End Sub

Private Sub UpdateOKBtn()
    'Update the OK button's caption to an appropriate value, depending on which radio is selected
    If rdoAdd.value Then
        btnOK.Caption = "Select Account"
    ElseIf rdoUpdate.value Then
        btnOK.Caption = "Update"
    ElseIf rdoRemove.value Then
        btnOK.Caption = "Remove"
    Else
        btnOK.Caption = "OK"
    End If
End Sub

Private Function AddOverride() As Boolean
    'Show a form with all accounts
    Dim frmSelect As frmSelectAccount
    Set frmSelect = New frmSelectAccount
    frmSelect.FillHouseholdList
    frmSelect.Show
    
    'Get the list household/account list indices from the form instead of the node
    Dim householdIndex As Integer, accountIndex As Integer
    householdIndex = frmSelect.SelectedHouseholdIndex
    accountIndex = frmSelect.SelectedAccountIndex

    'Get the list of households
    Dim households As IXMLDOMNodeList
    Set households = XMLClientList.SelectNodes("//Household[@Active='True']")

    'Get the household node
    Dim householdNode As IXMLDOMNode
    Set householdNode = households(householdIndex)

    'Get the household's accounts
    Dim accountNodes As IXMLDOMNodeList
    Set accountNodes = householdNode.SelectNodes(".//Account")

    'Get the selected account node
    Dim accountNode As IXMLDOMElement
    Set accountNode = accountNodes(accountIndex)

    If Not accountNode Is Nothing Then
        'The account node was found. Initialize the override node
        Dim accountOverride As IXMLDOMElement
        
        'Check if there's already an override node
        If accountNode.SelectSingleNode("Override") Is Nothing Then
            'The account doesn't already have an override, create the node and add it
            Set accountOverride = XMLClientList.createNode(1, "Override", "")
            accountOverride.setAttribute GetOverrideType, GetOverrideValue
            
            'Add the override node to the account node
            accountNode.appendChild accountOverride
            AddOverride = True
        Else
            'The account already has an override. See if the one being added contradicts the one already there
            Set accountOverride = accountNode.SelectSingleNode("Override")
            If GetOverrideType = "Active" And Not IsNull(accountOverride.getAttribute("Active")) Then
                'The existing override is for the Active status
                If CBool(accountOverride.getAttribute("Active")) = Not CBool(GetOverrideValue) Then
                    'The overrides contradict, so the one being added is how the account should be. Remove the active override entirely
                    accountOverride.removeAttribute "Active"
                Else
                    'The overrides are the same, so don't add another
                End If
            ElseIf GetOverrideType = "Tag" And Not IsNull(accountOverride.getAttribute("Tag")) Then
                'The existing override is for the Tag. Check if the override being added equals the one already there
                If accountOverride.getAttribute("Tag") <> GetOverrideValue Then
                    'The existing tag is different than the one being added. Add the override
                    accountOverride.setAttribute GetOverrideType, GetOverrideValue
                    AddOverride = True
                Else
                    'The overrides are the same, so don't add another
                End If
            Else
                'The existing override doesn't have any attributes. Add the override
                accountOverride.setAttribute GetOverrideType, GetOverrideValue
                AddOverride = True
            End If
        End If
    
        'Show confirmation if the override was added
        If AddOverride Then
            MsgBox GetOverrideType & ": " & GetOverrideValue & " override added to " & accountNode.Attributes(1).Text
        End If
    Else
        'The account node wasn't found. Throw an error
        MsgBox "Account not found in XML file; override not added."
        AddOverride = False
    End If
    
    'Unload the form
    Unload frmSelect
End Function

Private Function GetOverrideType() As String
    If rdoActive.value Then
        GetOverrideType = "Active"
    ElseIf rdoTag.value Then
        GetOverrideType = "Tag"
    End If
End Function

Private Function GetOverrideValue() As String
    If rdoTrue.value Then
        GetOverrideValue = "True"
    ElseIf rdoFalse.value Then
        GetOverrideValue = "False"
    ElseIf txtTag.value <> vbNullString Then
        GetOverrideValue = txtTag.value
    End If
End Function

Private Function UpdateOverride() As Boolean
    'Show a form for new override values
    Dim frmUpdate As frmOverrideUpdate
    Set frmUpdate = New frmOverrideUpdate
    frmUpdate.LoadOverride selectedNode
    frmUpdate.Show
    
    'Check if the override value changed
    If selectedNode.getAttribute(frmUpdate.OverrideType) <> frmUpdate.OverrideValue Then
        'Update the override node
        selectedNode.setAttribute frmUpdate.OverrideType, frmUpdate.OverrideValue
        UpdateOverride = True
    End If
    
    'Unload the form
    Unload frmUpdate
    
    'Show confirmation
    MsgBox "Override has been updated."
End Function

Private Function RemoveOverride() As Boolean
    'Remove the node from its parent
    selectedNode.parentNode.RemoveChild selectedNode
    
    'Show confirmation
    MsgBox "Override has been removed."
    
    'Return true
    RemoveOverride = True
End Function
