VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmAddBene 
   Caption         =   "Add a Beneficiary"
   ClientHeight    =   4710
   ClientLeft      =   48
   ClientTop       =   396
   ClientWidth     =   6732
   OleObjectBlob   =   "frmAddBene.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmAddBene"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_beneAdded As Boolean
Private m_selectedHousehold As clsHousehold
Private m_selectedAccount As clsAccount
Private m_households As Dictionary
Private m_accounts As Dictionary

Public Property Get wasBeneAdded() As Boolean
    wasBeneAdded = m_beneAdded
End Property

Private Property Get SelectedHousehold() As clsHousehold
    Set SelectedHousehold = m_selectedHousehold
End Property

Private Property Get SelectedAccount() As clsAccount
    Set SelectedAccount = m_selectedAccount
End Property

Private Sub UserForm_Initialize()
    'Start the form in the middle of the screen with Excel
    Me.StartUpPosition = 0
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)

    'Initialize the dictionaries
    Set m_households = New Dictionary
    Set m_accounts = New Dictionary
End Sub

Private Sub cbxHousehold_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Drop down the combo box if it's clicked on
    cbxHousehold.DropDown
End Sub

Private Sub cbxaccount_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Drop down the combo box if it's clicked on
    cbxAccount.DropDown
End Sub

Private Sub rdoPrimary_KeyDown(ByVal keyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    IfEnterPressedCallAddClick keyCode
End Sub

Private Sub rdoContingent_KeyDown(ByVal keyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    IfEnterPressedCallAddClick keyCode
End Sub

Private Sub txtBeneName_KeyDown(ByVal keyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    IfEnterPressedCallAddClick keyCode
End Sub

Private Sub txtPercent_KeyDown(ByVal keyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    IfEnterPressedCallAddClick keyCode
End Sub

Private Sub chkPerStirpes_KeyDown(ByVal keyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    IfEnterPressedCallAddClick keyCode
End Sub

Private Sub UserForm_KeyDown(ByVal keyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    IfEnterPressedCallAddClick keyCode
End Sub

Private Sub cbxHousehold_Change()
    'Clear the account list
    cbxAccount.Clear
    Set m_accounts = New Dictionary
    
    If cbxHousehold.ListIndex <> -1 Then
        'Show the selected household's accounts in the account combo box
        Set m_selectedHousehold = m_households.Items(cbxHousehold.ListIndex)
        FillAccountList SelectedHousehold
        
        'Enable the account combobox
        cbxAccount.Enabled = True
    Else
        Set m_selectedHousehold = Nothing
        cbxAccount.Enabled = False
    End If
End Sub

Private Sub cbxAccount_Change()
    If cbxAccount.ListIndex <> -1 Then
        Set m_selectedAccount = m_accounts.Items(cbxAccount.ListIndex)
    Else
        Set m_selectedAccount = Nothing
    End If
    
    EnableAddButtons
End Sub

Private Sub txtBeneName_Change()
    EnableAddButtons
End Sub

Private Sub rdoPrimary_Click()
    EnableAddButtons
End Sub

Private Sub rdoContingent_Click()
    EnableAddButtons
End Sub

Private Sub spnPercent_SpinUp()
    'Don't raise the value above 100
    If txtPercent.value < 100 Then
        txtPercent.value = txtPercent.value + 1
    End If
End Sub

Private Sub spnPercent_SpinDown()
    'Don't lower the value below 0
    If txtPercent.value > 0 Then
        txtPercent.value = txtPercent.value - 1
    End If
End Sub

Private Sub txtPercent_Change()
    If IsNumeric(txtPercent.value) Then
        'Don't allow values above 100
        If txtPercent.value > 100 Then
            txtPercent.value = 100
        End If
    Else
        'Don't allow non-numeric characters
        Dim char As Integer
        For char = 1 To Len(txtPercent.value)
            If Not IsNumeric(Mid(txtPercent.value, char, 1)) Then
                txtPercent.value = Replace(txtPercent.value, Mid(txtPercent.value, char, 1), vbNullString)
            End If
        Next char
    End If
    
    'Enable the buttons if they can be enabled
    EnableAddButtons
End Sub

Private Sub btnAdd_Click()
    'Add the beneficiary to this worksheet, the master worksheet, and the XML
    If AddBeneToSheetAndXML Then MsgBox "Beneficiary added to account"
    
    'Hide the form
    Me.Hide
End Sub

Private Sub btnSaveAddAnother_Click()
    'Add the beneficiary to this worksheet and the XML
    AddBeneToSheetAndXML
    
    'Clear the form except for the combo boxes
    txtBeneName.value = vbNullString
    rdoPrimary.value = False
    rdoContingent.value = False
    txtPercent.value = 0
    chkPerStirpes.value = False
    
    'Put focus on the beneficiary name text box
    txtBeneName.SetFocus
End Sub

Private Sub btnCancel_Click()
    'Hide the form
    Me.Hide
End Sub

Private Function AddBeneToSheetAndXML() As Boolean
    'Add the beneficiary
    Dim beneToAdd As clsBeneficiary
    Set beneToAdd = GetBeneFromForm
    
    'Add the beneficiary to the XML file
    Dim addedBeneNode As IXMLDOMNode
    Set addedBeneNode = XMLWrite.AddBeneficiaryToNode(beneToAdd, sheetName:=ProjectGlobals.m_manualBeneListName)
    If Not addedBeneNode Is Nothing Then
        beneToAdd.id = addedBeneNode.SelectSingleNode("ID").Text
    
        'Add the beneficiary to this and the master worksheets
        FormProcedures.AddToSheet acctToAdd:=SelectedAccount, beneToAdd:=beneToAdd, action:="Add"
        
        'Format and save the XML
        XMLProcedures.FormatAndSaveXML
    
        'Save the workbook
        ThisWorkbook.Save
    Else
        AddBeneToSheetAndXML = False
        MsgBox "The beneficiary wasn't able to be added to the account"
    End If
End Function

Private Sub IfEnterPressedCallAddClick(keyCode As MSForms.ReturnInteger)
    'Call btnAdd's click method if the button is enabled and enter is pressed
    If keyCode = 13 And btnAdd.Enabled Then
        btnAdd_Click
    End If
End Sub

Public Sub FillHouseholdList(households As Dictionary)
    'Add the list of households to the combobox and the m_households array
    Dim household As Integer
    For household = 0 To households.count - 1
        Dim householdItem As clsHousehold
        Set householdItem = households.Items(household)
        If households.Items(household).Active And HasBeneChangeEligibleAccount(householdItem) Then
            'The household is active and has an account that's eligible for beneficiary changes; add the household to the dictionary of households
            m_households.Add households.Items(household).NameOfHousehold, households.Items(household)
        End If
    Next household
    
    'Sort the household list
    Set m_households = FormProcedures.SortHouseholdList(m_households)
    
    'Add each household to the combobox
    Dim hhold As Integer
    For hhold = 0 To m_households.count - 1
        cbxHousehold.AddItem m_households.Keys(hhold)
    Next hhold
End Sub

Private Sub FillAccountList(household As clsHousehold)
    If Not household Is Nothing Then
        'Add the household's accounts to the combo box
        Dim member As Integer
        For member = 0 To household.members.count - 1
            Dim account As Integer
            For account = 0 To household.members.Items(member).Accounts.count - 1
                'Only allow the beneficiary to be modified if it's from an active non-TD account
                Dim householdMember As clsMember
                Set householdMember = household.members.Items(member)
                Dim memberAccount As clsAccount
                Set memberAccount = householdMember.Accounts.Items(account)
                If IsAccountBeneChangeEligible(memberAccount) Then
                    With memberAccount
                        'The account is active; add it to the combobox
                        cbxAccount.AddItem .NameOfAccount
                        
                        'Add the account to the array
                        m_accounts.Add .NameOfAccount & .Number, memberAccount
                    End With
                End If
            Next account
        Next member
    End If
End Sub

Private Function HasBeneChangeEligibleAccount(household As clsHousehold) As Boolean
    'Check each account and each member of the household for an account whose beneficiaries can be changed
    Dim householdMemberItem As Integer
    Do While householdMemberItem < household.members.count And Not HasBeneChangeEligibleAccount
        Dim householdMember As clsMember
        Set householdMember = household.members.Items(householdMemberItem)
        Dim memberAccountItem As Integer
        Do While memberAccountItem < householdMember.Accounts.count And Not HasBeneChangeEligibleAccount
            Dim memberAccount As clsAccount
            Set memberAccount = householdMember.Accounts.Items(memberAccountItem)
            HasBeneChangeEligibleAccount = IsAccountBeneChangeEligible(memberAccount)
            memberAccountItem = memberAccountItem + 1
        Loop
        householdMemberItem = householdMemberItem + 1
    Loop
End Function

Private Function IsAccountBeneChangeEligible(account As clsAccount) As Boolean
    'A beneficiary can be modified if it's from an active non-TD account
    IsAccountBeneChangeEligible = account.Active And account.custodian <> ProjectGlobals.DefaultCustodian
End Function

Private Sub EnableAddButtons()
    'Only enable the buttons if everything needed for the new beneficiary is available
    If Not SelectedAccount Is Nothing _
    And txtBeneName.value <> vbNullString _
    And (rdoPrimary.value = True Or rdoContingent.value = True) _
    And txtPercent.value > 0 And txtPercent.value <= 100 Then
        btnAdd.Enabled = True
        btnSaveAddAnother.Enabled = True
    Else
        btnAdd.Enabled = False
        btnSaveAddAnother.Enabled = False
    End If
End Sub

Private Function GetBeneFromForm() As clsBeneficiary
    'Get the beneficiary name
    Dim beneName As String
    If chkPerStirpes.value Then
        beneName = Trim(txtBeneName.value) & " Per Stirpes"
    Else
        beneName = Trim(txtBeneName.value)
    End If
    
    'Get the beneficiary level
    Dim beneLevel As String
    If rdoPrimary.value Then
        beneLevel = "P"
    Else
        beneLevel = "C"
    End If
    
    'Get the beneficiary percent
    Dim benePercent As Double
    benePercent = txtPercent.value
    
    'Declare the beneficiary and add ID and account
    Dim tempBene As clsBeneficiary
    Set tempBene = ClassConstructor.NewBene(beneName, beneLevel, benePercent)
    With tempBene
        .id = GetNextBeneID
        .account = SelectedAccount
    End With
    
    'Return the beneficiary
    Set GetBeneFromForm = tempBene
End Function

Private Function GetNextBeneID() As Integer
    'Get the client list node
    Dim clientListNode As IXMLDOMElement
    Set clientListNode = ClientListFile.SelectSingleNode("Client_List")
    
    'Increment and return the max beneficiary ID attribute
    clientListNode.setAttribute "Max_Beneficiary_ID", clientListNode.getAttribute("Max_Beneficiary_ID") + 1
    GetNextBeneID = clientListNode.getAttribute("Max_Beneficiary_ID")
End Function
