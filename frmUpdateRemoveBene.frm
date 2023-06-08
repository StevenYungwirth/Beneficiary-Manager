VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmUpdateRemoveBene 
   Caption         =   "Add a Beneficiary"
   ClientHeight    =   4704
   ClientLeft      =   48
   ClientTop       =   396
   ClientWidth     =   6768
   OleObjectBlob   =   "frmUpdateRemoveBene.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmUpdateRemoveBene"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_beneModified As Boolean
Private m_households As Dictionary
Private m_accounts As Dictionary

Public Property Get WasBeneModified() As Boolean
    WasBeneModified = m_beneModified
End Property

Private Property Get manualSheet() As Worksheet
    Set manualSheet = ThisWorkbook.Sheets("Manual Beneficiaries")
End Property

Private Property Get SelectedAccount() As clsAccount
    'Set the account selected in the listbox
    If cbxAccount.ListIndex <> -1 Then
        Set SelectedAccount = m_accounts.Items(cbxAccount.ListIndex)
    End If
End Property

Private Property Get SelectedBeneficiary() As clsBeneficiary
    If lbxBeneficiaries.ListIndex <> -1 Then
        Set SelectedBeneficiary = SelectedAccount.Benes.Items((lbxBeneficiaries.ListIndex))
    End If
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

Private Sub cbxHousehold_Change()
    If cbxHousehold.ListIndex <> -1 Then
        'A household was selected; clear the account list
        cbxAccount.Clear
        Set m_accounts = New Dictionary
        
        'Update the selected household
        Dim SelectedHousehold As clsHousehold
        Set SelectedHousehold = m_households.Items(cbxHousehold.ListIndex)
        
        'Show the selected household's accounts in the account combo box
        FillAccountList SelectedHousehold
        
        'Enable the account combobox
        cbxAccount.Enabled = True
    End If
End Sub

Private Sub cbxAccount_Change()
    'Clear the listbox
    lbxBeneficiaries.Clear
    
    'Fill the listbox
    If Not SelectedAccount Is Nothing Then
        FillBeneficiaryList SelectedAccount
    End If
End Sub

Private Sub btnUpdate_Click()
    If Not SelectedBeneficiary Is Nothing Then
        'Load the update beneficiary form
        Load frmUpdateBene
        frmUpdateBene.LoadBeneficiary SelectedBeneficiary
        frmUpdateBene.Show
        
        If frmUpdateBene.BeneUpdated Then
            'Put the updated beneficiary onto the sheet
            FormProcedures.AddToSheet SelectedAccount, frmUpdateBene.updatedBene, "Update"
        
            'Unload the form
            Unload frmUpdateBene
            
            'Hide this form
            m_beneModified = True
            Me.Hide
        Else
            'Unload the form
            m_beneModified = False
            Unload frmUpdateBene
        End If
    End If
End Sub

Private Sub btnRemove_Click()
    If Not SelectedBeneficiary Is Nothing Then
        'Get confirmation
        If MsgBox("Are you sure you want to remove " & SelectedBeneficiary.NameOfBeneficiary & " from the beneficiaries?", vbYesNo) = vbYes Then
            'Remove the selected beneficiary from the XML file
            If XMLWrite.UpdateRemoveBene(2, SelectedBeneficiary) Then
                'Add bene to the sheet
                FormProcedures.AddToSheet SelectedAccount, SelectedBeneficiary, "Delete"
                
                'Show confirmation and remove the beneficiary from the listbox
                lbxBeneficiaries.RemoveItem lbxBeneficiaries.ListIndex
                MsgBox "Beneficiary has been removed."
                
                'Hide this form
                m_beneModified = True
                Me.Hide
            Else
                m_beneModified = False
                MsgBox "Beneficiary wasn't able to be removed."
            End If
        End If
    End If
End Sub

Private Sub btnCancel_Click()
    'Return nothing and hide the form
    Me.Hide
End Sub

Public Sub FillHouseholdListasdf(households As Dictionary)
    'Add the list of households to the combobox and the m_households array
    Dim household As Integer
    For household = 0 To households.count - 1
        If households.Items(household).Active Then
            'The household is active; add them to the combobox
            cbxHousehold.AddItem households.Items(household).NameOfHousehold
            
            'Add the household to the array
            m_households.Add households.Items(household).NameOfHousehold, households.Items(household)
        End If
    Next household
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
    'Add the household's accounts to the combo box
    Dim member As Variant
    For Each member In household.members.Items
        Dim account As Variant
        For Each account In member.Accounts.Items
            'Only allow the beneficiary to be modified if it's from an active non-TD account
            If account.Active And account.custodian <> ProjectGlobals.DefaultCustodian Then
                'The account is active; add it to the combobox
                cbxAccount.AddItem account.NameOfAccount
                
                'Add the account to the array
                m_accounts.Add account.NameOfAccount & account.Number, account
            End If
        Next account
    Next member
End Sub

Private Sub FillBeneficiaryList(account As clsAccount)
    With lbxBeneficiaries
        'Add the account's beneficiaries to the listbox
        Dim bene As Variant
        For Each bene In account.Benes.Items
            'Get the number of items already in the list box
            Dim beneCount As Integer
            beneCount = .ListCount
            
            .AddItem
            .List(beneCount, 0) = bene.NameOfBeneficiary
            .List(beneCount, 1) = bene.Level
            .List(beneCount, 2) = bene.Percent
        Next bene
        
        .ColumnWidths = "180,50,30"
    End With
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
