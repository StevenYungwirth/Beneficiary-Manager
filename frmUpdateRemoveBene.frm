VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmUpdateRemoveBene 
   Caption         =   "Add a Beneficiary"
   ClientHeight    =   4710
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6735
   OleObjectBlob   =   "frmUpdateRemoveBene.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmUpdateRemoveBene"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_households As Dictionary
Private m_accounts As Dictionary

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
        Set SelectedBeneficiary = SelectedAccount.Benes(lbxBeneficiaries.ListIndex + 1)
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
        Dim selectedHousehold As clsHousehold
        Set selectedHousehold = m_households.Items(cbxHousehold.ListIndex)
        
        'Show the selected household's accounts in the account combo box
        FillAccountList selectedHousehold
        
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
            FormProcedures.AddToSheet SelectedAccount, frmUpdateBene.updatedBene, "Updated"
        
            'Unload the form
            Unload frmUpdateBene
            
            'Hide this form
            Me.Hide
        Else
            'Unload the form
            Unload frmUpdateBene
        End If
    End If
End Sub

Private Sub btnRemove_Click()
    If Not SelectedBeneficiary Is Nothing Then
        'Get confirmation
        If MsgBox("Are you sure you want to remove " & SelectedBeneficiary.NameOfBeneficiary & " from the beneficiaries?", vbYesNo) = vbYes Then
            'Remove the selected beneficiary from the XML file
            If XMLReadWrite.UpdateRemoveBene(2, SelectedBeneficiary) Then
                'Add bene to the sheet
                FormProcedures.AddToSheet SelectedAccount, SelectedBeneficiary, "Deleted"
                
                'Show confirmation and remove the beneficiary from the listbox
                lbxBeneficiaries.RemoveItem lbxBeneficiaries.ListIndex
                MsgBox "Beneficiary has been removed."
                
                'Hide this form
                Me.Hide
            Else
                MsgBox "Beneficiary wasn't able to be removed."
            End If
        End If
    End If
End Sub

Private Sub btnCancel_Click()
    'Return nothing and hide the form
    Me.Hide
End Sub

Public Sub FillHouseholdList(households As Dictionary)
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

Private Sub FillAccountList(household As clsHousehold)
    'Add the household's accounts to the combo box
    Dim member As Variant
    For Each member In household.Members.Items
        Dim account As Variant
        For Each account In member.accounts.Items
            If account.Active Then
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
        For Each bene In account.Benes
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
