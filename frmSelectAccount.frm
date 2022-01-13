VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSelectAccount 
   Caption         =   "Select an Account"
   ClientHeight    =   2535
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6735
   OleObjectBlob   =   "frmSelectAccount.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSelectAccount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_householdNodes As IXMLDOMNodeList
Private m_accountNodes As IXMLDOMNodeList
Private m_selectedAccount As IXMLDOMNode

Public Property Get SelectedAccount() As IXMLDOMNode
    Set SelectedAccount = m_selectedAccount
End Property

Public Property Get SelectedHouseholdIndex() As Integer
    SelectedHouseholdIndex = cbxHousehold.ListIndex
End Property

Public Property Get SelectedAccountIndex() As Integer
    SelectedAccountIndex = cbxAccount.ListIndex
End Property

Private Sub UserForm_Initialize()
    'Start the form in the middle of the screen with Excel
    Me.StartUpPosition = 0
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)
End Sub

Private Sub cbxHousehold_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Drop down the combo box if it's clicked on
    cbxHousehold.DropDown
End Sub

Private Sub cbxaccount_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Drop down the combo box if it's clicked on
    cbxAccount.DropDown
End Sub

Private Sub UserForm_KeyDown(ByVal keyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    IfEnterPressedCallAddClick keyCode
End Sub

Private Sub cbxHousehold_Change()
    If cbxHousehold.ListIndex <> -1 Then
        'A household was selected; clear the account list
        cbxAccount.Clear
        
        'Show the selected household's accounts in the account combo box
        FillAccountList
        
        'Enable the account combobox
        cbxAccount.Enabled = True
    End If
End Sub

Private Sub cbxAccount_Change()
    'Enable the select button if an account is selected
    If cbxAccount.ListIndex <> -1 Then
        btnSelect.Enabled = True
    Else
        btnSelect.Enabled = False
    End If
End Sub

Private Sub btnSelect_Click()
    'Return the selected account
    If cbxAccount.ListIndex <> -1 Then
        Set m_selectedAccount = m_accountNodes(cbxAccount.ListIndex)
    End If
    
    'Hide the form
    Me.Hide
End Sub

Private Sub btnCancel_Click()
    'Return nothing
    Set m_selectedAccount = Nothing
    
    'Hide the form
    Me.Hide
End Sub

Public Sub FillHouseholdList()
    'Load the XML file
    Dim xmlFile As DOMDocument60
    Set xmlFile = XMLReadWrite.LoadClientList
    
    'Get the household nodes
    Set m_householdNodes = xmlFile.SelectNodes("//Household[@Active='True']")
    
    'Add the list of households to the combobox
    Dim household As Integer
    For household = 0 To m_householdNodes.Length - 1
        'Get the household from the node
        Dim householdFromNode As clsHousehold
        Set householdFromNode = XMLReadWrite.ReadHouseholdFromNode(m_householdNodes(household))
        
        'Add the household to the combox if it's active
        If householdFromNode.Active Then
            cbxHousehold.AddItem householdFromNode.NameOfHousehold
        End If
    Next household
End Sub

Private Sub FillAccountList()
    'Don't do anything if no household is selected
    If cbxHousehold.ListIndex = -1 Then
        Exit Sub
    End If
    
    'Get the selected household
    Dim selectedHousehold As IXMLDOMNode
    Set selectedHousehold = m_householdNodes(cbxHousehold.ListIndex)
    
    'Get the account nodes in the selected household
    Set m_accountNodes = selectedHousehold.SelectNodes(".//Account")
    
    'Add the household's accounts to the combo box
    Dim account As Integer
    For account = 0 To m_accountNodes.Length - 1
        'Get the account from the node
        Dim accountFromNode As clsAccount
        Set accountFromNode = XMLReadWrite.ReadAccountFromNode(m_accountNodes(account))
        
        'Add the account to the combobox
        cbxAccount.AddItem accountFromNode.NameOfAccount
    Next account
End Sub

Private Sub IfEnterPressedCallAddClick(keyCode As MSForms.ReturnInteger)
    'Call btnAdd's click method if the button is enabled and enter is pressed
    If keyCode = 13 And btnSelect.Enabled Then
        btnSelect_Click
    End If
End Sub
