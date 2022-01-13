VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmAddBene 
   Caption         =   "Add a Beneficiary"
   ClientHeight    =   4710
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6735
   OleObjectBlob   =   "frmAddBene.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmAddBene"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_households As Dictionary
Private m_accounts As Dictionary

Private Property Get selectedHousehold() As clsHousehold
    If cbxHousehold.ListIndex <> -1 Then
        Set selectedHousehold = m_households.Items(cbxHousehold.ListIndex)
    End If
End Property

Private Property Get SelectedAccount() As clsAccount
    If cbxAccount.ListIndex <> -1 Then
        Set SelectedAccount = m_accounts.Items(cbxAccount.ListIndex)
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
    If cbxHousehold.ListIndex <> -1 Then
        'A household was selected; clear the account list
        cbxAccount.Clear
        Set m_accounts = New Dictionary
        
        'Show the selected household's accounts in the account combo box
        FillAccountList selectedHousehold
        
        'Enable the account combobox
        cbxAccount.Enabled = True
    End If
End Sub

Private Sub cbxAccount_Change()
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
    AddBeneToSheetAndXML
    
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

Private Sub AddBeneToSheetAndXML()
    'Add the beneficiary
    Dim beneToAdd As clsBeneficiary
    Set beneToAdd = GetBeneFromForm
    
    'Add the beneficiary to the XML file
    beneToAdd.ID = XMLReadWrite.AddBenesToXML(beneToAdd).getAttribute("ID")
    
    'Add the beneficiary to this and the master worksheets
    FormProcedures.AddToSheet acctToAdd:=SelectedAccount, beneToAdd:=beneToAdd, action:="Added"
End Sub

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

Private Sub EnableAddButtons()
    'Only enable the buttons if everything needed for the new beneficiary is available
    If cbxAccount.ListIndex <> -1 _
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
        .ID = GetNextBeneID(XMLReadWrite.LoadClientList)
        .account = SelectedAccount
    End With
    
    'Return the beneficiary
    Set GetBeneFromForm = tempBene
End Function

Private Function GetNextBeneID(xmlFile As DOMDocument60) As Integer
    'Get the client list node
    Dim clientListNode As IXMLDOMElement
    Set clientListNode = xmlFile.SelectSingleNode("Client_List")
    
    'Increment and return the max beneficiary ID attribute
    GetNextBeneID = clientListNode.getAttribute("Max_Beneficiary_ID") + 1
End Function
