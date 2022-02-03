VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmUpdateBene 
   Caption         =   "Update Beneficiary"
   ClientHeight    =   3030
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9390
   OleObjectBlob   =   "frmUpdateBene.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmUpdateBene"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_currentBene As clsBeneficiary
Private m_beneUpdated As Boolean

Public Property Get updatedBene() As clsBeneficiary
    Set updatedBene = GetBeneFromForm
End Property

Public Property Get BeneUpdated() As Boolean
    BeneUpdated = m_beneUpdated
End Property

Private Sub UserForm_Initialize()
    'Start the form in the middle of the screen with Excel
    Me.StartUpPosition = 0
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)
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
End Sub

Private Sub btnUpdate_Click()
    'Update the beneficiary in the XML file
    m_beneUpdated = XMLReadWrite.UpdateRemoveBene(1, m_currentBene, GetBeneFromForm)
    If m_beneUpdated Then
        'Beneficiary has been updated; show confirmation
        MsgBox "Beneficiary has been updated."
    End If
        
    'Hide the form
    Me.Hide
End Sub

Private Sub btnCancel_Click()
    'Hide the form
    Me.Hide
End Sub

Public Sub LoadBeneficiary(bene As clsBeneficiary)
    'Set the beneficiary
    Set m_currentBene = bene

    'Set the current beneficiary values
    lblCurrentName.Caption = "Name: " & m_currentBene.NameOfBeneficiary
    lblCurrentLevel.Caption = "Level: " & m_currentBene.Level
    lblCurrentPercent.Caption = "Percent: " & m_currentBene.Percent
    
    'Set the new beneficiary values to initially equal the old ones
    txtName.Text = m_currentBene.NameOfBeneficiary
    If m_currentBene.Level = "Primary" Or m_currentBene.Level = "P" Then
        rdoPrimary.value = True
    Else
        rdoContingent.value = True
    End If
    txtPercent.Text = m_currentBene.Percent
End Sub

Private Function GetBeneFromForm() As clsBeneficiary
    'Get the beneficiary name
    Dim beneName As String
    beneName = txtName.value
    
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
    
    'Declare the beneficiary to return
    Set GetBeneFromForm = ClassConstructor.NewBene(beneName, beneLevel, benePercent)
    
    'Add the account information
    GetBeneFromForm.account.NameOfAccount = m_currentBene.account.NameOfAccount
    GetBeneFromForm.account.Number = m_currentBene.account.Number
    GetBeneFromForm.account.ID = m_currentBene.account.ID
    
    'Add the beneficiary ID
    GetBeneFromForm.ID = m_currentBene.ID
End Function
