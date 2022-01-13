VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSelectHousehold 
   Caption         =   "Select a Household"
   ClientHeight    =   2295
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6735
   OleObjectBlob   =   "frmSelectHousehold.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSelectHousehold"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_households As Dictionary
Private m_selectedHousehold As clsHousehold

Public Property Get selectedHousehold() As clsHousehold
    Set selectedHousehold = m_selectedHousehold
End Property

Private Sub UserForm_Initialize()
    'Start the form in the middle of the screen Excel is on
    Me.StartUpPosition = 0
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)
    
    'Initialize the household dictionary
    Set m_households = New Dictionary
End Sub

Private Sub cbxHousehold_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Drop down the combo box when it's clicked on
    cbxHousehold.DropDown
End Sub

Private Sub btnSelectHousehold_Click()
    'Return the selected household and hide the form
    Set m_selectedHousehold = m_households.Items(cbxHousehold.ListIndex)
    Me.Hide
End Sub

Private Sub btnCancel_Click()
    'Return nothing and hide the form
    Set m_selectedHousehold = Nothing
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
