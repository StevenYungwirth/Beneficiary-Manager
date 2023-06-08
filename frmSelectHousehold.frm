VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSelectHousehold 
   Caption         =   "Select a Household"
   ClientHeight    =   2592
   ClientLeft      =   48
   ClientTop       =   396
   ClientWidth     =   6732
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

Public Property Get SelectedHousehold() As clsHousehold
    Set SelectedHousehold = m_selectedHousehold
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
            'Add the household to the array
            m_households.Add households.Items(household).NameOfHousehold, households.Items(household)
        End If
    Next household
    
    'Sort the household list
    SortHouseholdList
    
    'Add each household to the combobox
    Dim hhold As Integer
    For hhold = 0 To m_households.count - 1
        cbxHousehold.AddItem m_households.Keys(hhold)
    Next hhold
End Sub

Private Sub SortHouseholdList()
    'Turn off screen updating
    Application.ScreenUpdating = False

    'Create a temporary worksheet
    Dim tempSheet As Worksheet
    Set tempSheet = ThisWorkbook.Sheets.Add
    
    'Put all of the keys from m_households into the worksheet
    Dim rw As Integer
    Dim dictKey As Variant
    For Each dictKey In m_households.Keys
        rw = rw + 1
        tempSheet.Cells(rw, 1).value = dictKey
    Next dictKey
    
    'Sort the worksheet
    tempSheet.UsedRange.Sort Key1:=tempSheet.Cells(1, 1), order1:=xlAscending, Header:=xlNo
    
    'Create a temporary dictionary and set the key-item pairs with the sorted list
    Dim tempDict As Dictionary: Set tempDict = New Dictionary
    For rw = 1 To tempSheet.UsedRange.Rows.count
        tempDict.Add tempSheet.Cells(rw, 1).Value2, m_households.Item(tempSheet.Cells(rw, 1).Value2)
    Next rw
    
    'Overwrite the original dictionary with the temporary one
    Set m_households = tempDict
    
    'Delete the hidden worksheet
    Application.DisplayAlerts = False
    tempSheet.Delete
    Application.DisplayAlerts = True

    'Turn screen updating back on
    Application.ScreenUpdating = True
End Sub
