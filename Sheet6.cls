VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub btnGenerateReports_Click()
    'Generate the beneficiary report for each household
    GenerateReports.BeneficiaryReport XMLReadWrite.ReadHouseholdsFromFile
End Sub

Private Sub btnGenerateSingle_Click()
    'Get the list of households
    Dim households As Dictionary
    Set households = XMLReadWrite.ReadHouseholdsFromFile
    
    'Add the households to the form
    Load frmSelectHousehold
    frmSelectHousehold.FillHouseholdList households
    
    'From the form, get the household to run a report on
    frmSelectHousehold.Show
    Dim selectedHousehold As clsHousehold
    Set selectedHousehold = frmSelectHousehold.selectedHousehold
    
    If Not selectedHousehold Is Nothing Then
        'Turn off screen updating
        Application.ScreenUpdating = False
        
        'Show the household's beneficiaries
        Dim singleReportBook As Workbook
        Set singleReportBook = GenerateReports.GenerateReportFromHousehold(selectedHousehold)
        
        'Turn on screen updating
        Application.ScreenUpdating = True
    End If
    
    'Unload the form
    Unload frmSelectHousehold
End Sub

Private Sub btnAddBene_Click()
    'Get the list of households
    Dim households As Dictionary
    Set households = XMLReadWrite.ReadHouseholdsFromFile
    
    'Add the households to the form
    Load frmAddBene
    frmAddBene.FillHouseholdList households
    
    'Show the form
    frmAddBene.Show
    
    'Unload the form
    Unload frmAddBene
    
    'Save the workbook
    ThisWorkbook.Save
End Sub

Private Sub btnUpdateRemove_Click()
    'Get the list of households
    Dim households As Dictionary
    Set households = XMLReadWrite.ReadHouseholdsFromFile
    
    'Add the households to the form
    Load frmUpdateRemoveBene
    frmUpdateRemoveBene.FillHouseholdList households
    
    'Show the form
    frmUpdateRemoveBene.Show
    
    'Unload the form
    Unload frmUpdateRemoveBene
    
    'Save the workbook
    ThisWorkbook.Save
End Sub

Private Sub btnModifyOverride_Click()
    'Show the form
    Dim overrideForm As frmOverride
    Set overrideForm = New frmOverride
    overrideForm.Show
    
    'Unload the form
    Unload overrideForm
End Sub

Private Sub btnImport_Click()
    'Ask for a password to prevent accidental imports
    Dim password As String
    password = "importstuff"
    If InputBox("Please enter the password:", Title:="Enter Password") = password Then
        'Import new data from files
        ImportDataSheets.ImportData
    Else
        MsgBox "Password incorrect"
    End If
End Sub

'Private Sub btnCreateClientList_Click()
'    'Regenerate the Client List XML file
'    XMLCreateList.CreateHouseholdsXMLFile XMLReadWrite.ReadHouseholdsFromFile
'End Sub
