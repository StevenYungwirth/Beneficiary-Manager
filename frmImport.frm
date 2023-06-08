VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmImport 
   Caption         =   "Select Files to Import"
   ClientHeight    =   3756
   ClientLeft      =   48
   ClientTop       =   396
   ClientWidth     =   4092
   OleObjectBlob   =   "frmImport.frx":0000
End
Attribute VB_Name = "frmImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_tdaBeneFile As String
Private m_msAccountsFile As String
Private m_rtAccountsFile As String
Private m_rtContactsFile As String
Private m_msHouseholdExportFile As String

Public Property Get TDABeneFile() As String
    TDABeneFile = m_tdaBeneFile
End Property

Public Property Get MSAccountsFile() As String
    MSAccountsFile = m_msAccountsFile
End Property

Public Property Get RTAccountsFile() As String
    RTAccountsFile = m_rtAccountsFile
End Property

Public Property Get RTContactsFile() As String
    RTContactsFile = m_rtContactsFile
End Property

Public Property Get MSHouseholdExportFile() As String
    MSHouseholdExportFile = m_msHouseholdExportFile
End Property

Private Sub UserForm_Initialize()
    'Start the form in the middle of the screen with Excel
    Me.StartUpPosition = 0
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)
End Sub

Private Sub chkTDABene_Click()
    chkClick m_tdaBeneFile, chkTDABene
End Sub

Private Sub chkMSAccounts_Click()
    chkClick m_msAccountsFile, chkMSAccounts
End Sub

Private Sub chkRTAccounts_Click()
    chkClick m_rtAccountsFile, chkRTAccounts
End Sub

Private Sub chkRTContacts_Click()
    chkClick m_rtContactsFile, chkRTContacts
End Sub

Private Sub chkMSHouseholdExport_Click()
    chkClick m_msHouseholdExportFile, chkMSHouseholdExport
End Sub

Private Sub chkClick(ByRef filePath As String, chk As Object)
    If chk.value Then
        filePath = SelectFile
        
        'Uncheck the box if the user didn't select a file
        If Len(filePath) = 0 Then
            chk.value = False
        End If
    Else
        filePath = vbNullString
    End If
End Sub

Private Sub btnSelect_Click()
    Me.Hide
End Sub

Private Sub btnCancel_Click()
    Me.Hide
    End
End Sub

Private Function SelectFile() As String
    'Show a file dialog
    With Application.FileDialog(msoFileDialogFilePicker)
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "Excel Files", "*.csv; *.xls*", 1
        If InStr(ThisWorkbook.name, "working copy") > 0 Then
            .InitialFileName = "Z:\FPIS - Operations\Beneficiary Project\Archive\"
        Else
            .InitialFileName = "Z:\"
        End If
        .Show
        
        'Get the selected file
        Dim fileSelected As Variant
        For Each fileSelected In .SelectedItems
            SelectFile = .SelectedItems.Item(1)
        Next fileSelected
    End With
End Function
