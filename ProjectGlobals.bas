Attribute VB_Name = "ProjectGlobals"
Option Explicit
Private m_clientListFile As DOMDocument60
Private m_importTime As Date
Public Const manualSheetPassword As String = "unlockmanualbenes"
Public Const DefaultCustodian As String = "TD Ameritrade Institutional"
Public Const ArchiveFolder As String = "Z:\FPIS - Operations\Beneficiary Project\Archive\Households\"
Public Const ClientListFolder As String = "Z:\FPIS - Operations\Beneficiary Project\"
Public Const ClientListFilePath As String = "Z:\FPIS - Operations\Beneficiary Project\Assets\Households.xml"
'Public Const ClientListFilePath As String = "Z:\FPIS - Operations\Beneficiary Project\Assets\Sample Households.xml"
Public Const SampleClientListFilePath As String = "Z:\FPIS - Operations\Beneficiary Project\Assets\Sample Households.xml"
Public Const m_msExportName As String = "MS_Export"
Public Const m_msAccountName As String = "MS_Accounts"
Public Const m_rtAccountName As String = "RT_Accounts"
Public Const m_rtContactName As String = "RT_Contacts"
Public Const m_beneListName As String = "Bene_List"
Public Const m_manualBeneListName As String = "Manual_Sheet"
Public Const m_emptyMemberName As String = "EmptyMember"
Private Const associatedFileLocation As String = "Z:\FPIS - Operations\Beneficiary Project\Assets\associated accounts.txt"

Public Property Get ImportTime() As Date
    If m_importTime = CDate(0) Then
        m_importTime = Now
    End If
    
    ImportTime = m_importTime
End Property

Public Property Get ClientListFile(Optional isSample As Boolean) As DOMDocument60
    If m_clientListFile Is Nothing Then
        Set m_clientListFile = LoadClientList(isSample)
        m_clientListFile.preserveWhiteSpace = True
    End If
    
    Set ClientListFile = m_clientListFile
End Property

Public Sub ResetImportTime()
    m_importTime = CDate(0)
End Sub

Public Sub CloseClientFile()
    Set m_clientListFile = Nothing
End Sub

Public Property Get HouseholdNodeProperties(givenHousehold As clsHousehold) As Dictionary
    'Set properties for the household
    Dim nodeProperties As Dictionary
    Set nodeProperties = New Dictionary
    With givenHousehold
        nodeProperties.Add "Morningstar_ID", .morningstarID
        nodeProperties.Add "Name", .NameOfHousehold
        nodeProperties.Add "Active", CStr(.Active)
    End With
    
    'Return the properties
    Set HouseholdNodeProperties = nodeProperties
End Property

Public Property Get MemberNodeProperties(givenMember As clsMember) As Dictionary
    'Set properties for the Member
    Dim nodeProperties As Dictionary
    Set nodeProperties = New Dictionary
    With givenMember
        nodeProperties.Add "Redtail_ID", .redtailID
        nodeProperties.Add "Active", CStr(.Active)
        nodeProperties.Add "Deceased", CStr(.Deceased)
        nodeProperties.Add "Status", .Status
        nodeProperties.Add "Date_of_Death", .dateOfDeath
        nodeProperties.Add "First_Name", .fName
        nodeProperties.Add "Last_Name", .lName
        nodeProperties.Add "Full_Name", .NameOfMember
    End With
    
    'Return the properties
    Set MemberNodeProperties = nodeProperties
End Property

Public Property Get AccountNodeProperties(givenAccount As clsAccount) As Dictionary
    'Set properties for the Account
    Dim nodeProperties As Dictionary
    Set nodeProperties = New Dictionary
    With givenAccount
        nodeProperties.Add "Morningstar_ID", .morningstarID
        nodeProperties.Add "Redtail_ID", .redtailID
        nodeProperties.Add "Name", .NameOfAccount
        nodeProperties.Add "Number", .Number
        nodeProperties.Add "Type", .TypeOfAccount
        nodeProperties.Add "Custodian", .custodian
        nodeProperties.Add "Owner", .owner.NameOfMember
        nodeProperties.Add "Active", CStr(.Active)
        nodeProperties.Add "Balance", .Balance
        nodeProperties.Add "Tag", AutoTag(.NameOfAccount, LoadAssociatedAccounts)
        nodeProperties.Add "Open_Date", .openDate
        nodeProperties.Add "Close_Date", .closeDate
        nodeProperties.Add "Discretionary", .Discretionary
    End With
    
    'Return the properties
    Set AccountNodeProperties = nodeProperties
End Property

Public Property Get BeneficiaryNodeProperties(givenBeneficiary As clsBeneficiary) As Dictionary
    'Set properties for the Beneficiary
    Dim nodeProperties As Dictionary
    Set nodeProperties = New Dictionary
    With givenBeneficiary
        If .id = 0 Then
            nodeProperties.Add "ID", vbNullString
        Else
            nodeProperties.Add "ID", .id
        End If
        nodeProperties.Add "Name", .NameOfBeneficiary
        nodeProperties.Add "Relationship", .Relation
        nodeProperties.Add "Level", .Level
        nodeProperties.Add "Percent", .Percent
        nodeProperties.Add "Last_Updated", .addDate
        nodeProperties.Add "Updated_By", .AddedBy
    End With
    
    'Return the properties
    Set BeneficiaryNodeProperties = nodeProperties
End Property

Public Property Get RedtailActiveStatuses() As String()
    Dim statuses(0 To 3) As String
    statuses(0) = "Active"
    statuses(1) = "Active Spouse/Partner"
    statuses(2) = "Client Child"
    statuses(3) = "Client Relative"
    
    RedtailActiveStatuses = statuses
End Property

Public Function LoadAssociatedAccounts() As String()
    If Dir(associatedFileLocation) <> vbNullString Then
        'The file exists; load it
        Dim fs As FileSystemObject
        Set fs = New FileSystemObject
        Dim associatedFile As TextStream
        Set associatedFile = fs.OpenTextFile(associatedFileLocation, ForReading, True)
        
        'Return the array of Associated account names
        LoadAssociatedAccounts = Split(associatedFile.ReadAll, vbLf)
        
        'Close the file
        associatedFile.Close
    Else
        'Return an empty string in the first index
        ReDim LoadAssociatedAccounts(0) As String
    End If
End Function

Public Function AutoTag(accountName As String, associatedAccountNames() As String) As String
    'Add WEC or Associated tags if it's easily identifiable or in the list of known Associated Bank account names
    If Len(accountName) > 0 And (UBound(Filter(associatedAccountNames, accountName)) > -1 Or InStr(accountName, " Associated ") > 0) Then
        AutoTag = "Associated"
    ElseIf InStr(accountName, " WEC ") > 0 And InStr(accountName, " WEC HSA") = 0 Then
        AutoTag = "WEC"
    ElseIf InStr(accountName, "CHARITABLE") > 0 Then
        AutoTag = "Charitable"
    Else
        AutoTag = vbNullString
    End If
End Function

Private Function LoadClientList(Optional isSample As Boolean) As DOMDocument60
    'See if the households file is available
    Dim fso As FileSystemObject
    Set fso = New FileSystemObject
    
    'Set the filepath for the list
    Dim filePath As String
    If isSample Then
        filePath = SampleClientListFilePath
    Else
        filePath = ClientListFilePath
    End If
    
    If fso.FileExists(filePath) Then
        Set LoadClientList = New DOMDocument60
        LoadClientList.Load filePath
    Else
        MsgBox "Client List not found in default location."
        End
    End If
End Function

Public Sub StateToggle(turnScreenOn As Boolean)
    With Application
        .ScreenUpdating = turnScreenOn
        .EnableAnimations = turnScreenOn
        .EnableEvents = turnScreenOn
        .DisplayStatusBar = turnScreenOn
        
        If turnScreenOn Then
            .Calculation = xlCalculationAutomatic
        Else
            .Calculation = xlCalculationManual
        End If
    End With
End Sub
