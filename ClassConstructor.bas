Attribute VB_Name = "ClassConstructor"
Option Explicit

Public Function NewHousehold(householdName As String, Optional morningstarID As String, Optional redtailID As Long) As clsHousehold
    Dim tempHousehold As clsHousehold
    Set tempHousehold = New clsHousehold
    tempHousehold.NameOfHousehold = householdName
    If redtailID <> 0 Then tempHousehold.redtailID = redtailID
    If morningstarID <> vbNullString Then tempHousehold.morningstarID = morningstarID
    Set NewHousehold = tempHousehold
End Function

Public Function NewMember(memberFullName As String, memberFirstName As String, memberLastName As String, memberType As String, memberStatus As String, memberDateOfDeath As String, contactID As Long) As clsMember
    Set NewMember = New clsMember
    With NewMember
        .TypeOfMember = memberType
        .NameOfMember = memberFullName
        .fName = memberFirstName
        .lName = memberLastName
        .Status = memberStatus
        .dateOfDeath = memberDateOfDeath
        .redtailID = contactID
    End With
End Function

Public Function NewMemberNameOnly(memberFullName As String) As clsMember
    Set NewMemberNameOnly = New clsMember
    NewMemberNameOnly.NameOfMember = memberFullName
End Function

Public Function NewAccount(accountName As String, accountNumber As String, accountType As String, accountCustodian As String, accountTag As String, Optional marketValue As Double) As clsAccount
    Set NewAccount = New clsAccount
    With NewAccount
        .NameOfAccount = accountName
        .Number = accountNumber
        .Balance = marketValue
        .TypeOfAccount = accountType
        .custodian = accountCustodian
        .Tag = accountTag
        .Active = False
    End With
End Function

Public Function NewBene(beneName As String, beneLevel As String, benePercent As Double, Optional beneRelation As String) As clsBeneficiary
    Set NewBene = New clsBeneficiary
    With NewBene
        .NameOfBeneficiary = beneName
        .Level = beneLevel
        .Percent = benePercent
        .Relation = beneRelation
        .addDate = Format(ProjectGlobals.ImportTime, "m/d/yyyy h:mm;@")
        .AddedBy = VBA.Environ("username")
    End With
End Function

Public Function NewTDABeneList(filePath As String) As clsTDABeneList
    If filePath = vbNullString Then
        Set NewTDABeneList = Nothing
    Else
        Dim tempList As clsTDABeneList
        Set tempList = New clsTDABeneList
        tempList.ClassBuilder path:=filePath
        Set NewTDABeneList = tempList
    End If
End Function

Public Function NewMSAccountList(filePath As String) As clsMSAccountList
    If filePath = vbNullString Then
        Set NewMSAccountList = Nothing
    Else
        Set NewMSAccountList = New clsMSAccountList
        NewMSAccountList.ClassBuilder path:=filePath
    End If
End Function

Public Function NewRTAccountList(filePath As String) As clsRTAccountList
    If filePath = vbNullString Then
        Set NewRTAccountList = Nothing
    Else
        Set NewRTAccountList = New clsRTAccountList
        NewRTAccountList.ClassBuilder path:=filePath
    End If
End Function

Public Function NewRTContactList(filePath As String) As clsRTContactList
    If filePath = vbNullString Then
        Set NewRTContactList = Nothing
    Else
        Set NewRTContactList = New clsRTContactList
        NewRTContactList.ClassBuilder path:=filePath
    End If
End Function

Public Function NewMSHouseholdExport(filePath As String) As clsMSHouseholdExport
    If filePath = vbNullString Then
        Set NewMSHouseholdExport = Nothing
    Else
        Set NewMSHouseholdExport = New clsMSHouseholdExport
        NewMSHouseholdExport.ClassBuilder path:=filePath
    End If
End Function

Public Function NewManualSheet() As clsManualSheet
    Set NewManualSheet = New clsManualSheet
    NewManualSheet.ClassBuilder
End Function

Public Function NewDataSheet(filePath As String, wkstName As String, reqHeaders As Variant) As clsDataSheet
    Set NewDataSheet = New clsDataSheet
    NewDataSheet.ClassBuilder path:=filePath, worksheetName:=wkstName, requiredHeaders:=reqHeaders
End Function
