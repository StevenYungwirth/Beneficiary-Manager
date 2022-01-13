Attribute VB_Name = "ClassConstructor"
Option Explicit

Public Function NewHousehold(householdName As String, Optional householdID As Long) As clsHousehold
    Dim tempHousehold As clsHousehold
    Set tempHousehold = New clsHousehold
    tempHousehold.NameOfHousehold = householdName
    If householdID <> 0 Then
        tempHousehold.redtailID = householdID
    End If
    Set NewHousehold = tempHousehold
End Function

Public Function NewMember(memberFullName As String, memberFirstName As String, memberLastName As String, isActive As Boolean, isDeceased As Boolean) As clsMember
    Set NewMember = New clsMember
    With NewMember
        .FullName = memberFullName
        .FName = memberFirstName
        .LName = memberLastName
        .Active = isActive
        .Deceased = isDeceased
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
        .AddDate = Format(Now(), "m/d/yy h:mm;@")
        .AddedBy = VBA.Environ("username")
    End With
End Function

Public Function NewTDABeneList(filePath As String) As clsTDABeneList
    Dim tempList As clsTDABeneList
    Set tempList = New clsTDABeneList
    tempList.ClassBuilder path:=filePath
    Set NewTDABeneList = tempList
End Function

Public Function NewMSAccountList(filePath As String, Optional sort As Boolean) As clsMSAccountList
    Set NewMSAccountList = New clsMSAccountList
    NewMSAccountList.ClassBuilder path:=filePath, sort:=sort
End Function

Public Function NewRTAccountList(filePath As String) As clsRTAccountList
    Set NewRTAccountList = New clsRTAccountList
    NewRTAccountList.ClassBuilder path:=filePath
End Function

Public Function NewRTContactList(filePath As String, Optional sort As Boolean) As clsRTContactList
    Set NewRTContactList = New clsRTContactList
    NewRTContactList.ClassBuilder path:=filePath, sort:=sort
End Function

Public Function NewManualSheet() As clsManualSheet
    Set NewManualSheet = New clsManualSheet
    NewManualSheet.ClassBuilder
End Function

Public Function NewDataSheet(filePath As String, wkstName As String, reqHeaders As Variant, Optional sortColumnHeader As String) As clsDataSheet
    Set NewDataSheet = New clsDataSheet
    NewDataSheet.ClassBuilder path:=filePath, worksheetName:=wkstName, requiredHeaders:=reqHeaders, sortColumn:=sortColumnHeader
End Function
