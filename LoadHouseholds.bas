Attribute VB_Name = "LoadHouseholds"
Option Explicit

Public Sub CreateXMLFile()
    'Set the datasheet classes to be the sheets in the Excel file
    Dim tdList As clsTDABeneList, msList As clsMSAccountList, rtaList As clsRTAccountList, rtcList As clsRTContactList
    Set tdList = ClassConstructor.NewTDABeneList(vbNullString)
    Set msList = ClassConstructor.NewMSAccountList(vbNullString, True)
    Set rtaList = ClassConstructor.NewRTAccountList(vbNullString)
    Set rtcList = ClassConstructor.NewRTContactList(vbNullString, True)
    
    XMLCreateList.CreateHouseholdsXMLFile GetHouseholds(tdList, msList, rtaList, rtcList)
End Sub

Public Function GetHouseholds(tdaBenes As clsTDABeneList, msAccounts As clsMSAccountList, rtAccounts As clsRTAccountList, rtContacts As clsRTContactList) As Dictionary
    'Get collection of households from Redtail
    Dim householdDict As Dictionary
    Set householdDict = rtContacts.GetHouseholds
    
    'Add Morningstar accounts to the households
    Dim accountDict As Dictionary
    Set accountDict = msAccounts.GetAccounts(householdDict)
    
    'Add beneficiaries to accounts and set active TD accounts
    tdaBenes.AddBenesFromTD accountDict
    
    'Add manually added beneficiaries
    Dim manualSheet As clsManualSheet
    Set manualSheet = ClassConstructor.NewManualSheet
    manualSheet.AddManualBenes accountDict
    
    'Get a list of all members
    Dim MemberDict As Dictionary
    Set MemberDict = GetAllMembers(householdDict)
    
    'Go through RT Accounts sheet, get all members from the household dictionary, find the account by its number, and add account IDs
    rtAccounts.AddAccountIDs MemberDict, householdDict
    
    'Return the household array
    Set GetHouseholds = householdDict
End Function

Private Function GetAllMembers(households As Dictionary) As Dictionary
    'Initialize a temporary dictionary to return
    Dim tempMembers As Dictionary
    Set tempMembers = New Dictionary
    
    'Loop through each household and add their members to the dictionary
    Dim household As Variant
    For Each household In households.Items
        Dim member As Variant
        For Each member In household.Members.Items
            If Not tempMembers.Exists(member.FullName) Then
                'This name hasn't been added yet. Add it to the dictionary
                tempMembers.Add member.FullName, member
            Else
                'This name already exists in the dictionary. Panic'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            End If
        Next member
    Next household
    
    'Return the dictionary
    Set GetAllMembers = tempMembers
End Function
