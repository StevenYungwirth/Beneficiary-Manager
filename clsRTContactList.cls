VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsRTContactList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_datasheet As clsDataSheet
Private Const worksheetName As String = "RT Contacts"
Private Const idHeader As String = "ID"
Private Const familyNameHeader As String = "Family Name"
Private Const firstNameHeader As String = "First Name"
Private Const middleNameHeader As String = "Middle Name"
Private Const lastNameHeader As String = "Last Name"
Private Const suffixHeader As String = "Suffix"
Private Const nicknameHeader As String = "Nickname"
Private Const statusHeader As String = "Status"
Private Const categoryHeader As String = "Category"
Private Const dateOfDeathHeader As String = "Date Of Death"

Public Sub ClassBuilder(path As String)
    'Set the required headers
    Dim requiredHeaders() As String
    requiredHeaders = Split(idHeader & "," & familyNameHeader & "," & firstNameHeader & "," & middleNameHeader & "," & lastNameHeader & "," _
                      & suffixHeader & "," & nicknameHeader & "," & statusHeader & "," & categoryHeader & "," & dateOfDeathHeader, ",")
    
    'Set the datasheet in order to get the data
    Set m_datasheet = ClassConstructor.NewDataSheet(filePath:=path, wkstName:=worksheetName, reqHeaders:=requiredHeaders, sortColumnHeader:=familyNameHeader)
End Sub

Public Function GetHouseholds() As Dictionary
    'Set a temporary dictionary to return
    Dim tempDictionary As Dictionary
    Set tempDictionary = New Dictionary
    
    'Go through each row in the Redtail contacts and add the households and members
    Dim rw As Integer
    For rw = 1 To UBound(m_datasheet.Data, 1) - 1
        'Skip rows if they have no family name
        Dim familyName As String
        familyName = GetData(familyNameHeader, rw)
        If Len(familyName) > 0 Then
            'Get the household to look at
            Dim tempHousehold As clsHousehold
            If tempDictionary.Exists(familyName) Then
                'The family is in the dictionary
                Set tempHousehold = tempDictionary(familyName)
            Else
                'The family's not in the dictionary, add it
                Set tempHousehold = ClassConstructor.NewHousehold(householdName:=familyName, householdID:=CLng(GetData(idHeader, rw)))
                tempDictionary.Add familyName, tempHousehold
            End If
            
            'Get the member's dictionary key
            Dim memberFirstName As String, memberLastName As String, memberKey As String
            memberFirstName = GetData(firstNameHeader, rw)
            memberLastName = GetData(lastNameHeader, rw)
            memberKey = memberLastName & memberFirstName
            
            'Add the member if it's not already in the dictionary
            If Not tempHousehold.Members.Exists(memberKey) Then
                'The member isn't in the dictionary; add it
                Dim memberFullName As String
                memberFullName = CombineNameParts(memberFirstName, GetData(middleNameHeader, rw), memberLastName, GetData(suffixHeader, rw), GetData(nicknameHeader, rw))
                Dim isActive As Boolean
                isActive = (InStr(GetData(statusHeader, rw), "Active") > 0 And GetData(statusHeader, rw) <> "InActive")
                tempHousehold.AddMember ClassConstructor.NewMember(memberFullName:=memberFullName, memberFirstName:=memberFirstName, memberLastName:=memberLastName, _
                                                                    isActive:=isActive, isDeceased:=(Len(GetData(dateOfDeathHeader, rw)) > 0))
            End If
        End If
    Next rw
    
    'Return the dictionary
    Set GetHouseholds = tempDictionary
End Function

Public Function GetData(headerName As String, datapoint As Integer) As String
    GetData = m_datasheet.GetData(headerName, datapoint)
End Function

Public Sub FillWorksheet(sht As Worksheet)
    m_datasheet.FillWorksheet sht
End Sub

Private Function CombineNameParts(firstName As String, middleName As String, lastName As String, suffix As String, nickname As String) As String
    'Put the row's values into an array
    Dim rowArray(4) As String
    rowArray(0) = firstName
    rowArray(1) = middleName
    rowArray(2) = lastName
    rowArray(3) = suffix
    rowArray(4) = nickname
    
    'rowArray contains: First name, middle name, last name, suffix, nickname
    'Build the string as "[last name] [suffix], [first name] [middle name] ([nickname])"
    Dim returnName As String
    returnName = rowArray(2)
    returnName = AddToString(addThisString:=rowArray(3), toThisString:=returnName, separatedBy:=" ")
    returnName = AddToString(addThisString:=rowArray(0), toThisString:=returnName, separatedBy:=", ")
    returnName = AddToString(addThisString:=rowArray(1), toThisString:=returnName, separatedBy:=" ")
    returnName = AddToString(addThisString:=rowArray(4), toThisString:=returnName, separatedBy:=" (")
    If Len(rowArray(4)) > 0 Then
        returnName = returnName & ")"
    End If
    
    'Return the name
    CombineNameParts = returnName
End Function

Private Function AddToString(addThisString As String, toThisString As String, separatedBy As String) As String
    If Len(addThisString) > 0 Then
        AddToString = toThisString & separatedBy & addThisString
    Else
        AddToString = toThisString
    End If
End Function
