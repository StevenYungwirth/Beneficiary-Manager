VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsHousehold"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_name As String
Private m_members As Dictionary
Private m_morningstarID As String
Private m_redtailID As Long
Private m_active As Boolean

Private Sub Class_Initialize()
    Set m_members = New Dictionary
End Sub

Public Property Get morningstarID() As String
    morningstarID = m_morningstarID
End Property

Public Property Let morningstarID(value As String)
    m_morningstarID = value
End Property

Public Property Get redtailID() As Long
    redtailID = m_redtailID
End Property

Public Property Let redtailID(value As Long)
    m_redtailID = value
End Property

Public Property Get NameOfHousehold() As String
    NameOfHousehold = m_name
End Property

Public Property Let NameOfHousehold(value As String)
    m_name = value
End Property

Public Property Get Active() As Boolean
    'Return true if at least one member is active
    Active = m_active
End Property

Public Property Let Active(value As Boolean)
    m_active = value
End Property

Public Property Get members() As Dictionary
    Set members = m_members
End Property

Public Property Get ActiveMemberCount() As Integer
    Dim count As Integer
    Dim member As Variant
    For Each member In m_members.Items
        If member.Active Then
            count = count + 1
        End If
    Next member
    ActiveMemberCount = count
End Property

Public Property Get SortedMembers() As Dictionary
    If m_members.count <= 1 Then
        'Sorting isn't necessary
        Set SortedMembers = m_members
        Exit Property
    End If

    'Convert the dictionary into an array
    Dim memberArray() As clsMember
    ReDim memberArray(0 To m_members.count - 1) As clsMember
    Dim householdMember As Integer
    For householdMember = 0 To m_members.count - 1
        Set memberArray(householdMember) = m_members.Items(householdMember)
    Next householdMember
    
    'Sort the array by member name
    SortMembers memberArray
    
    'Convert the array back into a dictionary and return it
    Set SortedMembers = New Dictionary
    For householdMember = 0 To UBound(memberArray)
        SortedMembers.Add memberArray(householdMember).NameOfMember, memberArray(householdMember)
    Next householdMember
End Property

Public Sub AddMember(member As clsMember)
    If Not m_members.Exists(member.NameOfMember) Then
        'The member isn't in the household; add them
        member.ContainingHousehold = Me
        m_members.Add member.lName & ", " & member.fName, member
    End If
End Sub

Public Function GetMemberFromSeparatedName(lastName As String, firstName As String) As clsMember
    'If the member name is in the member dictionary, return that member
    Dim memberKey As String
    memberKey = lastName & ", " & firstName
    If m_members.Exists(memberKey) Then
        Set GetMemberFromSeparatedName = m_members(memberKey)
    End If
End Function

Public Function GetMemberFromLastCommaFirst(lastNameCommaFirstName As String) As clsMember
    'If the member name is in the member dictionary, return that member
    If m_members.Exists(lastNameCommaFirstName) Then
        Set GetMemberFromLastCommaFirst = m_members(lastNameCommaFirstName)
    End If
End Function

Public Function GetAccountByNumber(acctNum As String) As clsAccount
    'Iterate through each member and their accounts to find the account with matching name and number
    Dim acctMemberItem As Variant
    For Each acctMemberItem In m_members.Items
        Dim acctMember As clsMember
        Set acctMember = acctMemberItem
        If acctMember.Accounts.Exists(acctNum) Then
            Set GetAccountByNumber = acctMember.Accounts(acctNum)
            Exit Function
        End If
    
        Dim account As Variant
        For Each account In acctMember.Accounts.Items
            If account.Number = acctNum Then
                'The account number matches; return the account
                Set GetAccountByNumber = account
                Exit Function
            End If
        Next account
    Next acctMemberItem
End Function

Private Sub SortMembers(memberArr() As clsMember)
    Dim i As Integer
    Dim j As Integer
    For i = LBound(memberArr) To UBound(memberArr)
        'Set the lowest name to be the one in the ith index
        Dim lowestIndex As Integer
        lowestIndex = i
        
        'Loop through the rest of the names to find the lowest index
        For j = i + 1 To UBound(memberArr)
            If memberArr(j).NameOfMember < memberArr(lowestIndex).NameOfMember Then
                lowestIndex = j
            End If
        Next j
        
        'Swap the lowest index with the first one
        Dim temp As clsMember
        Set temp = memberArr(lowestIndex)
        Set memberArr(lowestIndex) = memberArr(i)
        Set memberArr(i) = temp
    Next i
End Sub
