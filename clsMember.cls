VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMember"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_redtailID As Long
Private m_type As String
Private m_name As String
Private m_fName As String
Private m_lName As String
Private m_nickname As String
Private m_status As String
Private m_dateOfDeath As String
Private m_household As clsHousehold
Private m_accounts As Dictionary
Private m_deceased As Boolean
Private m_active As Boolean

Private Sub Class_Initialize()
    Set m_household = New clsHousehold
    Set m_accounts = New Dictionary
End Sub

Public Property Get redtailID() As Long
    redtailID = m_redtailID
End Property

Public Property Let redtailID(value As Long)
    m_redtailID = value
End Property

Public Property Get TypeOfMember() As String
    TypeOfMember = m_type
End Property

Public Property Let TypeOfMember(value As String)
    m_type = value
End Property

Public Property Get NameOfMember() As String
    If Len(m_name) = 0 And Len(m_fName) > 0 And Len(m_lName) > 0 Then
        'Return "[Last], [First]"
        NameOfMember = m_lName & ", " & m_fName
    Else
        NameOfMember = m_name
    End If
End Property

Public Property Let NameOfMember(value As String)
    m_name = Trim(value)
    
    'Extract the first name, last name, and nickname from value
    SplitName Trim(value)
End Property

Public Property Get fName() As String
    fName = m_fName
End Property

Public Property Let fName(value As String)
    m_fName = Trim(value)
End Property

Public Property Get lName() As String
    lName = m_lName
End Property

Public Property Let lName(value As String)
    m_lName = Trim(value)
End Property

Public Property Get nickname() As String
    nickname = m_nickname
End Property

Public Property Let nickname(value As String)
    m_nickname = Trim(value)
End Property

Public Property Get Status() As String
    Status = m_status
End Property

Public Property Let Status(value As String)
    m_status = value
End Property

Public Property Get dateOfDeath() As String
    dateOfDeath = m_dateOfDeath
End Property

Public Property Let dateOfDeath(value As String)
    If value = vbNullString Then
        m_dateOfDeath = value
    Else
        m_dateOfDeath = Format(value, "mm/dd/yyyy")
    End If
End Property

Public Property Get Deceased() As Boolean
    Deceased = (m_dateOfDeath <> vbNullString)
End Property

Public Property Get Active() As Boolean
    Active = m_active
End Property

Public Property Let Active(value As Boolean)
    m_active = value
End Property

Public Property Get Accounts() As Dictionary
    Set Accounts = m_accounts
End Property

Public Property Get SortedAccounts() As Dictionary
    If m_accounts.count <= 1 Then
        'Sorting isn't necessary
        Set SortedAccounts = m_accounts
        Exit Property
    End If

    'Convert the dictionary into an array
    Dim accountArray() As clsAccount
    ReDim accountArray(0 To m_accounts.count - 1) As clsAccount
    Dim memberAccount As Integer
    For memberAccount = 0 To m_accounts.count - 1
        Set accountArray(memberAccount) = m_accounts.Items(memberAccount)
    Next memberAccount
    
    'Sort the array by account name
    Dim sortedAccountArray() As clsAccount
    sortedAccountArray = SortAccounts(accountArray)
    
    'Convert the array back into a dictionary and return it
    Set SortedAccounts = New Dictionary
    For memberAccount = 0 To UBound(sortedAccountArray)
        'Add the account to the dictionary, using different keys for heldaway accounts
        If sortedAccountArray(memberAccount).custodian = ProjectGlobals.DefaultCustodian Then
            If SortedAccounts.Exists(sortedAccountArray(memberAccount).Number) Then Stop
            SortedAccounts.Add sortedAccountArray(memberAccount).Number, sortedAccountArray(memberAccount)
        Else
            If SortedAccounts.Exists(sortedAccountArray(memberAccount).NameOfAccount & sortedAccountArray(memberAccount).Number) Then Stop
            SortedAccounts.Add sortedAccountArray(memberAccount).NameOfAccount & sortedAccountArray(memberAccount).Number, sortedAccountArray(memberAccount)
        End If
    Next memberAccount
End Property

Public Property Get ContainingHousehold() As clsHousehold
    Set ContainingHousehold = m_household
End Property

Public Property Let ContainingHousehold(value As clsHousehold)
    Set m_household = value
End Property

Public Sub AddAccount(acct As clsAccount)
    acct.owner = Me
    If m_accounts.Exists(acct.NameOfAccount & acct.Number) Then Stop
    m_accounts.Add acct.NameOfAccount & acct.Number, acct
End Sub

Public Function GetAccountByNumber(AcctNumber As String) As clsAccount
    'Look in each account for one with a matching number
    Dim acct As Variant
    For Each acct In m_accounts.Items
        If acct.Number = AcctNumber Then
            Set GetAccountByNumber = acct
            Exit Function
        End If
    Next acct
End Function

Public Function ActiveAccountsCount() As Integer
    Dim count As Integer
    Dim acct As Variant
    For Each acct In m_accounts.Items
        If acct.Active And acct.Balance > 0 Then
            count = count + 1
        End If
    Next acct
    ActiveAccountsCount = count
End Function

Private Sub SplitName(fullName As String)
    If InStr(fullName, ",") = 0 Then
        'Name cannot be split
        Exit Sub
    End If
    
    Dim firstName As String, lastName As String, nickname As String
    
    'Take everything before the comma as the last name
    lastName = Left(fullName, InStr(fullName, ",") - 1)
    
    'Take everything after the comma as the first name
    firstName = Right(fullName, Len(fullName) - InStr(fullName, ",") - 1)
    
    'Extract the nickname if the name has a parenthesis
    If InStr(firstName, "(") > 0 Then
        'Take everything between the parentheses as the nickname
        nickname = Mid(firstName, InStr("(", firstName) + 1, Len(firstName) - InStr("(", firstName) - 1)
        
        'Reset the first name
        firstName = Left(firstName, InStr(firstName, "(") - 2)
    End If
    
    'Take out the middle initial
    If InStr(firstName, " ") > 0 Then
        If Right(firstName, 1) = "." Then
            firstName = Left(firstName, Len(firstName) - 3)
        ElseIf Mid(firstName, Len(firstName) - 1, 1) = " " Then
            firstName = Left(firstName, Len(firstName) - 2)
        End If
    End If
    
    'Set the first, last, and nicknames
    m_fName = firstName
    m_lName = lastName
    m_nickname = nickname
End Sub

Private Function SortAccounts(accountArr() As clsAccount) As clsAccount()
    Dim i As Integer
    Dim j As Integer
    For i = LBound(accountArr) To UBound(accountArr)
        'Set the lowest name to be the one in the ith index
        Dim lowestIndex As Integer
        lowestIndex = i
        
        'Loop through the rest of the names to find the lowest index
        For j = i + 1 To UBound(accountArr)
            If accountArr(j).NameOfAccount < accountArr(lowestIndex).NameOfAccount Then
                lowestIndex = j
            End If
        Next j
        
        'Swap the lowest index with the first one
        Dim temp As clsAccount
        Set temp = accountArr(lowestIndex)
        Set accountArr(lowestIndex) = accountArr(i)
        Set accountArr(i) = temp
    Next i
    
    SortAccounts = accountArr
End Function
