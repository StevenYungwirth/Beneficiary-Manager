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
Private m_name As String
Private m_fName As String
Private m_lName As String
Private m_fullName As String
Private m_household As clsHousehold
Private m_accounts As Dictionary
Private m_deceased As Boolean
Private m_active As Boolean

Private Sub Class_Initialize()
    Set m_household = New clsHousehold
    Set m_accounts = New Dictionary
End Sub

Public Property Get NameOfMember() As String
    If Len(m_name) = 0 And Len(m_fName) > 0 And Len(m_lName) > 0 Then
        'Return "[Last], [First]"
        NameOfMember = m_lName & ", " & m_fName
    Else
        NameOfMember = m_name
    End If
End Property

Public Property Let NameOfMember(value As String)
    m_name = value
    
    'If name has a comma, split the first and last name
    If InStr(value, ",") > 0 Then
        'Take everything before the comma as the last name
        m_lName = Left(value, InStr(value, ",") - 1)
        
        'Take everything after the comma as the first name
        m_fName = Right(value, Len(value) - InStr(value, ",") - 1)
    End If
End Property

Public Property Get FName() As String
    FName = m_fName
End Property

Public Property Let FName(value As String)
    m_fName = value
End Property

Public Property Get LName() As String
    LName = m_lName
End Property

Public Property Let LName(value As String)
    m_lName = value
End Property

Public Property Get FullName() As String
    If Len(m_fullName) = 0 And Len(m_fName) > 0 And Len(m_lName) > 0 Then
        'Return "[Last], [First]"
        FullName = m_lName & ", " & m_fName
    Else
        FullName = m_fullName
    End If
End Property

Public Property Let FullName(value As String)
    m_fullName = value
End Property

Public Property Get Deceased() As Boolean
    Deceased = m_deceased
End Property

Public Property Let Deceased(value As Boolean)
    m_deceased = value
End Property

Public Property Get Active() As Boolean
    Active = m_active
End Property

Public Property Let Active(value As Boolean)
    m_active = value
End Property

Public Property Get accounts() As Dictionary
    Set accounts = m_accounts
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
    SortAccounts accountArray
    
    'Convert the array back into a dictionary and return it
    Set SortedAccounts = New Dictionary
    For memberAccount = 0 To UBound(accountArray)
        'Add the account to the dictionary, using different keys for heldaway accounts
        If accountArray(memberAccount).custodian = "TD Ameritrade Institutional" Then
            If SortedAccounts.Exists(accountArray(memberAccount).Number) Then Stop
            SortedAccounts.Add accountArray(memberAccount).Number, accountArray(memberAccount)
        Else
            If SortedAccounts.Exists(accountArray(memberAccount).NameOfAccount & accountArray(memberAccount).Number) Then Stop
            SortedAccounts.Add accountArray(memberAccount).NameOfAccount & accountArray(memberAccount).Number, accountArray(memberAccount)
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
    acct.Owner = Me
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

Private Sub SortAccounts(accountArr() As clsAccount)
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
End Sub
