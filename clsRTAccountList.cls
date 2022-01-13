VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsRTAccountList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_datasheet As clsDataSheet
Private Const worksheetName As String = "RT Accounts"
Private Const idHeader As String = "Account Id"
Private Const nameHeader As String = "Contact Name"
Private Const accountNumberHeader As String = "Account Number"
Private Const accountTypeHeader As String = "Type"
Private Const custodianHeader As String = "Company"

Public Sub ClassBuilder(path As String)
    'Set the required headers
    Dim requiredHeaders() As String
    requiredHeaders = Split(idHeader & "," & nameHeader & "," & accountNumberHeader & "," & accountTypeHeader & "," & custodianHeader, ",")
    
    'Set the datasheet in order to get the data
    Set m_datasheet = ClassConstructor.NewDataSheet(filePath:=path, wkstName:=worksheetName, reqHeaders:=requiredHeaders)
End Sub

Public Sub AddAccountIDs(MemberDict As Dictionary, householdDict As Dictionary)
    'Initialize a string to write to the error file
    Dim errorString As String
    
    'For each row, find the member, and find the account by its number
    Dim rw As Integer
    For rw = 1 To UBound(m_datasheet.data, 1) - 1
        'Get the contact name on this row. Redtail puts a double-space before the nickname; remove the extra space
        Dim contactName As String
        contactName = Trim(Replace(GetData(nameHeader, rw), "  ", " "))
        
        'See if the member on this row is in the member dictionary
        If MemberDict.Exists(contactName) Then
            'The member is in the dictionary; get the account on this row
            Dim accountNumber As String
            accountNumber = Trim(GetData(accountNumberHeader, rw))
            
            'Find the account number in the member's accounts
            Dim tempMember As clsMember
            Set tempMember = MemberDict.Item(contactName)
            Dim tempAccount As clsAccount
            Set tempAccount = tempMember.GetAccountByNumber(accountNumber)
            
            'If the account couldn't be found, see if the household has an adjoining AUA
            If tempAccount Is Nothing And householdDict.Exists(tempMember.ContainingHousehold.NameOfHousehold & " AUA") Then
                Dim auaHousehold As clsHousehold
                Set auaHousehold = householdDict.Item(tempMember.ContainingHousehold.NameOfHousehold & " AUA")
                Set tempAccount = auaHousehold.GetAccountByNumber(accountNumber)
            End If
        
            'If the account is there, add its ID
            If Not tempAccount Is Nothing Then
                tempAccount.ID = GetData(idHeader, rw)
            ElseIf tempMember.Active Then
                'The account wasn't found and the member is active, add the account as an inactive account
                Dim tempInactiveAccount As clsAccount
                Set tempInactiveAccount = ClassConstructor.NewAccount(accountName:=vbNullString, accountNumber:=accountNumber, _
                                                                      accountType:=GetData(accountTypeHeader, rw), accountCustodian:=GetData(custodianHeader, rw), _
                                                                      accountTag:=vbNullString)
                tempInactiveAccount.ID = GetData(idHeader, rw)
                tempMember.AddAccount tempInactiveAccount
                
                'Log the missing account to the error file
                errorString = errorString & vbNewLine & accountNumber & " is in Redtail, but not in Morningstar for " & contactName
            End If
        Else
            'The member wasn't in the dictionary, log it to the error file
            errorString = errorString & vbNewLine & contactName & " has an account in Redtail, but couldn't be found in Morningstar"
        End If
    Next rw
    
    'Log all errors to the error file
    ErrorHandling.LogErrorToFile errorString
End Sub

Public Function GetData(headerName As String, datapoint As Integer) As String
    GetData = m_datasheet.GetData(headerName, datapoint)
End Function

Public Sub SetData(newData As String, headerName As String, datapoint As Integer)
    m_datasheet.SetData newData, headerName, datapoint
End Sub

Public Sub FillWorksheet(sht As Worksheet)
    m_datasheet.FillWorksheet sht
End Sub
