VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsAccount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_morningstarID As String
Private m_redtailID As Double
Private m_name As String
Private m_number As String
Private m_balance As Double
Private m_owner As clsMember
Private m_type As String
Private m_custodian As String
Private m_active As Boolean
Private m_benes As Dictionary
Private m_openDate As String
Private m_closeDate As String
Private m_discretionary As Boolean
Private m_tag As String

Private Sub Class_Initialize()
    Set m_benes = New Dictionary
    Set m_owner = New clsMember
End Sub

Public Property Get morningstarID() As String
    morningstarID = m_morningstarID
End Property

Public Property Let morningstarID(value As String)
    m_morningstarID = value
End Property

Public Property Get redtailID() As Double
    redtailID = m_redtailID
End Property

Public Property Let redtailID(value As Double)
    m_redtailID = value
End Property

Public Property Get NameOfAccount() As String
    NameOfAccount = m_name
End Property

Public Property Let NameOfAccount(value As String)
    m_name = value
End Property

Public Property Get Number() As String
    Number = m_number
End Property

Public Property Let Number(value As String)
    m_number = value
End Property

Public Property Get Balance() As Double
    Balance = Round(m_balance, 2)
End Property

Public Property Let Balance(value As Double)
    m_balance = value
End Property

Public Property Get owner() As clsMember
    Set owner = m_owner
End Property

Public Property Let owner(value As clsMember)
    Set m_owner = value
End Property

Public Property Get TypeOfAccount() As String
    TypeOfAccount = m_type
End Property

Public Property Let TypeOfAccount(value As String)
    m_type = value
End Property

Public Property Get custodian() As String
    If m_custodian = "Unknown" Then
        custodian = "N/A (Non-TD Ameritrade)"
    Else
        custodian = m_custodian
    End If
End Property

Public Property Let custodian(value As String)
    m_custodian = value
End Property

Public Property Get Active() As Boolean
    Active = m_active
End Property

Public Property Let Active(value As Boolean)
    m_active = value
End Property

Public Property Let openDate(value As String)
    If value = vbNullString Then
        m_openDate = value
    Else
        m_openDate = Format(value, "mm/dd/yyyy")
    End If
End Property

Public Property Get openDate() As String
    openDate = m_openDate
End Property

Public Property Let closeDate(value As String)
    If value = vbNullString Then
        m_closeDate = value
    Else
        m_closeDate = Format(value, "mm/dd/yyyy")
    End If
End Property

Public Property Get closeDate() As String
    closeDate = m_closeDate
End Property

Public Property Let Discretionary(value As Boolean)
    m_discretionary = value
End Property

Public Property Get Discretionary() As Boolean
    Discretionary = m_discretionary
End Property

Public Property Get Benes() As Dictionary
    Set Benes = m_benes
End Property

Public Property Get SortedBenes() As Dictionary
    If m_benes.count <= 1 Then
        'Sorting isn't necessary
        Set SortedBenes = m_benes
        Exit Property
    End If
    
    'Split the beneficiaries into primary and contingent arrays
    Dim primaryArray() As clsBeneficiary
    ReDim primaryArray(0) As clsBeneficiary
    Dim contingentArray() As clsBeneficiary
    ReDim contingentArray(0) As clsBeneficiary
    Dim accountBene As Variant
    For Each accountBene In m_benes.Items
        Dim bene As clsBeneficiary
        Set bene = accountBene
        If bene.Level = "P" Then
            AddToTopOfArray bene, primaryArray
        ElseIf bene.Level = "C" Then
            AddToTopOfArray bene, contingentArray
        End If
    Next accountBene
    
    'Sort each array by percent and then by name
    SortBenes beneArr:=primaryArray, One_Percent_Two_Name:=1
    SortBenes beneArr:=primaryArray, One_Percent_Two_Name:=2
    SortBenes beneArr:=contingentArray, One_Percent_Two_Name:=1
    SortBenes beneArr:=contingentArray, One_Percent_Two_Name:=2
    
    'Convert the arrays back into a collection and return it
    Set SortedBenes = New Dictionary
    If Not primaryArray(0) Is Nothing Then
        For accountBene = LBound(primaryArray) To UBound(primaryArray)
            SortedBenes.Add accountBene, primaryArray(accountBene)
        Next accountBene
    End If
    If Not contingentArray(0) Is Nothing Then
        For accountBene = LBound(contingentArray) To UBound(contingentArray)
            SortedBenes.Add UBound(primaryArray) + accountBene + 1, contingentArray(accountBene)
        Next accountBene
    End If
End Property

Public Property Let Tag(value As String)
    m_tag = value
End Property

Public Property Get Tag() As String
    Tag = m_tag
End Property

Public Function BenesUpdated() As Date
    'Get the most recent update date from the beneficiaries
    Dim mostRecentDate As Date
    Dim beneItem As Variant
    For Each beneItem In m_benes.Items
        Dim bene As clsBeneficiary
        Set bene = beneItem
        Dim beneUpdateDate As Date
        If bene.UpdatedDate = vbNullString Then
            beneUpdateDate = CDate(0)
        Else
            beneUpdateDate = CDate(bene.UpdatedDate)
        End If
        If beneUpdateDate > mostRecentDate Then
            mostRecentDate = beneUpdateDate
        End If
    Next beneItem
    BenesUpdated = mostRecentDate
End Function

Public Sub AddBene(beneToAdd As clsBeneficiary, hasID As Boolean)
    If Not beneToAdd Is Nothing Then
        Dim beneKey As String
        If hasID Then
            beneKey = beneToAdd.id
        Else
            beneKey = beneToAdd.NameOfBeneficiary & beneToAdd.Level & beneToAdd.Percent
        End If
        
        Dim counter As Integer
        Do While m_benes.Exists(beneKey)
            counter = counter + 1
            beneKey = beneKey & counter
            
            Dim errorString As String
            With beneToAdd
                errorString = "Duplicate beneficiary found for " & .account.NameOfAccount & " " & .account.Number & ": " & .NameOfBeneficiary & " " & .Level & " " & .Percent
            End With
            ErrorHandling.LogErrorToFile errorString
        Loop
        
        beneToAdd.account = Me
        m_benes.Add beneKey, beneToAdd
    End If
End Sub

Public Function GetBene(beneName As String) As clsBeneficiary
    Dim beneItem As Variant
    For Each beneItem In m_benes.Items
        Dim bene As clsBeneficiary
        Set bene = beneItem
        If bene.NameOfBeneficiary = beneName Then
            Set GetBene = bene
        End If
    Next beneItem
End Function

Public Sub UpdateBene(beneIDToUpdate As Double, NewBene As clsBeneficiary)
    Dim beneToUpdate As clsBeneficiary
    Set beneToUpdate = GetBeneByID(beneIDToUpdate)
    With beneToUpdate
        .NameOfBeneficiary = NewBene.NameOfBeneficiary
        .Level = NewBene.Level
        .Percent = NewBene.Percent
        .UpdatedDate = NewBene.addDate
        .updatedBy = NewBene.AddedBy
    End With
End Sub

Public Sub RemoveBene(beneToRemove As clsBeneficiary)
    If m_benes.Exists(CStr(beneToRemove.id)) Then
        m_benes.Remove CStr(beneToRemove.id)
    Else
        Dim beneItem As Variant
        For Each beneItem In m_benes.Items
            Dim bene As clsBeneficiary
            Set bene = beneItem
            If bene.id = beneToRemove.id Then
                If Not m_benes.Exists(CStr(beneToRemove.id)) Then Stop
                m_benes.Remove CStr(beneToRemove.id)
                Exit Sub
            ElseIf bene.NameOfBeneficiary = beneToRemove.NameOfBeneficiary And bene.Level = beneToRemove.Level And bene.Percent = beneToRemove.Percent Then
                If Not m_benes.Exists(CStr(bene.id)) Then Stop
                m_benes.Remove CStr(bene.id)
                Exit Sub
            End If
        Next beneItem
    End If
End Sub

Private Sub AddToTopOfArray(value As clsBeneficiary, arr() As clsBeneficiary)
    If arr(0) Is Nothing Then
        Set arr(0) = value
    Else
        ReDim Preserve arr(0 To UBound(arr) + 1) As clsBeneficiary
        Set arr(UBound(arr)) = value
    End If
End Sub

Private Sub SortBenes(beneArr() As clsBeneficiary, One_Percent_Two_Name As Integer)
    Dim i As Integer
    Dim j As Integer
    For i = LBound(beneArr) To UBound(beneArr)
        'Set the lowest name to be the one in the ith index
        Dim lowestIndex As Integer
        lowestIndex = i
        
        'Loop through the rest of the names to find the lowest index
        For j = i + 1 To UBound(beneArr)
            If One_Percent_Two_Name = 1 And beneArr(j).NameOfBeneficiary < beneArr(lowestIndex).NameOfBeneficiary Then
                lowestIndex = j
            ElseIf One_Percent_Two_Name = 2 And beneArr(j).Percent > beneArr(lowestIndex).Percent Then
                lowestIndex = j
            End If
        Next j
        
        'Swap the lowest index with the first one
        Dim temp As clsBeneficiary
        Set temp = beneArr(lowestIndex)
        Set beneArr(lowestIndex) = beneArr(i)
        Set beneArr(i) = temp
    Next i
End Sub

Private Function GetBeneByID(beneID As Double) As clsBeneficiary
    Dim beneItem As Variant
    For Each beneItem In m_benes.Items
        Dim bene As clsBeneficiary
        Set bene = beneItem
        If bene.id = beneID Then
            Set GetBeneByID = bene
        End If
    Next beneItem
End Function
