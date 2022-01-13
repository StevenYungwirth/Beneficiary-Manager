Attribute VB_Name = "SampleDatabase"
Option Explicit
Private Const lNameFilePath As String = "Z:\YungwirthSteve\Beneficiary Report\Assets\LastNames.txt"
Private Const mFNameFilePath As String = "Z:\YungwirthSteve\Beneficiary Report\Assets\MFirstNames.txt"
Private Const fFNameFilePath As String = "Z:\YungwirthSteve\Beneficiary Report\Assets\FFirstNames.txt"
Private Const arbitraryCutoff As Integer = 1000
Private mFNames() As String
Private fFNames() As String
Private lNames() As String

Private Sub CreateSampleDatabase()
    'Load the name files
    mFNames = LoadNameFile(mFNameFilePath)
    fFNames = LoadNameFile(fFNameFilePath)
    lNames = LoadNameFile(lNameFilePath)
    
    'Load database
    Dim xmlFile As DOMDocument60
    Set xmlFile = XMLReadWrite.LoadClientList
    
    'Get the households
    Dim householdNodes As IXMLDOMNodeList
    Set householdNodes = xmlFile.SelectNodes("//Household")
    
    'Conceal the sensitive information for each household and their components
    Dim householdNode As IXMLDOMElement
    Dim householdnumber As Integer
    For householdnumber = 0 To householdNodes.Length - 1
        'Get the household
        Set householdNode = householdNodes(householdnumber)
        
        'Set the member's new names
        SetNewMemberNames householdNode
        
        'Change the household name
        householdNode.setAttribute "Name", NewHouseholdName(householdNode)
    
        'Get the accounts
        Dim accountNodes As IXMLDOMNodeList
        Set accountNodes = householdNode.SelectNodes("./Member/Account")
        
        'Change account names, numbers, owner names, and beneficiaries
        Dim accountNumber As Integer
        For accountNumber = 0 To accountNodes.Length - 1
            ChangeAccount accountNodes(accountNumber)
        Next accountNumber
    Next householdnumber
    
    'Remove the household nodes after the cutoff
    For householdnumber = householdNodes.Length - 1 To arbitraryCutoff Step -1
        'Get the household
        Set householdNode = householdNodes(householdnumber)
        
        'Remove the node
        householdNode.parentNode.RemoveChild householdNode
    Next householdnumber
    
    'Save the XML in a new location
    xmlFile.Save XMLReadWrite.SampleClientListFile
End Sub

Private Function LoadNameFile(filePath As String) As String()
    If Dir(filePath) <> vbNullString Then
        'The file exists; load it
        Dim fs As FileSystemObject
        Set fs = New FileSystemObject
        Dim nameFile As TextStream
        Set nameFile = fs.OpenTextFile(filePath, ForReading, True)
        
        'Return the array of Associated account names
        Dim returnArr() As String
        returnArr = Split(nameFile.ReadAll, Chr(13))
        
        'Remove extra spaces
        Dim name As Single
        For name = LBound(returnArr) To UBound(returnArr)
            returnArr(name) = Replace(returnArr(name), Chr(10), vbNullString)
        Next name
        
        'Return the array
        LoadNameFile = returnArr
        
        'Close the file
        nameFile.Close
    Else
        'Return an empty string in the first index
        ReDim LoadNameFile(0) As String
    End If
End Function

Private Sub SetNewMemberNames(householdNode As IXMLDOMElement)
    'Get the members
    Dim memberNodes As IXMLDOMNodeList
    Set memberNodes = householdNode.SelectNodes("./Member")
    
    'Get the last name
    Dim lastName As String
    lastName = Trim(lNames(Int((UBound(lNames) + 1) * Rnd)))
    
    'Change the first names
    Dim memberNode As IXMLDOMElement
    If memberNodes.Length = 1 Then
        RenameMember memberNodes(0), lastName, "Random"
    Else
        'Rename the first member
        RenameMember memberNodes(0), lastName, "Male"
    
        'Rename the second member
        RenameMember memberNodes(1), lastName, "Female"
        
        'Rename any extra members
        If memberNodes.Length > 2 Then
            Dim memberNumber As Integer
            For memberNumber = 2 To memberNodes.Length - 1
                RenameMember memberNodes(memberNumber), lastName, "Random"
            Next memberNumber
        End If
    End If
End Sub

Private Sub RenameMember(memberNode As IXMLDOMElement, lastName As String, gender As String)
    'Set the first name
    If gender = "Male" Then
        memberNode.setAttribute "First_Name", Trim(mFNames(Int((UBound(mFNames) + 1) * Rnd)))
    ElseIf gender = "Female" Then
        memberNode.setAttribute "First_Name", Trim(fFNames(Int((UBound(fFNames) + 1) * Rnd)))
    Else
        memberNode.setAttribute "First_Name", RandomFirstName
    End If
    
    'Set the last name
    memberNode.setAttribute "Last_Name", lastName
End Sub

Private Function RandomFirstName() As String
    'Select whether the single person is male or female
    Dim rndNum
    rndNum = Int(2 * Rnd)
    
    If rndNum = 0 Then
        'Male
        RandomFirstName = Trim(mFNames(Int((UBound(mFNames) + 1) * Rnd)))
    Else
        'Female
        RandomFirstName = Trim(fFNames(Int((UBound(fFNames) + 1) * Rnd)))
    End If
End Function

Private Function NewHouseholdName(householdNode As IXMLDOMElement) As String
    'Get the members
    Dim memberNodes As IXMLDOMNodeList
    Set memberNodes = householdNode.SelectNodes("./Member")
    
    'Get the first member
    Dim memberNode As IXMLDOMElement
    Set memberNode = memberNodes(0)
    
    'Get the last name
    Dim lastName As String
    lastName = memberNode.getAttribute("Last_Name")
    
    'Set the household name
    NewHouseholdName = lastName & ", " & memberNode.getAttribute("First_Name")
    If memberNodes.Length > 1 Then
        Set memberNode = memberNodes(1)
        NewHouseholdName = NewHouseholdName & " & " & memberNode.getAttribute("First_Name")
    End If
End Function

Private Sub ChangeAccount(accountNode As IXMLDOMElement)
    'Get the member first and last names
    Dim firstName As String, lastName As String
    firstName = accountNode.parentNode.Attributes(0).Text
    lastName = accountNode.parentNode.Attributes(1).Text
    
    'Change the owner
    accountNode.setAttribute "Owner", lastName & ", " & firstName
    
    'Change the name
    If accountNode.getAttribute("Custodian") = "TD Ameritrade Institutional" Then
        accountNode.setAttribute "Name", firstName & " " & lastName & " " & UCase(accountNode.getAttribute("Type"))
    Else
        accountNode.setAttribute "Name", firstName & " " & lastName & " HELDAWAY"
    End If
    
    'Change the number
    accountNode.setAttribute "Number", NewNumber
    
    'Get the beneficiaries
    Dim beneNodes As IXMLDOMNodeList
    Set beneNodes = accountNode.SelectNodes("./Beneficiary")
    
    'Change the beneficiaries
    Dim beneNumber As Integer
    For beneNumber = 0 To beneNodes.Length - 1
        ChangeBeneficiary beneNodes(beneNumber)
    Next beneNumber
End Sub

Private Sub ChangeBeneficiary(beneNode As IXMLDOMElement)
    'Get the account owner's last name
    Dim lastName As String
    lastName = beneNode.parentNode.parentNode.Attributes(1).Text
    
    'Change the beneficiary name
    beneNode.setAttribute "Name", RandomFirstName & " " & lastName
    
    'Remove "Added_By"/"Updated_By" names, if present
    If Not IsNull(beneNode.getAttribute("Added_By")) Then
        beneNode.setAttribute "Added_By", vbNullString
    End If
    If Not IsNull(beneNode.getAttribute("Updated_By")) Then
        beneNode.setAttribute "Updated_By", vbNullString
    End If
End Sub

Private Function NewNumber() As String
    Dim returnStr As String
    Dim i As Integer
    For i = 0 To 5
        returnStr = returnStr & CStr(Chr(Int((57 - 49 + 1) * Rnd + 49)))
    Next i
    NewNumber = returnStr
End Function
