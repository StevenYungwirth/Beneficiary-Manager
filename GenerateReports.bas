Attribute VB_Name = "GenerateReports"
Option Explicit
Private PrintRange As Range
Private Const IneligibleBlurb As String = "Account not eligible for beneficiaries"
Private Const AssociatedBlurb As String = "Beneficiary information isn't available online. Please call Associated Bank at 800-236-8866 to verify or change your beneficiaries."
Private Const WECBlurb As String = "Beneficiary information isn't available online. Please contact WEC's Human Resources department at 800-499-2800 to verify or change your beneficiaries."

Private Property Get IneligibleAccounts() As String()
    'Return the list of account types that aren't eligible for beneficiaries
    IneligibleAccounts = Split("Endowment|Foundation|Non-Profit Organization-ex|Non-Profit Organization-nonex|Partnership|Club/Association|" _
    & "Custodian|Custodian Others|Estate|Guardian|Omnibus|Outside Trustee Plans|Power of Attorney|Retirement Non-Fund|Non-Profit/Exempt|" _
    & "Corporation - Non-taxable|Corporation - Taxable|UGMA/UTMA|529 Plan|Corporate", "|")
End Property

Private Property Get TrustAccounts() As String()
    'Return the list of account types with trust documents
    TrustAccounts = Split("Charitable Trust|Trust|Annual Trust|Charitable Remainder Trust|Charitable Trust|Revocable Trust|Trust|Trust, Qualified Plan", "|")
End Property

Public Sub BeneficiaryReport(households As Dictionary)
    'Set the folder to save each report to
    Dim saveFolder As String
    saveFolder = "Z:\Beneficiary Reports\"
    
    'Turn off screen updating except for the status bar
    UpdateScreen "Off"
    Application.DisplayStatusBar = True
    
    'For each household, generate the beneficiary report
    Dim reportsRan As Integer
    Dim household As Integer
    For household = 0 To 10 'households.Count
        'Update status bar to show household name and number
        Application.StatusBar = "Generating report for " & household + 1 & "/" & households.count + 1 & " - " & households.Items(household).NameOfHousehold
        DoEvents
        
        'Generate and save the report
        Dim beneReport As Workbook
        Set beneReport = GenerateReportFromHousehold(households.Item(household))
        
        'Set the location to save each report
        Dim savePath As String
        savePath = saveFolder & households.Item(household).NameOfHousehold & ".xlsx"
        
        'Save the workbook
        Application.DisplayAlerts = False
        beneReport.SaveAs savePath, ConflictResolution:=xlLocalSessionChanges
        Application.DisplayAlerts = True
        
        'Close the workbook
        beneReport.Close
        
        'Increment the report counter
        reportsRan = reportsRan + 1
    Next household
    
    'Turn on screen updating and reset the status bar
    UpdateScreen "On"
    Application.StatusBar = False
    
    'Show completion message
    If households.count = 1 Then
        MsgBox reportsRan & "/" & households.count + 1 & " beneficiary report ran and saved in " & saveFolder
    Else
        MsgBox reportsRan & "/" & households.count + 1 & " beneficiary reports ran and saved in " & saveFolder
    End If
End Sub

Private Sub UpdateScreen(OnOrOff As String)
    Dim Reset As Long
    If OnOrOff = "Off" Then
        Application.ScreenUpdating = False
        Application.EnableEvents = False
        Application.DisplayStatusBar = False
        Application.Calculation = xlCalculationManual
    ElseIf OnOrOff = "On" Then
        Application.ScreenUpdating = True
        Application.EnableEvents = True
        Application.DisplayStatusBar = True
        Application.Calculation = xlCalculationAutomatic
        Reset = ActiveSheet.UsedRange.Rows.count
    End If
End Sub

Public Function GenerateReportFromHousehold(household As clsHousehold) As Workbook
    On Error GoTo BackOn
    
    'Create a new workbook
    Dim reportBook As Workbook
    Set reportBook = Workbooks.Add("Z:\YungwirthSteve\Beneficiary Report\Assets\Bene Template.xltx")
    Dim reportSheet As Worksheet
    Set reportSheet = reportBook.Sheets(1)
    
    'Build report
    BuildReport reportSheet, household
    
    'Return the report workbook
    Set GenerateReportFromHousehold = reportBook
    
    On Error GoTo 0
    Exit Function
BackOn:
    Application.ScreenUpdating = True
    End
End Function

Private Sub BuildReport(reportSheet As Worksheet, household As clsHousehold)
    'Add the date to the header
    reportSheet.PageSetup.LeftHeader = vbLf & "&""Arial""&12&BBeneficiary Report - " & Date & vbLf & "For Informational Purposes Only"

    'Add thick line below header
    reportSheet.Range("A1:E1").Borders(xlEdgeTop).LineStyle = xlContinuous
    reportSheet.Range("A1:E1").Borders(xlEdgeTop).Weight = xlMedium
    
    'Set the print range
    Set PrintRange = reportSheet.Range("A1")
    
    ShortLine
    NextLine
    
    'Put household name on report
    PrintRange.Value2 = household.NameOfHousehold
    PrintRange.Font.Bold = True
    PrintRange.Font.Size = 12
    reportSheet.Range(PrintRange, PrintRange.Offset(0, 2)).Merge across:=True
    
    NextLine
    NextLine
    
    'Get the sorted members
    Dim householdMembers As Dictionary
    Set householdMembers = household.SortedMembers

    'Declare a variable to keep track of if the first member has been added to the report
    Dim firstMemberAdded As Boolean

    'Put each member onto the report
    Dim member As Integer
    For member = 0 To householdMembers.count - 1
        'Add the member if they are not deceased and have accounts
        If householdMembers.Items(member).Active And householdMembers.Items(member).ActiveAccountsCount > 0 Then
            'Add the client's name unless it's the first member and it matches the household name
            Dim addName As Boolean
            addName = Not (member = 0 And householdMembers.Items(member).NameOfMember = household.NameOfHousehold)
            
            'Add the member to the report
            firstMemberAdded = AddMemberToReport(reportSheet, householdMembers.Items(member), addName, firstMemberAdded)
            
            If member < householdMembers.count - 1 Then
                'There are more members, see if they have active accounts
                If householdMembers.Items(member + 1).Active And householdMembers.Items(member + 1).ActiveAccountsCount > 0 Then
                    'More members will be added to the report. Get the current number of page breaks
                    Dim startPageBreaks As Integer
                    startPageBreaks = PageBreakTestStart
        
                    'Check for a page break
                    If Not PageBreakTestEnd(startPageBreaks, PrintRange) Then
                        'The members are on the same page. Put top border between members
                        reportSheet.Range(PrintRange, PrintRange.Offset(0, 4)).Borders(xlEdgeTop).LineStyle = xlContinuous
                        reportSheet.Range(PrintRange, PrintRange.Offset(0, 4)).Borders(xlEdgeTop).Weight = xlThin
            
                        'Put a double space between members
                        NextLine
                        NextLine
                    End If
                End If
            End If
        End If
    Next member
    
    'Set the formatting for the sheet
    SetReportFont reportSheet
    SetColumnWidths reportSheet
End Sub

Private Sub SetReportFont(sht As Worksheet)
    With sht.UsedRange.Font
        .name = "Arial"
        .Size = 11
    End With
End Sub

Private Sub SetColumnWidths(sht As Worksheet)
    With sht
        .Columns(1).ColumnWidth = 45
        .Columns(2).ColumnWidth = 11
        .Columns(3).ColumnWidth = 10
        .Columns(4).ColumnWidth = 15
        .Columns(5).ColumnWidth = 5.86
    End With
End Sub

Private Function AddMemberToReport(reportSheet As Worksheet, member As Variant, addName As Boolean, isFirstMemberAdded As Boolean) As Boolean
    If isFirstMemberAdded Then
        'Get the starting position in case a page break needs to be added
        Dim memberStartRange As Range
        Set memberStartRange = PrintRange
                
        'Get the current number of page breaks
        Dim memberStartPageBreaks As Integer
        memberStartPageBreaks = PageBreakTestStart
    End If
    
    If addName Then
        'Add the member's name
        PrintRange.Value2 = member.FName & " " & member.LName
        PrintRange.Font.Bold = True
        PrintRange.Font.Size = 12
        NextLine
    End If
    
    ShortLine
    NextLine
    
    'Get the sorted accounts
    Dim memberAccounts As Dictionary
    Set memberAccounts = member.SortedAccounts
    
    'Declare a variable to track if the first account has been added
    Dim isFirstAccountAdded As Boolean
    
    'Add the member's accounts
    Dim account As Integer
    For account = 0 To memberAccounts.count - 1
        Dim memberAccount As clsAccount
        Set memberAccount = memberAccounts.Items(account)
        If memberAccount.Active And memberAccount.Balance > 0 Then
            'Get the starting position in case a page break needs to be added
            Dim startRange As Range
            Set startRange = PrintRange
            
            'Get the current number of page breaks
            Dim startPageBreaks As Integer
            startPageBreaks = PageBreakTestStart
            
            'Add the account to the report
            AddAccountToReport memberAccount
            
            If isFirstMemberAdded And Not isFirstAccountAdded Then
                'This is the first account, check if there's been a page break after the member name
                PageBreakTestEnd memberStartPageBreaks, memberStartRange
            Else
                'This is after the first account, check if the account is split between pages
                PageBreakTestEnd startPageBreaks, startRange
            End If
        
            'Put a double space between accounts
            NextLine
            NextLine
            
            isFirstAccountAdded = True
        End If
    Next account
    
    'The member was added to the report; return true
    AddMemberToReport = True
End Function

Private Function PageBreakTestStart() As Integer
    'Get the current number of page breaks
    Dim startPageBreaks As Integer
    startPageBreaks = PrintRange.Worksheet.HPageBreaks.count
            
    'Put a test value into the PrintRange cell to see if it would cause a new page
    PrintRange.Value2 = 0
    
    'Return the starting number of page breaks
    PageBreakTestStart = startPageBreaks
    
    'Clear the test value
    PrintRange.Value2 = vbNullString
End Function

Private Function PageBreakTestEnd(startPageBreaks As Integer, startRange As Range) As Boolean
    'If the account is split between pages, put it all on the new page
    Dim endPageBreaks As Integer
    endPageBreaks = startRange.Worksheet.HPageBreaks.count
    If startPageBreaks <> endPageBreaks Then
        AddPageBreak startRange
        PageBreakTestEnd = True
    End If
End Function

Private Sub AddPageBreak(rng As Range)
    If Not rng Is Nothing Then
        'Add a page break before the account information
        rng.Worksheet.Rows(rng.Offset(-1, 0).Row).PageBreak = xlPageBreakManual
        
        'Put a border at the top between the header and the content
        rng.Worksheet.Range(rng, rng.Offset(-1, 4)).Borders(xlEdgeTop).LineStyle = xlContinuous
        rng.Worksheet.Range(rng, rng.Offset(-1, 4)).Borders(xlEdgeTop).Weight = xlMedium
    End If
End Sub

Private Function AddAccountToReport(account As clsAccount) As Boolean
    'Add account name
    With PrintRange
        .Value2 = ChrW(&H25BA) & account.NameOfAccount & AddAsOfDate(account)
        With .Characters(Len(account.NameOfAccount) + 2).Font
            .Italic = True
            .Size = 9
        End With
    End With
    
    NextLine
    ShortLine
    NextLine
    
    'Add custodian heading
    With PrintRange
        .Value2 = "Custodian"
        .HorizontalAlignment = xlLeft
        .Font.Underline = True
    End With
        
    'Add account type heading
    With PrintRange.Offset(0, 1)
        .Value2 = "Account Type"
        .HorizontalAlignment = xlLeft
        .Font.Underline = True
    End With
    
    NextLine
    
    'Add custodian and account type
    With PrintRange
        If account.custodian = "Aggregated" Or account.custodian = "Other" Then
            .Value2 = "Aggregated (Non-TD Ameritrade)"
        Else
            .Value2 = account.custodian
        End If
        
        .HorizontalAlignment = xlLeft
        .Offset(0, 1).Value2 = account.TypeOfAccount
        .Offset(0, 1).HorizontalAlignment = xlLeft
    End With
    
    NextLine
    ShortLine
    NextLine
    
    'Get the sorted accounts
    Dim accountBenes As Collection
    Set accountBenes = account.SortedBenes
    
    'Add beneficiaries
    If accountBenes.count = 0 Then
        'No beneficiaries
        PrintNoBeneficiaryInfo account
        NextLine
    Else
        'Add the beneficiary headings
        AddBeneHeadings

        NextLine
            
        Dim primaryTotal As Double
        Dim contingentTotal As Double
        Dim bene As Integer
        For bene = 1 To accountBenes.count
            'Get the beneficiary
            Dim accountBene As clsBeneficiary
            Set accountBene = accountBenes(bene)
            
            'Add the beneficiary's percentage to its respective total
            If accountBene.Level = "P" Then
                primaryTotal = primaryTotal + accountBene.Percent
            Else
                contingentTotal = contingentTotal + accountBene.Percent
            End If
            
            'Add the beneficiaries
            AddBeneToReport accountBene
            NextLine
        Next bene
        
        'Check if the beneficiary percentages don't equal 100%
        If Round(primaryTotal, 2) <> 100 Or (contingentTotal > 0 And Round(contingentTotal, 2) <> 100) Then
            'Throw an error
            Dim errorMessage As String
            errorMessage = "Beneficiary percentages don't equal 100% for " & account.NameOfAccount & ". Proceed with report?" & vbLf & vbLf
            
            Dim errorBene As Variant
            For Each errorBene In accountBenes
                If primaryTotal <> 100 And errorBene.Level = "P" Then
                    errorMessage = errorMessage & errorBene.NameOfBeneficiary & vbTab & errorBene.Level & vbTab & errorBene.Percent & vbLf
                ElseIf (contingentTotal > 0 And contingentTotal <> 100) And errorBene.Level = "C" Then
                    errorMessage = errorMessage & errorBene.NameOfBeneficiary & vbTab & errorBene.Level & vbTab & errorBene.Percent & vbLf
                End If
            Next errorBene
            
            If MsgBox(errorMessage, vbYesNo) = vbNo Then
                'Close the report without saving
                Dim reportBook As Workbook
                Set reportBook = PrintRange.Worksheet.Parent
                reportBook.Close savechanges:=False
                
                'End the macro
                Application.ScreenUpdating = True
                End
            End If
        End If
    End If
    
    'The account has been added to the report; return true
    AddAccountToReport = True
End Function

Private Function AddAsOfDate(account As clsAccount) As String
    If account.custodian <> "TD Ameritrade Institutional" Then
        If account.BenesUpdated > CDate(0) Then
            AddAsOfDate = " (As of " & Format(account.BenesUpdated, "m/d/yyyy") & ")"
        Else
            AddAsOfDate = vbNullString
        End If
    Else
        AddAsOfDate = vbNullString
    End If
End Function

Private Sub PrintNoBeneficiaryInfo(account As clsAccount)
    'Print the proper blurb if the account is labeled
    If Len(account.Tag) > 0 Then
        If account.Tag = "Associated" Then
            PrintRange.value = AssociatedBlurb
        ElseIf account.Tag = "WEC" Then
            PrintRange.value = WECBlurb
        ElseIf account.Tag = "Charitable" Then
            PrintRange.value = IneligibleBlurb
        ElseIf account.Tag = "Form" Then
            PrintRange.value = ""
        ElseIf account.Tag = "Online" Then
            PrintRange.value = ""
        ElseIf account.Tag = "Phone" Then
            PrintRange.value = ""
        End If
    Else
        'Print a blurb depending on the account's type
        PrintTypeSpecificInfo account.TypeOfAccount
    End If
    
    'Set the row height and merge the info across the report
    SetRowHeightAndMerge
End Sub

Private Sub PrintTypeSpecificInfo(accountType As String)
    'Add the appropriate verbiage based on the account type
    If UBound(Filter(TrustAccounts, accountType)) > -1 Then
        PrintRange.Value2 = "See most recent Trust documents for beneficiary information"
    ElseIf UBound(Filter(IneligibleAccounts, accountType)) > -1 Then
        If accountType = "Other" Then
            'Account fell in here because filtering the ineligible accounts for "Other" returns "Custodian Others"
            PrintRange.Value2 = "No beneficiaries found for this account"
        ElseIf accountType = "Custodian" Or accountType = "Custodian Others" Or accountType = "UGMA/UTMA" Then
            PrintRange.Value2 = IneligibleBlurb
        ElseIf accountType = "Estate" Then
            PrintRange.Value2 = IneligibleBlurb
        ElseIf accountType = "529 Plan" Then
            PrintRange.Value2 = IneligibleBlurb
        Else
            PrintRange.Value2 = IneligibleBlurb
        End If
    ElseIf accountType = "Individual" Or accountType = "Joint (JTWROS)" Then
        PrintRange.Value2 = "No beneficiaries on record; Transfer on Death (TOD) designation is available. Please give us a call if you would like to add beneficiaries."
    Else
        PrintRange.Value2 = "No beneficiaries found for this account"
    End If
End Sub

Private Sub SetRowHeightAndMerge()
    'Temporarily change column width to be the resultant merged columns' width
    PrintRange.ColumnWidth = 81
    
    'Wrap the text
    PrintRange.WrapText = True
    
    'Store the row height
    Dim rHeight As Integer
    rHeight = PrintRange.RowHeight
    
    'Reset the column width
    PrintRange.ColumnWidth = 45
    
    'Merge the cell across the report
    PrintRange.Worksheet.Range(PrintRange, PrintRange.Offset(0, 3)).Merge across:=True
    
    'Set the row height
    PrintRange.RowHeight = rHeight
End Sub

Private Sub AddBeneHeadings()
    With PrintRange
        .Value2 = "Beneficiary Name"
        .IndentLevel = 2
        .Font.Underline = True
        .Offset(0, 1).Value2 = "Bene Type"
        .Offset(0, 1).Font.Underline = True
        .Offset(0, 1).HorizontalAlignment = xlLeft
        .Offset(0, 2).Value2 = "Percent"
        .Offset(0, 2).Font.Underline = True
        .Offset(0, 2).HorizontalAlignment = xlCenter
    End With
End Sub

Private Sub AddBeneToReport(bene As clsBeneficiary)
    'Add the beneficiary
    With PrintRange
        'Name
        .Value2 = UCase(bene.NameOfBeneficiary)
        .IndentLevel = 2
        .WrapText = True
        
        'Type
        With .Offset(0, 1)
            If bene.Level = "P" Then
                .Value2 = "Primary"
            ElseIf bene.Level = "C" Then
                .Value2 = "Contingent"
            Else
                .Value2 = bene.Level
            End If
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlCenter
        End With
        
        'Percent
        With .Offset(0, 2)
            .Value2 = bene.Percent & "%"
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
    End With
End Sub

Private Sub NextLine()
    Set PrintRange = PrintRange.Offset(1, 0)
End Sub

Private Sub ShortLine()
    PrintRange.RowHeight = 4.5
End Sub


