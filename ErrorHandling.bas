Attribute VB_Name = "ErrorHandling"
Option Explicit
Private Const errorFileName As String = "Z:\YungwirthSteve\Beneficiary Report\Assets\errors.txt"

Public Sub LogErrorToFile(errorString As String)
    'Open a text file to write errors to
    Dim fs As FileSystemObject
    Set fs = New FileSystemObject
    Dim errorFile As TextStream
    Set errorFile = fs.OpenTextFile(errorFileName, ForAppending, True)
    
    'Write the error string
    Dim ele As Integer
    errorFile.WriteLine errorString

    'Close the error file
    errorFile.Close
End Sub

Public Function AccountsNotInMSError(missingAccounts() As String) As String
    Dim missingAccount As Integer
    For missingAccount = LBound(missingAccounts) To UBound(missingAccounts)
        AccountsNotInMSError = AccountsNotInMSError & vbNewLine & missingAccounts(missingAccount) & " is in TD, but not in Morningstar"
    Next missingAccount
End Function
