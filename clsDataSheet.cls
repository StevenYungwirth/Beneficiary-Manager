VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDataSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_data() As Variant

Public Property Get data() As Variant
    data = m_data
End Property

Public Sub ClassBuilder(path As String, worksheetName As String, requiredHeaders As Variant, sortColumn As String)
    'Get the worksheet via the given filepath's workbook and the given worksheet name
    Dim dataSheet As Worksheet
    If Len(path) = 0 Then
        'No filepath was given, by default the sheet should be on the current workbook
        Set dataSheet = ThisWorkbook.Worksheets(worksheetName)
    Else
        'Turn off screen updating
        On Error GoTo BackOn
        Application.ScreenUpdating = False
        
        'Open the workbook with the given file name
        Dim dataSheetBook As Workbook
        Set dataSheetBook = Workbooks.Open(path)
        Set dataSheet = dataSheetBook.Worksheets(1)
    End If
    
    'Sort the sheet by the given column header
    If sortColumn <> vbNullString Then
        SortSheet dataSheet, sortColumn
    End If
    
    'Get the data from the worksheet
    m_data = dataSheet.UsedRange.Value2
    
    'Make sure all the needed headers are present
    CheckDataSheetHeaders requiredHeaders, dataSheet.name
    
    'If a workbook was opened, close it
    If Len(path) > 0 Then
        dataSheetBook.Close savechanges:=False
        Application.ScreenUpdating = True
    End If
    
    On Error GoTo 0
    Exit Sub

BackOn:
    Application.ScreenUpdating = True
End Sub

Private Sub SortSheet(dataSheet As Worksheet, columnName As String)
    'Get the column to sort by
    Dim sortCol As Range
    Set sortCol = dataSheet.Cells(1, GetHeaderColumn(dataSheet.Rows(1), columnName))
    
    'Sort the worksheet
    dataSheet.UsedRange.sort key1:=sortCol, Header:=xlYes
End Sub

Private Function GetHeaderColumn(headerLine As Range, Header As String) As Integer
    GetHeaderColumn = headerLine.Find(Header, lookat:=xlWhole).Column
End Function

Public Sub FillWorksheet(sht As Worksheet)
    'Clear the worksheet
    sht.UsedRange.Value2 = vbNullString
    
    'Fill the worksheet with every datapoint in the data
    Dim rw As Integer, col As Integer
    For rw = 1 To UBound(m_data, 1)
        For col = 1 To UBound(m_data, 2)
            sht.Cells(rw, col).Value2 = m_data(rw, col)
        Next col
    Next rw
End Sub

Public Function GetData(headerName As String, datapoint As Integer) As String
    'Return the nth data point of a given header
    Dim dataFound As Boolean
    Dim col As Integer
    col = 1
    Do While Not dataFound And col <= UBound(m_data, 2)
        'Look for the inputted header
        If Trim(m_data(1, col)) = headerName Then
            'Get the data. Datapoint is +1 since the first row is the header
            GetData = m_data(datapoint + 1, col)
            dataFound = True
        End If
        
        'Look for the header in the next column
        col = col + 1
    Loop
End Function

Public Sub SetData(newData As Variant, headerName As String, datapoint As Integer)
    'Set the nth data point of a given header
    Dim dataFound As Boolean
    Dim col As Integer
    col = 1
    Do While Not dataFound And col <= UBound(m_data, 2)
        'Look for the inputted header
        If Trim(m_data(1, col)) = headerName Then
            If datapoint <> -1 Then
                'Set the data. Datapoint is +1 since the first row is the header
                m_data(datapoint + 1, col) = newData
                dataFound = True
            Else
                'Put the data into the first empty datapoint
                Dim firstEmpty As Integer
                firstEmpty = LBound(m_data, 1)
                Do While Not dataFound And firstEmpty <= UBound(m_data, 1)
                    If m_data(firstEmpty, col) = vbNullString Then
                        m_data(firstEmpty, col) = newData
                        dataFound = True
                    End If
                    
                    'See if the next row is empty
                    firstEmpty = firstEmpty + 1
                Loop
            End If
        End If
        
        'Look for the header in the next column
        col = col + 1
    Loop
End Sub

Public Function FirstEmptyRow() As Integer
    'Find the first datapoint that doesn't have data
End Function

Private Sub CheckDataSheetHeaders(requiredHeaders As Variant, worksheetName As String)
    'Set the array to be the values of the first row
    Dim headers() As String
    ReDim headers(1 To UBound(m_data, 2)) As String
    Dim col As Integer
    For col = 1 To UBound(data, 2)
        headers(col) = Trim(m_data(1, col))
    Next col
    
    'Check to make sure these column headers can be found. If one's not available, throw an error
    Dim i As Integer
    For i = 0 To UBound(requiredHeaders)
        'Check if the header is in the list of headers
        If Not DoesHeaderExist(headers, requiredHeaders(i)) Then
            'The needed column wasn't found
            MsgBox requiredHeaders(i) & " not found on " & worksheetName & ". Macro has been halted."
            End
        End If
    Next i
End Sub

Private Function DoesHeaderExist(headerArr As Variant, Header As Variant) As Boolean
    'If the header doesn't exist, GetIndexOf will return -1
    If GetIndexOf(headerArr, Header) = -1 Then
        DoesHeaderExist = False
    Else
        DoesHeaderExist = True
    End If
End Function

Private Function GetIndexOf(arr As Variant, str As Variant) As Integer
    'Return the index of an array's element, or -1 in case the element can't be found
    Dim i As Integer
    GetIndexOf = -1
    i = 1
    Do While GetIndexOf = -1 And i <= UBound(arr)
        If arr(i) = str Then
            GetIndexOf = i
        End If
        i = i + 1
    Loop
End Function
