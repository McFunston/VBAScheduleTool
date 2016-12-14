Sub Location2ScheduleFinder_Click()
    Dim FileSelect As Variant
    'hold in memory
    Application.ScreenUpdating = False
    'locate the file path
    FileSelect = Application.GetOpenFilename(fileFilter:="Excel Files,*.xl*", _
    MultiSelect:=False)
    'check if a file is selected
    If FileSelect = False Then
        MsgBox "Select the file name"
        Exit Sub
    End If
    Sheet1.Range("A7").Value = FileSelect
    Call LoadList(Sheet1.Range("A7").Value, "Location2")
End Sub

Sub Location1ScheduleFinder_Click()
    Dim FileSelect As Variant
    'hold in memory
    Application.ScreenUpdating = False
    'locate the file path
    FileSelect = Application.GetOpenFilename(fileFilter:="Excel Files,*.xl*", _
    MultiSelect:=False)
    'check if a file is selected
    If FileSelect = False Then
        MsgBox "Select the file name"
        Exit Sub
    End If
    Sheet1.Range("A12").Value = FileSelect
    Call LoadList(Sheet1.Range("A12").Value, "Location1")
End Sub

Sub PrintflowScheduleFinder_Click()
    Dim FileSelect As Variant
    'hold in memory
    Application.ScreenUpdating = False
    'locate the file path
    FileSelect = Application.GetOpenFilename(fileFilter:="Excel Files,*.xl*", _
    MultiSelect:=False)
    'check if a file is selected
    If FileSelect = False Then
        MsgBox "Select the file name"
        Exit Sub
    End If
    Sheet1.Range("A17").Value = FileSelect
    Call LoadList(Sheet1.Range("A17").Value, "PrintflowToDo")
End Sub

Sub PrintflowDone_Click()
    Dim FileSelect As Variant
    'hold in memory
    Application.ScreenUpdating = False
    'locate the file path
    FileSelect = Application.GetOpenFilename(fileFilter:="Excel Files,*.xl*", _
    MultiSelect:=False)
    'check if a file is selected
    If FileSelect = False Then
        MsgBox "Select the file name"
        Exit Sub
    End If
    Sheet1.Range("A22").Value = FileSelect
End Sub

Sub Clear_Click()
    Range("C2:M1000").ClearContents
End Sub
Sub LoadDocketNumbers(FileName As String, ColNum As String, ByRef JobsNumberList As Collection, ColName As String, ByRef JobsNamesList As Collection)
    
    Dim JobNumber As String
    'Dim JobsNumberList As New Collection
    Dim NumberOfRows As Integer
    Dim i As Integer
    Dim JobName As String
    Application.ScreenUpdating = False
    Worksheets(FileName).Activate
    NumberOfRows = ActiveSheet.Range("A" & Rows.Count).End(xlUp).Row 'Getting the row count of the source
    For i = 2 To NumberOfRows
    JobNumber = ActiveSheet.Range(ColNum & i)
    JobName = ActiveSheet.Range(ColName & i)
    'If JobsNumberList.Count = 0 Then
    '    JobsNumberList.Add (JobNumber)
    If CollectionContains(JobsNumberList, JobNumber) Then
    ElseIf IsNumeric(JobNumber) Then
    'MsgBox JobName
        JobsNumberList.Add (JobNumber)
        JobsNamesList.Add (JobName)
    End If
    Next i
    'MsgBox JobsNumberList.Count
    Worksheets("Sheet1").Activate
    
End Sub

Function LoadCSR(JobNumber As String) As String

    Dim NumberOfRows As Integer
    Dim CSR As String
    Worksheets("PrintflowToDo").Activate
    NumberOfRows = ActiveSheet.Range("A" & Rows.Count).End(xlUp).Row 'Getting the row count of the source
    For i = 2 To NumberOfRows
        If ActiveSheet.Range("D" & i) = JobNumber Then
            CSR = ActiveSheet.Range("I" & i).Value
            Exit For
        End If
    Next i
    LoadCSR = CSR
    Worksheets("Sheet1").Activate
End Function

Function GetLocation2Time(JobNumber) As Variant
    Dim FirstOnPress As String
    Dim DateOnPress As Date
    Dim NumberOfRows As Integer
    FirstOnPress = "2"
    Worksheets("Location2").Activate
    
        NumberOfRows = ActiveSheet.Range("A" & Rows.Count).End(xlUp).Row 'Getting the row count of the source
        For i = 2 To NumberOfRows
        FirstOnPress = "2"
        If ActiveSheet.Range("B" & i).Value = JobNumber Then
            FirstOnPress = ActiveSheet.Range("A" & i).Value
            DateOnPress = ConvertDate(FirstOnPress)
            Exit For
        End If
        DateOnPress = CDate(FirstOnPress)
        Next i
    
    GetLocation2Time = DateOnPress
    Worksheets("Sheet1").Activate
End Function

Function GetLocation1Time(JobNumber) As Variant
    Dim FirstOnPress As String
    Dim DateOnPress As Date
    Dim NumberOfRows As Integer
    FirstOnPress = "2"
    Worksheets("Location1").Activate
    
        NumberOfRows = ActiveSheet.Range("A" & Rows.Count).End(xlUp).Row 'Getting the row count of the source
        For i = 2 To NumberOfRows
        FirstOnPress = "2"
        If ActiveSheet.Range("B" & i).Value = JobNumber Then
            FirstOnPress = ActiveSheet.Range("A" & i).Value
            DateOnPress = ConvertDate(FirstOnPress)
            Exit For
        End If
        DateOnPress = CDate(FirstOnPress)
        Next i
    GetLocation1Time = DateOnPress
    
    
    Worksheets("Sheet1").Activate
End Function

Sub GetStatus(JobNumber As String, ByRef FilesIn As String, ByRef ProofOut As String, ByRef ProofIn As String)
       
    Dim NumberOfRows As Integer
    Dim CostCenter As String
    Worksheets("PrintflowToDo").Activate
    NumberOfRows = ActiveSheet.Range("A" & Rows.Count).End(xlUp).Row 'Getting the row count of the source
  
    For i = 2 To NumberOfRows
        CostCenter = ActiveSheet.Range("C" & i)
        If ActiveSheet.Range("D" & i).Value = JobNumber Then
            If CostCenter = "Location2 Printing-Files In" Or CostCenter = "Hunt Club-Files In" Then
                FilesIn = FixPrintflowDate(ActiveSheet.Range("A" & i))
            End If
            If CostCenter = "Location2 Printing-Proof Out" Or CostCenter = "Hunt Club-Proof Out" Then
                ProofOut = FixPrintflowDate(ActiveSheet.Range("A" & i))
            End If
            If CostCenter = "Location2 Printing-Proof In" Or CostCenter = "Hunt Club-Proof In" Then
                
                ProofIn = FixPrintflowDate(ActiveSheet.Range("A" & i))
            End If
        End If
    Next i
    Worksheets("Sheet1").Activate
End Sub

Function CollectionContains(myCol As Collection, checkVal As Variant) As Boolean
    On Error Resume Next
    CollectionContains = False
    Dim it As Variant
    For Each it In myCol
        If it = checkVal Then
            CollectionContains = True
            Exit Function
        End If
    Next
End Function

Sub GetSchedule_Click()
    Dim PrintflowFile As String
    Dim Location2File As String
    Dim Location1File As String
    Dim JobsNumberList As New Collection
    Dim JobsNamesList As New Collection
    Dim FilesIn As String
    Dim ProofOut As String
    Dim ProofIn As String

    Dim JobsList As New Collection
    Dim ii As Integer
    
    Range("C2:M1000").Range("C2:M1000").ClearContents
    
    ii = 2
    PrintflowFile = Sheet1.Range("A17").Value
    Location2File = Sheet1.Range("A7").Value
    Location1File = Sheet1.Range("A12").Value
    Call LoadDocketNumbers("PrintflowToDo", "D", JobsNumberList, "G", JobsNamesList)
    Call LoadDocketNumbers("Location2", "B", JobsNumberList, "C", JobsNamesList)
    Call LoadDocketNumbers("Location1", "B", JobsNumberList, "C", JobsNamesList)
    For Each JN In JobsNumberList
        FilesIn = "2"
        ProofOut = "2"
        ProofIn = "1"
        Sheet1.Range("C" & ii) = JN
        Sheet1.Range("E" & ii) = LoadCSR(CStr(JN))
        Sheet1.Range("F" & ii) = GetLocation1Time(CStr(JN))
        Sheet1.Range("G" & ii) = GetLocation2Time(CStr(JN))
        Call GetStatus(CStr(JN), FilesIn, ProofOut, ProofIn)
        Sheet1.Range("H" & ii) = CDate(FilesIn)
        Sheet1.Range("I" & ii) = CDate(ProofOut)
        Sheet1.Range("J" & ii) = CDate(ProofIn)
        ii = ii + 1
    Next JN
    ii = 2
    For Each Job In JobsNamesList
        Sheet1.Range("D" & ii) = Job
        ii = ii + 1
    Next Job
    
End Sub

Sub LoadPrintflowToDo()
    Dim PrintflowFile As String
    PrintflowFile = Sheet1.Range("A17").Value
    Dim wbk1 As Workbook
    Set wbk1 = ActiveWorkbook
    Dim PrintflowWB As Workbook
    Set PrintflowWB = Workbooks.Open(PrintflowFile)
    PrintflowWB.Sheets("Sheet0").Copy After:=wbk1.Sheets(1)
    PrintflowWB.Close False
    Sheets(2).Name = "PrintflowToDo"
    Worksheets("Sheet1").Activate
End Sub

Sub LoadLocation2()
    Dim Location2File As String
    Location2File = Sheet1.Range("A7").Value
    Dim wbk1 As Workbook
    Set wbk1 = ActiveWorkbook
    Dim Location2WB As Workbook
    Set Location2WB = Workbooks.Open(Location2File)
    Location2WB.Sheets(1).Copy After:=wbk1.Sheets(1)
    Location2WB.Close False
    Sheets(2).Name = "Location2"
    Worksheets("Sheet1").Activate
End Sub

Sub LoadLocation1()
    Dim Location1File As String
    Location1File = Sheet1.Range("A12").Value
    Dim wbk1 As Workbook
    Set wbk1 = ActiveWorkbook
    Dim Location1WB As Workbook
    Set Location1WB = Workbooks.Open(Location1File)
    Location1WB.Sheets(1).Copy After:=wbk1.Sheets(1)
    Location1WB.Close False
    Sheets(2).Name = "Location1"
    Worksheets("Sheet1").Activate
End Sub
Public Function ConvertDate(suppliedDate As String) As Variant
If suppliedDate <> "Unknown/NA" Then
    Dim YY As String
    Dim MM As String
    Dim DD As String
    Dim FormatedDate
    YY = Left(suppliedDate, 2)
    MM = Mid(suppliedDate, 4, 2)
    DD = Mid(suppliedDate, 7, 2)
    FormatedDate = DD & "/" & MM & "/" & "20" & YY & Right(suppliedDate, 6) & ":00"
    'MsgBox (FormatedDate)
    Dim ProperDate
    ProperDate = FormatedDate
    ConvertDate = CDate(ProperDate)
End If
ConvertDate = suppliedDate
End Function
Public Function FixPrintflowDate(suppliedDate As String) As String
Dim MM As String
Dim DD As String
Dim YY As String
Dim FixedDate As String
DD = Left(suppliedDate, 2)
MM = Mid(suppliedDate, 4, 2)
FixedDate = MM & "/" & DD & "/" & Right(suppliedDate, 16)
FixPrintflowDate = FixedDate
End Function
Public Function LoadList(path As String, tabName As String)
Application.DisplayAlerts = False
Dim wbk1 As Workbook
Set wbk1 = ActiveWorkbook
If sheetExists(tabName) Then
    wbk1.Sheets(tabName).Delete
End If
Dim NewWorkbook As Workbook
Set NewWorkbook = Workbooks.Open(path)
NewWorkbook.Sheets(1).Copy After:=wbk1.Sheets(1)
NewWorkbook.Close False
Sheets(2).Name = tabName
Worksheets("Sheet1").Activate
Application.DisplayAlerts = True
End Function

Function sheetExists(sheetToFind As String) As Boolean
    sheetExists = False
    For Each Sheet In Worksheets
        If sheetToFind = Sheet.Name Then
            sheetExists = True
            Exit Function
        End If
    Next Sheet
End Function

Public Function Contains(strBaseString As String, strSearchTerm As String) As Boolean
'Purpose: Returns TRUE if one string exists within another
'On Error GoTo ErrorMessage
    Contains = InStr(strBaseString, strSearchTerm)
Exit Function
'ErrorMessage:
'MsgBox "The database has generated an error. Please contact the database administrator, quoting the following error message: '" & Err.Description & "'", vbCritical, "Database Error"
End
End Function