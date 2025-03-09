Attribute VB_Name = "Module2"
Sub ImportanceTracker()
    Dim Task As String
    Dim DaysLeft As Integer
    Dim InputImportance As Integer
    Dim InputImportanceString As String
    Dim HoursInDay As Variant
    Dim HoursPerDay As Double
    Dim HoursRequired As Double
    Dim DaysImportance As Integer
    Dim GeneralImportance As Integer
    Dim ws As Worksheet
    Dim os As Worksheet
    Dim TaskAmount As Integer
    Dim i As Integer
    Dim j As Integer
    Dim tempImportance As Integer
    Dim tempTask As String
    Dim tempHours As Double
    Dim TotalHours As Variant
    Dim TaskCount As Integer
    Dim Test As Double
    Dim endDate As Date
    Dim Program As String
    Dim Year As Variant
    Dim DegreeLevel As String
    Dim ps As Worksheet
    Dim tempDueDate As Date

    Set ws = ThisWorkbook.Sheets("Task Tracking Sheet")
    Set os = ThisWorkbook.Sheets("Data Processing")
    Set ps = ThisWorkbook.Sheets("Personal Profile")
    TaskAmount = Application.WorksheetFunction.CountA(ws.Range("B5:B100"))
    
    Dim ImportanceList() As Integer
    Dim TaskList() As String
    Dim HoursList() As Double
    Dim DueDateList() As Date

    ReDim ImportanceList(TaskAmount - 1)
    ReDim TaskList(TaskAmount - 1)
    ReDim HoursList(TaskAmount - 1)
    ReDim DueDateList(TaskAmount - 1)
    
    ' TODO:Assign hoursInDay based oninputed program details
    
    Program = ps.Cells(5, 4).Value
    Year = ps.Cells(5, 3).Value
    DegreeLevel = ps.Cells(5, 2).Value
    
    HoursInDay = 2
    
    If Program = "Mathematics" Or Program = "Engineering" Or Program = "Science" Then
        HoursInDay = HoursInDay * 2
    End If
    Select Case DegreeLevel
        Case "Undergraduate": HoursInDay = HoursInDay * 1
        Case "Diploma": HoursInDay = HoursInDay * 1.2
        Case "Masters": HoursInDay = HoursInDay * 1.4
        Case "PHD": HoursInDay = HoursInDay * 1.5
        Case "Other": HoursInDay = HoursInDay * 1
    End Select
    
    Select Case Year
        Case 1: HoursInDay = HoursInDay * 1
        Case 2: HoursInDay = HoursInDay * 1.2
        Case 3: HoursInDay = HoursInDay * 1.3
        Case 4: HoursInDay = HoursInDay * 1.4
        Case Else: HoursInDay = HoursInDay * 1.5
    End Select
    
    

    ' Determine importance for each task
    For i = 5 To TaskAmount + 4
        InputImportanceString = ws.Cells(i, 7).Value
        
        Select Case InputImportanceString
            Case "High": InputImportance = 3
            Case "Medium": InputImportance = 2
            Case "Low": InputImportance = 1
        End Select
    
            
        endDate = ws.Cells(i, 6)
        DaysLeft = endDate - Date
        
        DueDateList(i - 5) = endDate
    
        Select Case DaysLeft
            Case Is > 7: DaysImportance = 0
            Case 4 To 7: DaysImportance = 2
            Case 3: DaysImportance = 4
            Case 2: DaysImportance = 6
            Case 1: DaysImportance = 11
            Case 0: DaysImportance = 12
            Case Else: DaysImportance = -1000
        End Select
        If DaysLeft = 0 Then
            DaysLeft = 1
        End If
        
        If DaysLeft > 0 Then
            GeneralImportance = DaysImportance + InputImportance
            TaskList(i - 5) = ws.Cells(i, 2).Value
            ImportanceList(i - 5) = GeneralImportance
            HoursRequired = ws.Cells(i, 9).Value
            HoursList(i - 5) = HoursRequired / DaysLeft
        End If
    Next i

    ' Bubble sort the three lists
    For i = 0 To TaskAmount - 2
        For j = 0 To TaskAmount - i - 2
            If ImportanceList(j) < ImportanceList(j + 1) Then
               
                tempImportance = ImportanceList(j)
                ImportanceList(j) = ImportanceList(j + 1)
                ImportanceList(j + 1) = tempImportance

              
                tempTask = TaskList(j)
                TaskList(j) = TaskList(j + 1)
                TaskList(j + 1) = tempTask

        
                tempHours = HoursList(j)
                HoursList(j) = HoursList(j + 1)
                HoursList(j + 1) = tempHours
                
                tempDueDate = DueDateList(j)
                DueDateList(j) = DueDateList(j + 1)
                DueDateList(j + 1) = tempDueDate
            End If
        Next j
    Next i

    ' Tally up tasks that fit within a day
    TotalHours = 0#
    TaskCount = 0
    
    Dim TestValue As Variant
    
    For i = 0 To TaskAmount - 1
        If TotalHours + HoursList(i) <= HoursInDay Then
            TotalHours = TotalHours + HoursList(i)
            TaskCount = TaskCount + 1
        Else
            HoursList(i) = HoursInDay - TotalHours
            TaskCount = TaskCount + 1
            Exit For
        End If
    Next i
    ' Output result
    os.Cells.Range("A3:C50").ClearContents

    For i = 0 To TaskCount - 1
        os.Cells(i + 3, 1).Value = TaskList(i)
        os.Cells(i + 3, 2).Value = Round(HoursList(i), 1)
        os.Cells(i + 3, 3).Value = DueDateList(i)
    Next i
End Sub

