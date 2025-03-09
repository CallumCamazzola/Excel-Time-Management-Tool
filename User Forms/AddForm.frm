VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AddForm 
   Caption         =   "UserForm1"
   ClientHeight    =   5870
   ClientLeft      =   100
   ClientTop       =   420
   ClientWidth     =   5520
   OleObjectBlob   =   "AddForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "AddForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CommandButton1_Click()
    Dim startMonth As Integer, startDay As Integer, startYear As Integer
    Dim endMonth As Integer, endDay As Integer, endYear As Integer
    Dim startDate As Date, endDate As Date
    
    ' Start Date Validation
    startMonth = Val(TextBox3.Value)
    startDay = Val(TextBox4.Value)
    startYear = Val(TextBox5.Value)
    
    ' End Date Validation
    endMonth = Val(TextBox8.Value)
    endDay = Val(TextBox7.Value)
    endYear = Val(TextBox6.Value)

    ' Validate Start Date
    If startMonth < 1 Or startMonth > 12 Then
        MsgBox "Please enter a valid month (1 to 12) for the start date.", vbCritical
        Exit Sub
    End If
    If startYear > Year(Date) Then
        MsgBox "Please enter a valid year (before or equal to the current year) for the start date.", vbCritical
        Exit Sub
    End If
    
    Select Case startMonth
        Case 1, 3, 5, 7, 8, 10, 12 ' Months with 31 days
            If startDay < 1 Or startDay > 31 Then
                MsgBox "Please enter a valid day (1 to 31) for the start date.", vbCritical
                Exit Sub
            End If
        Case 4, 6, 9, 11 ' Months with 30 days
            If startDay < 1 Or startDay > 30 Then
                MsgBox "Please enter a valid day (1 to 30) for the start date.", vbCritical
                Exit Sub
            End If
        Case 2 ' February
            If (startYear Mod 4 = 0 And (startYear Mod 100 <> 0 Or startYear Mod 400 = 0)) Then
                If startDay < 1 Or startDay > 29 Then
                    MsgBox "Please enter a valid day (1 to 29) for February in a leap year.", vbCritical
                    Exit Sub
                End If
            Else
                If startDay < 1 Or startDay > 28 Then
                    MsgBox "Please enter a valid day (1 to 28) for February in a non-leap year.", vbCritical
                    Exit Sub
                End If
            End If
    End Select
    
    ' Validate End Date
    If endMonth < 1 Or endMonth > 12 Then
        MsgBox "Please enter a valid month (1 to 12) for the end date.", vbCritical
        Exit Sub
    End If
    If endYear > Year(Date) Then
        MsgBox "Please enter a valid year (before or equal to the current year) for the end date.", vbCritical
        Exit Sub
    End If
    
    Select Case endMonth
        Case 1, 3, 5, 7, 8, 10, 12 ' Months with 31 days
            If endDay < 1 Or endDay > 31 Then
                MsgBox "Please enter a valid day (1 to 31) for the end date.", vbCritical
                Exit Sub
            End If
        Case 4, 6, 9, 11 ' Months with 30 days
            If endDay < 1 Or endDay > 30 Then
                MsgBox "Please enter a valid day (1 to 30) for the end date.", vbCritical
                Exit Sub
            End If
        Case 2 ' February
            If (endYear Mod 4 = 0 And (endYear Mod 100 <> 0 Or endYear Mod 400 = 0)) Then
                If endDay < 1 Or endDay > 29 Then
                    MsgBox "Please enter a valid day (1 to 29) for February in a leap year.", vbCritical
                    Exit Sub
                End If
            Else
                If endDay < 1 Or endDay > 28 Then
                    MsgBox "Please enter a valid day (1 to 28) for February in a non-leap year.", vbCritical
                    Exit Sub
                End If
            End If
    End Select
    
    On Error GoTo ErrorHandler
    startDate = DateSerial(startYear, startMonth, startDay)
    endDate = DateSerial(endYear, endMonth, endDay)
    On Error GoTo 0
    
    If startDate > endDate Then
        MsgBox "Start date must be before end date.", vbCritical
        Exit Sub
    End If

    Set WB = ThisWorkbook
    Set ws = WB.Worksheets("Task Tracking Sheet")

    Dim intRow As Integer
    intRow = 5
    Do While ws.Cells(intRow, "B") <> ""
        intRow = intRow + 1
    Loop
    
    ws.Cells(intRow, "B") = TextBox1.Value
    ws.Cells(intRow, "C") = TextBox2.Value
    ws.Cells(intRow, "D") = ComboBox3.Value
    ws.Cells(intRow, "E") = startDate
    ws.Cells(intRow, "E").NumberFormat = "yyyy-mm-dd;@"
    ws.Cells(intRow, "F") = endDate
    ws.Cells(intRow, "F").NumberFormat = "yyyy-mm-dd;@"
    ws.Cells(intRow, "G") = ComboBox1.Value
    ws.Cells(intRow, "I") = ComboBox2.Value
    ws.Cells(intRow, "H") = "0"
    ws.Cells(intRow, "H").NumberFormat = "0%"

    Sheets("Graphical Output").PivotTables("PivotTable1").PivotCache.Refresh
    Sheets("Graphical Output").PivotTables("PivotTable3").PivotCache.Refresh

    Exit Sub

ErrorHandler:
    MsgBox "Invalid date input. Please check the entered date values.", vbCritical
End Sub


Private Sub UserForm_Initialize()
    
    Set WB = ThisWorkbook
    Set ws = WB.Worksheets("Task Tracking Sheet")
    
    
    ComboBox1.AddItem ("High")
    ComboBox1.AddItem ("Medium")
    ComboBox1.AddItem ("Low")
    
  
    ComboBox2.AddItem ("1")
    ComboBox2.AddItem ("2")
    ComboBox2.AddItem ("3")
    ComboBox2.AddItem ("4")
    ComboBox2.AddItem ("5")
    ComboBox2.AddItem ("6")
    ComboBox2.AddItem ("7")
    ComboBox2.AddItem ("8")
    ComboBox2.AddItem ("9")
    ComboBox2.AddItem ("10")
    ComboBox2.AddItem ("11")
    ComboBox2.AddItem ("12")
    ComboBox2.AddItem ("13")
    ComboBox2.AddItem ("14")
    ComboBox2.AddItem ("15+")
    
    'Intialize Type Level
    ComboBox3.AddItem ("Project")
    ComboBox3.AddItem ("Assignment")
    ComboBox3.AddItem ("Test")
    ComboBox3.AddItem ("Exam")
    ComboBox3.AddItem ("Quiz")
    ComboBox3.AddItem ("Lab")
    ComboBox3.AddItem ("Report")
    ComboBox3.AddItem ("Essay")
    ComboBox3.AddItem ("Other")
End Sub
