VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} TaskFilterInput 
   Caption         =   "UserForm2"
   ClientHeight    =   4860
   ClientLeft      =   100
   ClientTop       =   460
   ClientWidth     =   3700
   OleObjectBlob   =   "Task Filter Input.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "TaskFilterInput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
    Dim startDate As Date
    Dim endDate As Date
    Dim wsFilter As Worksheet
    Dim wsTracking As Worksheet
    Dim lastRow As Long
    Dim outputRow As Long
    Dim taskRow As Range

    startDate = DateSerial(StartY.Value, StartM.Value, StartD.Value)
    endDate = DateSerial(EndY.Value, EndM.Value, EndD.Value)

    Set wsFilter = ThisWorkbook.Sheets("Task Filter")
    Set wsTracking = ThisWorkbook.Sheets("Task Tracking Sheet")

    If Application.WorksheetFunction.CountA(wsFilter.Range("G5:M" & wsFilter.Rows.Count)) > 0 Then
        MsgBox "Please clear the Task Filter sheet before running this filter.", vbExclamation, "Sheet Not Empty"
        Exit Sub
    End If

    wsFilter.Range("A4").Value = startDate
    wsFilter.Range("A6").Value = endDate

    lastRow = wsTracking.Cells(wsTracking.Rows.Count, "B").End(xlUp).Row
    outputRow = 5

    For Each taskRow In wsTracking.Range("B5:B" & lastRow).Rows
        If IsDate(wsTracking.Cells(taskRow.Row, "E").Value) And IsDate(wsTracking.Cells(taskRow.Row, "F").Value) Then
            If wsTracking.Cells(taskRow.Row, "E").Value >= startDate And wsTracking.Cells(taskRow.Row, "F").Value <= endDate Then
                wsFilter.Cells(outputRow, "G").Value = wsTracking.Cells(taskRow.Row, "B").Value
                wsFilter.Cells(outputRow, "H").Value = wsTracking.Cells(taskRow.Row, "C").Value
                wsFilter.Cells(outputRow, "I").Value = wsTracking.Cells(taskRow.Row, "D").Value
                wsFilter.Cells(outputRow, "J").Value = wsTracking.Cells(taskRow.Row, "E").Value
                wsFilter.Cells(outputRow, "K").Value = wsTracking.Cells(taskRow.Row, "F").Value
                wsFilter.Cells(outputRow, "L").Value = wsTracking.Cells(taskRow.Row, "G").Value
                wsFilter.Cells(outputRow, "M").Value = wsTracking.Cells(taskRow.Row, "H").Value
                outputRow = outputRow + 1
            End If
        End If
    Next taskRow

    wsFilter.Columns("M").NumberFormat = "0%"
End Sub


