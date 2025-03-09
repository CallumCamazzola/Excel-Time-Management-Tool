Attribute VB_Name = "Module4"
Sub Graphic3_Click()
    
    Application.ScreenUpdating = False
    
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Calendar Breakdown")
    
    
    Dim taskSheet As Worksheet
    Set taskSheet = ThisWorkbook.Sheets("Task Tracking Sheet")
    
   
    Dim dataProcessingSheet As Worksheet
    Set dataProcessingSheet = ThisWorkbook.Sheets("Data Processing")
    
   
    Dim dateCell As Range
    Dim taskStartDate As Date
    Dim taskEndDate As Date
    Dim taskName As String
    Dim taskHours As Double
    Dim taskRow As Long
    Dim commentText As String
    
    ' Loop over each date cell in B8:H8
    For Each dateCell In ws.Range("B8:H8")
        
        Dim currentDate As Date
        currentDate = ws.Cells(4, dateCell.Column).Value
        
        
        commentText = "Tasks for today: "
        
        
        For taskRow = 2 To taskSheet.Cells(taskSheet.Rows.Count, "E").End(xlUp).Row
           
            If IsDate(taskSheet.Cells(taskRow, "E").Value) And IsDate(taskSheet.Cells(taskRow, "F").Value) Then
                taskStartDate = taskSheet.Cells(taskRow, "E").Value ' Start Date
                taskEndDate = taskSheet.Cells(taskRow, "F").Value ' End Date
                taskName = taskSheet.Cells(taskRow, "B").Value ' Task Name
                
                
                If currentDate >= taskStartDate And currentDate <= taskEndDate Or currentDate = taskStartDate Or currentDate = taskEndDate Then
                    
                    Dim dataProcessingRow As Long
                    For dataProcessingRow = 3 To dataProcessingSheet.Cells(dataProcessingSheet.Rows.Count, "A").End(xlUp).Row
                        If dataProcessingSheet.Cells(dataProcessingRow, "A").Value = taskName Then
                            taskHours = dataProcessingSheet.Cells(dataProcessingRow, "B").Value ' Hours worked on task
                            
                            
                            commentText = commentText & taskName & " for " & taskHours & " hrs, "
                        End If
                    Next dataProcessingRow
                End If
            End If
        Next taskRow
        
        
        If Len(commentText) > 15 Then
            commentText = Left(commentText, Len(commentText) - 2)
        Else
            commentText = commentText & "No tasks."
        End If
        
        
        dateCell.ClearComments
        dateCell.AddComment
        dateCell.Comment.Text Text:=commentText
    Next dateCell

   
    Application.ScreenUpdating = True
End Sub
