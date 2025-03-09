Attribute VB_Name = "Module6"
Sub Graphic4_Click()
    ' Disable screen updating for better performance
    Application.ScreenUpdating = False
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Calendar Breakdown")
    
    
    Dim personalProfileSheet As Worksheet
    Set personalProfileSheet = ThisWorkbook.Sheets("Personal Profile")
    
    
    Dim lastRow As Long
    lastRow = personalProfileSheet.Cells(personalProfileSheet.Rows.Count, "J").End(xlUp).Row
    
    Dim dateCell As Range
    Dim activityName As String
    Dim activityHours As Double
    Dim commentText As String
    Dim i As Long
    Dim hourText As String
    
    
    For Each dateCell In ws.Range("B9:H9")
        
        commentText = "Tasks for today: "
        
    
        dateCell.ClearComments
        
     
        For i = 5 To lastRow
            activityName = personalProfileSheet.Cells(i, "J").Value
            activityHours = personalProfileSheet.Cells(i, "K").Value
            
            
            If activityHours = 1 Then
                hourText = "hr"
            Else
                hourText = "hrs"
            End If
            
           
            commentText = commentText & activityName & " " & activityHours & " " & hourText & ", "
        Next i
        
        
        If Len(commentText) > 0 Then
            commentText = Left(commentText, Len(commentText) - 2)
        End If
        
        
        dateCell.AddComment
        dateCell.Comment.Text Text:=commentText
    Next dateCell

    ' Re-enable screen updating
    Application.ScreenUpdating = True
End Sub

