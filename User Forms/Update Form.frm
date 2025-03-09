VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UpdateForm 
   Caption         =   "UserForm1"
   ClientHeight    =   2930
   ClientLeft      =   100
   ClientTop       =   420
   ClientWidth     =   7780
   OleObjectBlob   =   "Update Form.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UpdateForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
   
    Set ws = ThisWorkbook.Sheets("Task Tracking Sheet")
    
    
    If ListBox1.ListIndex = -1 Then
        MsgBox "Please select a task from the list."
        Exit Sub
    End If
    

    If ComboBox1.ListIndex = -1 Then
        MsgBox "Please select a progress percentage."
        Exit Sub
    End If
    
    
    Dim tbl As ListObject
    Set tbl = ws.ListObjects("InputTable")
    If tbl.DataBodyRange Is Nothing Then
        MsgBox "No tasks found in the 'Task Tracking Sheet'."
        Exit Sub
    End If
    
    ' If progress is 100%
    If ComboBox1.Value = "100%" Then
        Dim response As VbMsgBoxResult
        response = MsgBox("Are you sure you want to delete the task '" & ListBox1.Value & "'?", _
                          vbYesNo + vbQuestion, "Confirm Deletion")
        If response = vbNo Then
            Exit Sub
        End If
    End If
    
    
    Dim selectedItem As String
    selectedItem = ListBox1.Value
    Dim progressValue As String
    progressValue = ComboBox1.Value
    
    Dim cell As Range
    For Each cell In ws.Range("B1:B100000")
        If ComboBox1.Value = "100%" Then
            If cell.Value = selectedItem Then
                cell.EntireRow.Delete
                Exit For
            End If
        ElseIf cell.Value = selectedItem Then
            cell.Offset(0, 6).Value = ComboBox1.Value
            cell.Offset(0, 7).Value = cell.Offset(0, 7).Value - (cell.Offset(0, 7).Value * cell.Offset(0, 6).Value)
        End If
    Next
End Sub


Private Sub ListBox1_Click()

End Sub

Private Sub UserForm_Click()

End Sub

Private Sub UserForm_Initialize()
    'Intialize dropdown
    ComboBox1.AddItem ("10%")
    ComboBox1.AddItem ("20%")
    ComboBox1.AddItem ("30%")
    ComboBox1.AddItem ("40%")
    ComboBox1.AddItem ("50%")
    ComboBox1.AddItem ("60%")
    ComboBox1.AddItem ("70%")
    ComboBox1.AddItem ("80%")
    ComboBox1.AddItem ("90%")
    ComboBox1.AddItem ("100%")
    
    'intialize listbox
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim nameColumn As ListColumn
    Dim cell As Range

    ' Set the worksheet and table references
    Set ws = ThisWorkbook.Sheets("Task Tracking Sheet")
    Set tbl = ws.ListObjects("InputTable")
    Set nameColumn = tbl.ListColumns("Name")

    ' Clear the ListBox before adding items
    Me.ListBox1.Clear

    ' Loop through each cell in the Name column and add it to the ListBox
    For Each cell In nameColumn.DataBodyRange
        Me.ListBox1.AddItem cell.Value
    Next cell
End Sub
