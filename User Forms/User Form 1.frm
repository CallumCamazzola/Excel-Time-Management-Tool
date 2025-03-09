VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Time Spending Input Form"
   ClientHeight    =   5890
   ClientLeft      =   120
   ClientTop       =   480
   ClientWidth     =   9060.001
   OleObjectBlob   =   "User Form 2.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Sub UserForm_Initialize()
    ListBox1.AddItem "Class"
    ListBox1.AddItem "Self - Study"
    ListBox1.AddItem "Commute"
    ListBox1.AddItem "Entertainment"
    ListBox1.AddItem "Others"
    ListBox1.AddItem "Sleep"
End Sub

Private Sub CommandButton1_Click()
    'CommandButton1 will input values to sheet.
    Dim lRow As Long
    Dim ws As Worksheet
    Dim lookupbool As Boolean
    lookupbool = False
    
    Set ws = Worksheets("Time Spending Input")
    lRow = ws.Cells(Rows.Count, 1).End(xlUp).Offset(1, 0).Row
    
    Dim recommendedTime As Double
    'A1:B7 is TimeList -> is a named range in the LookupList sheet that contains recommended time for each category
    'It is CRUCIAL to have 'false' as the 4th parameter of vlookup. True will treat 1 and 5 as an approximate match. Hence making it not useful in our context.
    recommendedTime = Application.WorksheetFunction.VLookup(Me.ListBox1.Value, Worksheets("LookupList").Range("TimeList"), 2, lookupbool)
    
    Dim actualTime As Double
    actualTime = Me.TextBox2.Value
    
    Dim description As String
    description = Me.TextBox1.Value
    
    'Use aforedescribed variables for each row-column combination
    With ws
        .Cells(lRow, 1).Value = Me.ListBox1.Value
        .Cells(lRow, 2).Value = actualTime
        .Cells(lRow, 3).Value = recommendedTime
        .Cells(lRow, 4).Value = description
    End With
    
    'Clear input controls.
    Me.ListBox1.Value = ""
    Me.TextBox1.Value = ""
    Me.TextBox2.Value = ""

End Sub



