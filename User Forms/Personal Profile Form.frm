VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PersonalProfileForm 
   Caption         =   "UserForm1"
   ClientHeight    =   6050
   ClientLeft      =   100
   ClientTop       =   420
   ClientWidth     =   8520.001
   OleObjectBlob   =   "Personal Profile Form.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "PersonalProfileForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub ComboBox4_Change()

End Sub
Private Sub CommandButton1_Click()
    If ComboBox1.Value = "" Then
        MsgBox "Degree Level cannot be empty.", vbExclamation
        Exit Sub
    End If

    If ComboBox2.Value = "" Then
        MsgBox "Current Year cannot be empty.", vbExclamation
        Exit Sub
    End If

    If ListBox1.ListIndex = -1 Then
        MsgBox "Program cannot be empty.", vbExclamation
        Exit Sub
    End If

    Set WB = ThisWorkbook
    Set ws = WB.Worksheets("Personal Profile")

    ws.Cells(5, "B") = ComboBox1.Value
    ws.Cells(5, "C") = ComboBox2.Value
    ws.Cells(5, "D") = ListBox1.Value
End Sub
Private Sub CommandButton2_Click()
    If ComboBox4.Value = "" Then
        MsgBox "Commuter cannot be empty.", vbExclamation
        Exit Sub
    End If

    If Not IsNumeric(TextBox1.Value) Then
        MsgBox "Commute Length must contain a number.", vbExclamation
        Exit Sub
    End If

    Set WB = ThisWorkbook
    Set ws = WB.Worksheets("Personal Profile")

    ws.Cells(5, "F") = ComboBox4.Value
    ws.Cells(5, "G") = TextBox1.Value
End Sub
Private Sub CommandButton3_Click()
    If Not IsNumeric(TextBox2.Value) Then
        MsgBox "Average Time Spent Per Day must contain a number.", vbExclamation
        Exit Sub
    End If

    If ComboBox5.Value = "" Then
        MsgBox "Activity Type cannot be empty.", vbExclamation
        Exit Sub
    End If

    If ComboBox6.Value = "" Then
        MsgBox "Priority Level cannot be empty.", vbExclamation
        Exit Sub
    End If

    Set WB = ThisWorkbook
    Set ws = WB.Worksheets("Personal Profile")
    Dim intRow As Integer
    intRow = 5

    Do While (ws.Cells(intRow, "J") <> "")
        intRow = intRow + 1
    Loop

    ws.Cells(intRow, "J") = ComboBox5.Value
    ws.Cells(intRow, "K") = TextBox2.Value
    ws.Cells(intRow, "L") = ComboBox6.Value
End Sub
Private Sub CommandButton4_Click()
    If ComboBox7.Value = "" Or ComboBox8.Value = "" Or ComboBox9.Value = "" Or ComboBox10.Value = "" Or ComboBox11.Value = "" Then
        MsgBox "All hour selection must have a value.", vbExclamation
        Exit Sub
    End If

    Set WB = ThisWorkbook
    Set ws = WB.Worksheets("Personal Profile")

    ws.Cells(5, "N") = ComboBox7.Value
    ws.Cells(5, "O") = ComboBox8.Value
    ws.Cells(5, "P") = ComboBox9.Value
    ws.Cells(5, "Q") = ComboBox10.Value
    ws.Cells(5, "R") = ComboBox11.Value
End Sub


Private Sub UserForm_Initialize()
    'Intitalize Degree Level ComboBox
    ComboBox1.AddItem ("Undergraduate")
    ComboBox1.AddItem ("Diploma")
    ComboBox1.AddItem ("Masters")
    ComboBox1.AddItem ("PHD")
    ComboBox1.AddItem ("Other")
    
    'Intialize Current Year Form
    ComboBox2.AddItem ("1")
    ComboBox2.AddItem ("2")
    ComboBox2.AddItem ("3")
    ComboBox2.AddItem ("4")
    ComboBox2.AddItem ("5+")
    
    'Intialize Program ListBox
    ListBox1.AddItem ("Arts")
    ListBox1.AddItem ("Engineering")
    ListBox1.AddItem ("Science")
    ListBox1.AddItem ("Business")
    ListBox1.AddItem ("Health")
    ListBox1.AddItem ("Mathematics")
    ListBox1.AddItem ("Music")
    ListBox1.AddItem ("Government & Law")
    ListBox1.AddItem ("Education")
    ListBox1.AddItem ("Other")
    
    'initialize commuter combobox
    ComboBox4.AddItem ("Yes")
    ComboBox4.AddItem ("No")
    
    'Initialize activity type
    ComboBox5.AddItem ("Job")
    ComboBox5.AddItem ("Club")
    ComboBox5.AddItem ("Sport")
    ComboBox5.AddItem ("Hobby")
    ComboBox5.AddItem ("Free Time")
    ComboBox5.AddItem ("Shopping")
    ComboBox5.AddItem ("Spending Time With Friends And Family")
    ComboBox5.AddItem ("Other")
    
    'intialize priority level
    ComboBox6.AddItem ("High (Necessary)")
    ComboBox6.AddItem ("Medium (Preffered But Not Necessary)")
    ComboBox6.AddItem ("Low (Unecessary)")
    
    'Initialize monday hours
    ComboBox7.AddItem ("1")
    ComboBox7.AddItem ("2")
    ComboBox7.AddItem ("3")
    ComboBox7.AddItem ("4")
    ComboBox7.AddItem ("5")
    ComboBox7.AddItem ("6")
    ComboBox7.AddItem ("7+")
    'Initialize tuesday hours
    ComboBox8.AddItem ("1")
    ComboBox8.AddItem ("2")
    ComboBox8.AddItem ("3")
    ComboBox8.AddItem ("4")
    ComboBox8.AddItem ("5")
    ComboBox8.AddItem ("6")
    ComboBox8.AddItem ("7+")
    'Initialize wednesday hours
    ComboBox9.AddItem ("1")
    ComboBox9.AddItem ("2")
    ComboBox9.AddItem ("3")
    ComboBox9.AddItem ("4")
    ComboBox9.AddItem ("5")
    ComboBox9.AddItem ("6")
    ComboBox9.AddItem ("7+")
    'Initialize thursday hours
    ComboBox10.AddItem ("1")
    ComboBox10.AddItem ("2")
    ComboBox10.AddItem ("3")
    ComboBox10.AddItem ("4")
    ComboBox10.AddItem ("5")
    ComboBox10.AddItem ("6")
    ComboBox10.AddItem ("7+")
    'Initialize friday hours
    ComboBox11.AddItem ("1")
    ComboBox11.AddItem ("2")
    ComboBox11.AddItem ("3")
    ComboBox11.AddItem ("4")
    ComboBox11.AddItem ("5")
    ComboBox11.AddItem ("6")
    ComboBox11.AddItem ("7+")
End Sub
