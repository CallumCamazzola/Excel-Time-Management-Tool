Attribute VB_Name = "Module1"
Sub FilterTaskClear()
Attribute FilterTaskClear.VB_Description = "clears everything on filter task board so user can see new set of tasks"
Attribute FilterTaskClear.VB_ProcData.VB_Invoke_Func = " \n14"
'
' FilterTaskClear Macro
' clears everything on filter task board so user can see new set of tasks
'

'
    Selection.ClearContents
    Range("A6").Select
    Selection.ClearContents
    Range("G5:M6").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    Range("A4").Select
End Sub
