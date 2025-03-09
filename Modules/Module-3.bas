Attribute VB_Name = "Module3"
Sub Motivational_Quote()
    Dim ws As Worksheet
    Dim quotesTable As ListObject
    Dim totalQuotes As Long
    Dim quote As String
    Dim randomRow As Long
    
    Set ws = ThisWorkbook.Sheets("Quotes List")
    Set quotesTable = ws.ListObjects("QuotesTable")
    totalQuotes = quotesTable.ListRows.Count
    Randomize
    randomRow = Int(totalQuotes * Rnd + 1)
    quote = quotesTable.ListRows(randomRow).Range.Cells(1, 2)
    
    MsgBox "Today's motivational quote is: " & vbNewLine & quote
End Sub
