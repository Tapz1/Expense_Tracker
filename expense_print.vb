''prints only the income entered
Sub PrintIncome()
        
    wsLR = Cells(Rows.Count, 2).End(xlUp).Row
    wsLC = Cells(2, Columns.Count).End(xlToLeft).Column
    Set PrintA = Range("b13:e" & wsLR) 'this sets the range
    'set the print area
    ActiveSheet.PageSetup.PrintArea = PrintA.Address(0, 0)
    ActiveSheet.PrintPreview
    
End Sub


''prints only the expenses entered
Sub PrintExpenses()
    
    wsLR = Cells(Rows.Count, 7).End(xlUp).Row
    wsLC = Cells(7, Columns.Count).End(xlToLeft).Column
    Set PrintA = Range("g13:k" & wsLR) 'this sets the range
    'set the print area
    ActiveSheet.PageSetup.PrintArea = PrintA.Address(0, 0)
    ActiveSheet.PrintPreview
    
End Sub
