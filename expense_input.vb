'''
'Programmer: Chris Tapia, TapzTracking
'Last Modified: 01/20/19
'Details: This code enables you to add input information to each line, print only a specific range of cells and to reproduce
'       the Main worksheet for more templates
'''


'''income
Sub submitIncome()

    Dim lastCell As Range
    
    
    MsgBox "Let's submit your income", vbOKOnly, "Income"
    the_date = InputBox("When was this earned?")
        If the_date = "" Then
            MsgBox "Cancelled!"
        Else
            amount = InputBox("How much did you earn?")
                If amount = "" Then
                    MsgBox "Cancelled!"
                Else
                    Details = InputBox("Provide details on this income")
                        If amount = "" Then
                            MsgBox "Cancelled!"
                        Else
                            MsgBox "Perfect, you can now view this under your income!"
                        
                        End If
                End If
        End If
                        
    
    'entering the date
    Set lastCell = Range("B" & Rows.Count).End(xlUp)
    Range("B" & lastCell.Row).Offset(1, 0).Value = the_date
    
    'entering the amount
    Set lastCell = Range("E" & Rows.Count).End(xlUp)
    Range("E" & lastCell.Row).Offset(1, 0).Value = amount
    
    'entering the details
    Set lastCell = Range("C" & Rows.Count).End(xlUp)
    Range("C" & lastCell.Row).Offset(1, 0).Value = Details

End Sub



'''expenses
Sub submitExpenses()

    Dim lastCell As Range
    
    
    MsgBox "Let's submit an expense", vbOKOnly, "Expenses"
    the_date = InputBox("When did this expense occur?")
        If the_date = "" Then
            MsgBox "Cancelled!"
        Else
            amount = InputBox("How much did you spend?")
                If amount = "" Then
                    MsgBox "Cancelled!"
                Else
                    Details = InputBox("Provide details on this expense")
                        If Details = "" Then
                            MsgBox "Cancelled!"
                        Else
                            MsgBox "Perfect, you can now view this under your expenses!"
                        
                        End If
                End If
        End If
    
    
    'entering the date
    Set lastCell = Range("G" & Rows.Count).End(xlUp)
    Range("G" & lastCell.Row).Offset(1, 0).Value = the_date
    
    'entering the amount
    Set lastCell = Range("K" & Rows.Count).End(xlUp)
    Range("K" & lastCell.Row).Offset(1, 0).Value = amount
    
    'entering the details
    Set lastCell = Range("H" & Rows.Count).End(xlUp)
    Range("H" & lastCell.Row).Offset(1, 0).Value = Details


End Sub



''' creates a copy of the primary sheet
Sub duplicateSheet()
    Dim I As Long
    Dim xNumber As Integer
    Dim pageName As String
    Dim xName As String
    Dim xActiveSheet As Worksheet
    On Error Resume Next
    Application.ScreenUpdating = False
    Set xActiveSheet = ThisWorkbook.Sheets("Main")
    xNumber = InputBox("Enter the number of templates you'd like to make")
    pageName = InputBox("What would you like to call it?")
    For I = 1 To xNumber
        xName = ActiveSheet.Name
        xActiveSheet.Copy After:=ActiveWorkbook.Sheets(xName)
        ActiveSheet.Name = pageName & I
    Next
    xActiveSheet.Activate
    Application.ScreenUpdating = True
End Sub



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

