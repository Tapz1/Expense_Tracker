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
