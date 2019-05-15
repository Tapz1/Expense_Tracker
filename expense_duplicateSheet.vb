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
