Attribute VB_Name = "a_FileReformatter"
Function RawReformatter()
' Add headers for later - "D,DATE,ACCOUNT,LOCATION,CURRENCY,CATEGORY,DESCRIPTION,TYPE,AMOUNT,DETAILS"
    Sheet1.Rows(1).EntireRow.Insert
    Sheet1.Range("A1:K1") = Split("D,DATE,ROUTING,ACCOUNT,LOCATION,CURRENCY,CATEGORY,DESCRIPTION,TYPE,AMOUNT,DETAILS", ",")
' Unwrap text
    Sheet1.Cells.WrapText = False
' Autofit columns
    Sheet1.Columns("A:k").AutoFit
' Format Column D and L as number no decimals
    Sheet1.Columns("D").NumberFormat = "0"
    Sheet1.Columns("L").NumberFormat = "0"
' Column J to currency no decimals or symbols
    Sheet1.Columns("J").NumberFormat = "#,##0_);(#,##0)"
' Delete Column L+
    Sheet1.Columns("L:Z").Delete
' Sort by Column A
    Sheet1.Range("A:K").Sort key1:=Sheet1.Columns("A"), Order1:=xlAscending, Header:=xlYes
' Delete everything except stuff with "D" in Column A
    firstH = Application.WorksheetFunction.Match("H", Sheet1.Columns("A"), 0)
    Sheet1.Rows(firstH & ":50000").Delete
' All entries with acct #'s of <=399 convert to negative values
    Sheet1.Range("A:K").Sort key1:=Sheet1.Columns("G"), Order1:=xlAscending, Header:=xlYes
    last399 = Application.WorksheetFunction.Match(400, Sheet1.Columns("G"), 1)
    For i = 2 To last399
        Sheet1.Range("J" & i).Value = Sheet1.Range("J" & i).Value * -1
    Next i
End Function
