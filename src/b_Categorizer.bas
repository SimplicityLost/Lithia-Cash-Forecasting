Attribute VB_Name = "b_Categorizer"
Function Categorize()
   
    'Read strings from string sheet into array
    Dim SearchArray(), SearchStrings
    lastrow = Sheet3.Cells(Sheet3.Rows.Count, "A").End(xlUp).Row
    ReDim SearchArray(1 To lastrow - 1)
    For i = 1 To lastrow - 1
        SearchStrings = Split(Sheet3.Range("A" & i + 1) & ";" & Sheet3.Range("B" & i + 1) & ";" & Sheet3.Range("C" & i + 1), ";")
        SearchArray(i) = SearchStrings
    Next i
    
    lastrow = Sheet1.Cells(Sheet1.Rows.Count, "A").End(xlUp).Row
    
    'For each category and type
    For Each cattype In SearchArray()
        For Each srchstr In Split(cattype(2), ",")
            Sheet3.Range("F2") = cattype(0)
            Sheet3.Range("G2") = srchstr
            
'            If cattype(1) = "MTG" Then
'            sheet1.Range("A:K").AdvancedFilter _
'                Action:=xlFilterInPlace, _
'                CriteriaRange:=sheet3.Range("F1:H2"), _
'                Unique:=False
'            Else
            Sheet1.Range("A:K").AdvancedFilter _
                Action:=xlFilterInPlace, _
                CriteriaRange:=Sheet3.Range("F1:G2"), _
                Unique:=False
'            End If
            
            Dim workingRng As Range
            Set workingRng = Sheet1.Range("A:K").SpecialCells(xlCellTypeVisible)
                
            For Each rngArea In workingRng.Areas
                For Each cell In rngArea.Columns("I").Cells
                    If Not cell.Value = "TYPE" And Not cell.Value = "" Then
                        cell.Value = cattype(1)
                    End If
                    If cell.Row > lastrow Or cell.Value = "" Then Exit For
                Next cell
            Next rngArea
            
            Sheet1.ShowAllData
        Next srchstr
    Next cattype
    'For each string
    'Search for string
    'If found, update the Type
    
End Function
