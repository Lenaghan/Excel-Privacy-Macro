Attribute VB_Name = "piiScrubber"
Sub GenerateNameCombinations()
    Dim ws                  As Worksheet
    Dim namesTable          As ListObject
    Dim firstNameCol        As ListColumn
    Dim lastNameCol         As ListColumn
    Dim i                   As Long
    Dim j                   As Long
    Dim k                   As Long
    Dim combinedNames()     As String
    
    Application.ScreenUpdating = False
    
    ' Assuming the NamesTable is in the active worksheet. You can modify this if needed.
    Set ws = ActiveSheet
    Set namesTable = ws.ListObjects("NamesTable")
    
    ' Assuming the columns are named "FirstName" and "LastName". Modify if needed.
    Set firstNameCol = namesTable.ListColumns("FirstName")
    Set lastNameCol = namesTable.ListColumns("LastName")
    
    ' Resize the array to hold all combinations
    ReDim combinedNames(1 To firstNameCol.DataBodyRange.Rows.Count * lastNameCol.DataBodyRange.Rows.Count, 1 To 1)
    
    ' Generate combinations
    k = 1
    For i = 1 To firstNameCol.DataBodyRange.Rows.Count
        For j = 1 To lastNameCol.DataBodyRange.Rows.Count
            combinedNames(k, 1) = firstNameCol.DataBodyRange.Cells(i).Value & " " & lastNameCol.DataBodyRange.Cells(j).Value
            k = k + 1
        Next j
    Next i
    
    square = WorksheetFunction.RoundDown(VBA.Sqr(UBound(combinedNames)), 0)
    
    ' Output loop
    rowCounter = 1
    For i = 1 To UBound(combinedNames) Step square
        For j = 1 To square
            ws.Cells(rowCounter, 5 + j - 1).Value = combinedNames(i + j - 1, 1)
        Next j
        rowCounter = rowCounter + 1
    Next i
End Sub
Sub ReplaceNamesWithFakes()
    Dim rng                 As Range
    Dim cell                As Range
    Dim targetSheet         As Worksheet
    Dim randomValue         As Variant
    Dim lastRow             As Long
    Dim firstNonEmptyCell   As Range
    Dim wb                  As Workbook
    Dim usedNames           As Collection
    Dim duplicateChance     As Double
    Dim randomRow           As Long
    
    ' adjust the constant to reflect the column to check
    Const NameColumn   As String = "C"
    
    Set wb = Workbooks("Names.xlsx") ' Set the workbook where the named range is located
    Set rng = wb.Names("FullNames").RefersToRange ' Set the named range you want to choose from
    Set targetSheet = ActiveSheet ' Set the target sheet as the active sheet
    
    lastRow = targetSheet.Cells(targetSheet.Rows.Count, NameColumn).End(xlUp).Row ' Find the last used row in the column
    
    ' Find the first non-empty cell in the column
    On Error Resume Next
    Set firstNonEmptyCell = targetSheet.Range(NameColumn & "1:" & NameColumn & lastRow).SpecialCells(xlCellTypeConstants).Cells(2)
    On Error GoTo 0
    
    If Not firstNonEmptyCell Is Nothing Then
        ' Initialize collection to store used names
        Set usedNames = New Collection
        
        ' Loop through each cell in column C starting from the second non-empty cell
        For Each cell In targetSheet.Range(firstNonEmptyCell, targetSheet.Cells(lastRow, NameColumn))
            ' Check if the cell is empty
            If Not IsEmpty(cell.Value) Then
                ResetCell cell
                ' Generate a random value to decide whether to use a duplicate name
                duplicateChance = Rnd()
                If duplicateChance <= 0.15 And usedNames.Count > 0 Then
                    ' Use a previously used name
                    randomValue = usedNames.Item(Int(Rnd() * usedNames.Count) + 1)
                    ' Randomize the row to paste the name
                    Do
                        randomRow = Int(Rnd() * (lastRow - firstNonEmptyCell.Row + 1)) + firstNonEmptyCell.Row
                    Loop While IsEmpty(targetSheet.Cells(randomRow, NameColumn).Value)
                    targetSheet.Cells(randomRow, NameColumn).Value = randomValue
                Else
                    ' Generate a random index within the range
                    randomValue = rng.Cells(Int(Rnd() * rng.Cells.Count) + 1).Value
                    ' Add the name to the used names collection
                    usedNames.Add randomValue
                    cell.Value = randomValue
                End If
            End If
        Next cell
    Else
        MsgBox "No non-empty cells found in column.", vbExclamation
    End If
' would be nice to make a pop up dialog for the choice of columns
' also need to change it to skip cells that have the text "Name" in it so it skips headers, instead of just the first non empty cell
End Sub
Sub ResetCell(cell)
    With cell
        .Interior.Pattern = xlNone
        .Interior.TintAndShade = 0
        .Interior.PatternTintAndShade = 0
        .Font.Bold = False
        .Font.Italic = False
        .Font.Underline = False
        .Font.ColorIndex = xlAutomatic
        .Font.Name = "Arial"
        .Font.Size = 11
        .WrapText = False
    End With
End Sub
Private Sub ReplacePIIWithRegex()

    Dim regEx               As New RegExp
    Dim strPattern          As Variant
    Dim strReplace          As String
    Dim strInput            As String
    Dim strOuput            As String
    Dim myRange             As Range
    Dim lastRow             As Long
    Dim lastColumn          As Long
    Dim arrayIndex          As Long
    Dim replacementCount    As Long
    
    strReplace = ""
    strPattern = Array("^d{3}-\d{2}-\d{4}$", "^\d{9}$", "^\d{3}\w\d{2}\w\d{4}$$", "^[A-PR-WY][1-9]\d\s?\d{4}[1-9]$", "^[a-zA-Z0-9._%-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,6}", "^[\w\.=-]+@[\w\.-]+\.[\w]{2,3}$")
    replacementCount = 0
    
    lastRow = Cells(Rows.Count, "A").End(xlUp).Row ' The amount of stuff in the column
    lastColumn = Cells(1, Columns.Count).End(xlToLeft).Column ' The amount of stuff in the row
    
    Set regEx = Nothing
    Set myRange = ActiveSheet.Range("A1:" & Cells(lastRow, lastColumn).Address)

    
    For Each Pattern In strPattern
        For Each c In myRange
            
            If strPattern(arrayIndex) <> "" Then
                strInput = c.Value
                
                With regEx
                    .Global = True
                    .MultiLine = True
                    .IgnoreCase = False
                    .Pattern = strPattern(arrayIndex)
                End With
    
                If regEx.test(strInput) Then
                    strOutput = regEx.Replace(strInput, strReplace)
                    
                    If strOutput <> strInput Then
                        c.Value = strOutput
                        replacementCount = replacementCount + 1
                    End If ' end of replacement coujnter increment
                    
                End If 'end of replace if found
                
            End If 'end of if cell not blank, check if it matches
            
        Next 'end of loop through range
        arrayIndex = arrayIndex + 1
    Next 'end of loop through patterns
    MsgBox "Total successful replacements: " & replacementCount, vbInformation, "Replacement Count"
    End Sub
