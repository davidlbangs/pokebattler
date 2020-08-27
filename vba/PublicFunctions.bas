Attribute VB_Name = "PublicFunctions"
Option Explicit
' Functions meant to be included directly in a spreadsheet, assigned to buttons, _
' assigned to the quick access toolbar, or called by worksheets to respond to events

Sub ToggleCommentIndicators()
    ' Recommended to Add this to Quick Action Toolbar

    If Application.DisplayCommentIndicator = xlNoIndicator Then
        Application.DisplayCommentIndicator = xlCommentIndicatorOnly
    Else
        Application.DisplayCommentIndicator = xlNoIndicator
    End If

End Sub

Sub PositionBattleScoreComment(ByVal cell As Range)

' When you click on a cell with a comment hovering, the comment should move to a fully visible position.

' Add this (below) to the each worksheet in which BattleComments should behave intelligently.
' Private Sub Worksheet_SelectionChange(ByVal Target As Range)
'    Call PositionBattleScoreComment(Target)
' End Sub

    If Application.DisplayCommentIndicator = xlNoIndicator Then Application.DisplayCommentIndicator = xlNoIndicator

    If Not (cell.Comment Is Nothing) Then
        Dim leftWindow As Single, rightWindow As Single
        
        ' The values Excel provides for right edge of Window is absolutely wrong. Frequently too high, so
        ' subtract a sufficiently high constant but do assume the cell with the comment is fully visible.
        
        leftWindow = ActiveWindow.VisibleRange.left
        rightWindow = Max(leftWindow + ActiveWindow.width - 350, cell.left + cell.width)
        
        With cell.Comment.Shape
        
        .left = Max(leftWindow + 5, Min(cell.left + (cell.width - .width) / 2, rightWindow - .width))
        .top = cell.top + cell.Height + 3
        
        If .top + .Height > ActiveWindow.VisibleRange.top + ActiveWindow.VisibleRange.Height Then
            .top = cell.top - .Height - 3
        End If

        .Visible = True
        
        End With
    End If

End Sub

Sub HideDisplayedComments()
    
    ' xlNoIndicator - Comments would not be visible unless we showed them.  We need to hide them.
    ' xlCommentAndIndicator - All comments are visible.  Leave them alone
    ' xlCommentIndicatorOnly - Excel is hiding comments every time the mouse is moved.  No action necessary.
    
    ' This hides all displayed comments!
    
    If Application.DisplayCommentIndicator = xlNoIndicator Then Application.DisplayCommentIndicator = xlNoIndicator

End Sub

Function ScoreFromComment(cellUpd As Range, Optional fMetaScores As Boolean = False) As Single
Dim strText As String

Dim iColor As Integer
Dim rng As Range

    If cellUpd.Comment Is Nothing Then GoTo NoComment
    
    strText = ParseSubstringBetween(cellUpd.Comment.Text, "Battle Score: ", Chr(10))
    
    If Not (IsNumeric(strText)) Then
        GoTo NoComment
    End If
    
    If fMetaScores Then
        ScoreFromComment = 1000 - CDec(strText)
    Else
        ScoreFromComment = CDec(strText)
    End If
    
    Exit Function
    
NoComment:
    ScoreFromComment = -1

End Function


Function IsTextInComment(cellUpd As Range, strSearch As String) As Boolean
Dim strText As String, valScore As Single

Dim iColor As Integer
Dim rng As Range, iLineFeed
    
    If cellUpd.Comment Is Nothing Then GoTo NoComment
    
    strText = cellUpd.Comment.Text
    IsTextInComment = InStr(strText, strSearch) <> 0
    
    Exit Function
    
NoComment:
    IsTextInComment = False

End Function


Function CountBattles(rng As Range) As Integer
    Dim iRow As Integer, iCol As Integer
    Dim count As Integer

    count = 0
    
    With rng
    If .Columns.count = 1 Then If .Columns(1).EntireColumn.Hidden Then GoTo Finished
    
    For iRow = 1 To .Rows.count
        If .Rows(iRow).EntireRow.Hidden Then GoTo NextRow
        
        For iCol = 1 To .Columns.count
            If .Columns.count > 1 Then If .Columns(iCol).EntireColumn.Hidden Then GoTo NextCol

            If ScoreFromComment(.Cells(iRow, iCol)) <> -1 Then count = count + 1
NextCol:
        Next iCol
NextRow:
    Next iRow

    
Finished:
    CountBattles = count
    End With

End Function

Function CountCommentsWithText(rng As Range, strSearch As String) As Integer
    Dim iRow As Integer, iCol As Integer
    Dim count As Integer

    count = 0
    
    With rng
    If .Columns.count = 1 Then If .Columns(1).EntireColumn.Hidden Then GoTo Finished
    
    For iRow = 1 To .Rows.count
        If .Rows(iRow).EntireRow.Hidden Then GoTo NextRow
        
        For iCol = 1 To .Columns.count
            If .Columns.count > 1 Then If .Columns(iCol).EntireColumn.Hidden Then GoTo NextCol

            If IsTextInComment(.Cells(iRow, iCol), strSearch) Then count = count + 1
NextCol:
        Next iCol
NextRow:
    Next iRow
    
Finished:
    CountCommentsWithText = count
    End With

End Function

Function CountScoresAbove(rng As Range, valScoreCompare As Single, Optional fMetaScores As Boolean = False) As Integer
    Dim iRow As Integer, iCol As Integer
    Dim count As Integer
    
    count = 0
    
    With rng
    If .Columns.count = 1 Then If .Columns(1).EntireColumn.Hidden Then GoTo Finished
    

    For iRow = 1 To .Rows.count
        If .Rows(iRow).EntireRow.Hidden Then GoTo NextRow
        
        For iCol = 1 To .Columns.count
            If .Columns.count > 1 Then If .Columns(iCol).EntireColumn.Hidden Then GoTo NextCol
        
            If ScoreFromComment(.Cells(iRow, iCol), fMetaScores) > valScoreCompare Then count = count + 1
NextCol:
        Next iCol

NextRow:
    Next iRow

    
Finished:
    CountScoresAbove = count
    End With

End Function

Function CountScoresBelow(rng As Range, valScoreCompare As Single, Optional fMetaScores As Boolean = False) As Integer
    Dim iRow As Integer, iCol As Integer
    Dim count As Integer, score As Integer
    
    count = 0
    
    With rng
    If .Columns.count = 1 Then If .Columns(1).EntireColumn.Hidden Then GoTo Finished

    For iRow = 1 To .Rows.count
        If .Rows(iRow).EntireRow.Hidden Then GoTo NextRow
        
        For iCol = 1 To .Columns.count
            If .Columns.count > 1 Then If .Columns(iCol).EntireColumn.Hidden Then GoTo NextCol
            
            score = ScoreFromComment(.Cells(iRow, iCol), fMetaScores)
            
            If score >= 0 And score < valScoreCompare Then count = count + 1
NextCol:
        Next iCol

NextRow:
    Next iRow

Finished:
    CountScoresBelow = count
    End With

End Function

Function AverageScore(rng As Range, Optional fMetaScores As Boolean = False) As Single

    Dim valScore As Single, sumScores As Single
    Dim iRow As Integer, iCol As Integer
    Dim count As Integer
    
    AverageScore = -1
    
    With rng
    If .Columns.count = 1 Then If .Columns(1).EntireColumn.Hidden Then Exit Function
    
    count = 0
    sumScores = 0
    
    For iRow = 1 To .Rows.count
        If .Rows(iRow).EntireRow.Hidden Then GoTo NextRow
        
        For iCol = 1 To .Columns.count
            If .Columns.count > 1 Then If .Columns(iCol).EntireColumn.Hidden Then GoTo NextCol
            
            valScore = ScoreFromComment(.Cells(iRow, iCol), fMetaScores)
            If valScore <> -1 Then
                sumScores = sumScores + valScore
                count = count + 1
            End If
NextCol:
        Next iCol

NextRow:
    Next iRow
    End With
    
    If count > 0 Then AverageScore = sumScores / count

End Function

Function StdDevOfScores(rng As Range, Optional fMetaScores As Boolean = False) As Single
    Dim valScore As Single, sumSquares As Single
    Dim valAverage As Single
    Dim iRow As Integer, iCol As Integer
    Dim count As Integer
    
    StdDevOfScores = -1
    
    With rng
    If .Columns.count = 1 Then If .Columns(1).EntireColumn.Hidden Then Exit Function
    
    count = 0
    sumSquares = 0
    valAverage = AverageScore(rng)
    
    For iRow = 1 To .Rows.count
        If .Rows(iRow).EntireRow.Hidden Then GoTo NextRow
        
        For iCol = 1 To .Columns.count
            If .Columns.count > 1 Then If .Columns(iCol).EntireColumn.Hidden Then GoTo NextCol
            
            valScore = ScoreFromComment(.Cells(iRow, iCol), fMetaScores)
            If valScore <> -1 Then
                sumSquares = sumSquares + (valScore - valAverage) ^ 2
                count = count + 1
            End If
NextCol:
            Next iCol
NextRow:
    Next iRow
    End With

    If count > 0 Then StdDevOfScores = Sqr(sumSquares / count)

End Function

Function ReportWins(rng As Range, Optional fMetaScores As Boolean = False) As String
    Dim val As Single
    
    val = CountScoresAbove(rng, 500, fMetaScores)
    ReportWins = "Wins:" & StrStatVal(val)
End Function

Function ReportLosses(rng As Range, Optional fMetaScores As Boolean = False) As String
    Dim val As Single
    
    val = CountScoresBelow(rng, 500, fMetaScores)
    ReportLosses = "Losses:" & StrStatVal(val)
End Function

Function ReportBigWins(rng As Range, Optional fMetaScores As Boolean = False) As String
    Dim val As Single
    
    val = CountScoresAbove(rng, 550, fMetaScores)
    ReportBigWins = "Big Wins:" & StrStatVal(val)
End Function

Function ReportBigLosses(rng As Range, Optional fMetaScores As Boolean = False) As String
    Dim val As Single
    
    val = CountScoresBelow(rng, 450, fMetaScores)
    ReportBigLosses = "Big Losses:" & StrStatVal(val)
End Function

Function ReportAverageScore(rng As Range, Optional fMetaScores As Boolean = False) As String
    Dim val As Single
    
    val = AverageScore(rng, fMetaScores)
    If val = -1 Then
        ReportAverageScore = "Ave Score:None"
    Else
        ReportAverageScore = "Ave Score:" & StrStatVal(val)
    End If

End Function

Function ReportStdDevOfScores(rng As Range, Optional fMetaScores As Boolean = False) As String
    Dim val As Single
    
    val = StdDevOfScores(rng, fMetaScores)
    If val = -1 Then
        ReportStdDevOfScores = "Std Dev:None"
    Else
        ReportStdDevOfScores = "Std Dev:" & StrStatVal(val)
    End If
End Function

Function ReportCommentsWithText(rng As Range, strSearch As String, Optional strReport As String = "") As String
    Dim count As Single
    
    count = CountCommentsWithText(rng, strSearch)
    If strReport = "" Then strReport = strSearch
    ReportCommentsWithText = strReport & ":" & StrStatVal(count)

End Function


Function BeautifyCsv(csvUgly As String) As String
    Dim str As String, csvPretty As String
    Dim iNextMove As Integer, strNextMove As String
        
    csvPretty = ParsePokemonName(csvUgly)
    
    iNextMove = 1
    strNextMove = ParseMoveName(csvUgly, iNextMove)
    
    While strNextMove <> ""
        csvPretty = csvPretty & ", " & strNextMove
        iNextMove = iNextMove + 1
        strNextMove = ParseMoveName(csvUgly, iNextMove)
    Wend
    
    BeautifyCsv = csvPretty

End Function

Function ValidateCsv(csv As String, Optional fExplicit As Boolean = False) As String
Dim str As String, iNextMove As Integer
Dim pk As Pokemon

    Call InitPokemon(pk, csv)
    ValidateCsv = StrValidatePk(pk)
    
    If fExplicit Then
        ValidateCsv = "Valid CSV"
    End If
    
End Function

Function DebuffWarning(csv As String) As String
    Dim strSymbols As String, strNextMove As String, iNextMove As Integer, str As String
    
    'On Windows, these warning symbols look best Verdana 18.
    
    str = ValidateCsv(csv)
    If str <> "" Then
        DebuffWarning = "Error"
        Exit Function
    End If
    
    iNextMove = 2
    strNextMove = ParseMoveName(csv, iNextMove)
    
    DebuffWarning = ""
    
    While strNextMove <> ""
    
        strSymbols = SpecialEffectWarning(strNextMove)
        
        If strSymbols <> "" Then
            ' DebuffWarning = SymbolForType(TypeOfChargeMove(strNextMove)) & " " & strSymbols
            DebuffWarning = strSymbols
            Exit Function
        End If
        
        iNextMove = iNextMove + 1
        strNextMove = ParseMoveName(csv, iNextMove)
    Wend
    
End Function

