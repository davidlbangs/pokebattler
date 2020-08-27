Attribute VB_Name = "BattleDisplay"
' Pokemon Go Battle Planner
' (c) 2020 David Bangs.  All rights reserved
'
' A tool to help PVP particants plan an ideal team and to guide them through battles by providing a heads up view of
' best moves and matchups.

Option Explicit

Private Const font_Text As String = "Verdana"
Private Const font_Symbol As String = "Segoe UI Symbol"
Private Const color_ShadeError As Double = &HAAFFAA
Private Const color_BattleTableBorder As Double = &HAAAAAA
Private Const color_PokemonGroupBorder As Double = &HAA4444
Private Const color_NeutralScore As Double = 15855598  ' Should be color for 500 in _ColorGradientTable

Private Const str_MoveRecommend As String = " used "

Private Type BestMoveAnalysis
    strSpecialInstruction As String ' Special instruction given through column dropbown.
    fPokemonBenched As Boolean ' This Pokemon is benched.  Do not include in best move analysis.

    'for displaying the best counter , and second best counter the team could use against an oponent
    iColBest As Integer
    valScoreBest As Single
    iColSecondBest As Integer
    valScoreSecondBest As Single
    strPokemonBest As String

    'for marking the best move that each individual pokemon could use against the oponent.
    cSamePokemonBlocks As Integer
    iColSamePokemon As Integer
    iColBestSamePokemon As Integer
    iColLastSamePokemon As Integer
    strSamePokemon As String
    valScoreBestSamePokemon As Single
    valScoreWorstSamePokemon As Single
End Type

Private Type BattleDisplayMode
    strDisplayMethod As String
    strBattleLeague As String
    rngDataBattleLeague As Range

    fChargeMoveAlerts As Boolean
    fChargeMoveAlerts_Shields As Boolean
    fChargeMoveAlerts_Bolts As Boolean
    
    fBestCounterIcons As Boolean
    fBestCounterIcons_Large As Boolean
    fBestCounterIcons_CloseSecond As Boolean
End Type

Option Compare Text

Dim breakPoint As Integer

Sub UpdateDisplayQuick()

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    Call UpdateTeamDisplay
'   Call UpdateMetaDisplay ' use formulas for now

    Call UpdateBattleDisplay(True) ' incremental update

    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic

End Sub

Sub RecalcBattles()
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    Call UpdateTeamDisplay
'   Call UpdateMetaDisplay ' use formulas for now

    Call UpdateBattleDisplay(False) ' full update
    
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic

End Sub


Sub UpdateTeamDisplay()
Dim rngCsv As Range, rngDisplay As Range, rngTeamData As Range
Dim csv As String
Dim cCol As Integer, iCol As Integer
Dim cRow As Integer, iRow As Integer
  
    Set rngCsv = ValidateRange("TeamList").Rows(1).EntireRow
    Set rngDisplay = ValidateRange("TeamDisplay")
    
    With rngDisplay
    
    cCol = .Columns.count
    cRow = .Rows.count
    
    'First and last columns of Team Display are left blank, as they align with TableTable border columns.

    For iCol = 2 To cCol - 1
        If IsColumnVisible(.Columns(iCol)) Then
            Set rngTeamData = Application.Intersect(.Columns(iCol).EntireColumn, rngCsv)
            csv = rngTeamData.Text
            
            Call DisplayPokemonName(.Cells(1, iCol), csv)
            
            If cRow > 1 Then
                Call DisplayQuickMoveName(.Cells(2, iCol), csv, 1)
                
                For iRow = 3 To cRow
                    Call DisplayChargeMoveName(.Cells(iRow, iCol), csv, iRow - 1)
                Next iRow
            End If
        End If
    Next iCol

    End With

End Sub


Sub UpdateBattleDisplay(Optional fScoresFromComments As Boolean = False)

Dim bdm As BattleDisplayMode
Dim selSave As Range
Dim rngUpdate As Range, rngTeam As Range, rngTeamDisplay As Range, rngMeta As Range
Dim rngBlockBorders As Range ' Display border around every other pokemon block
Dim cellForSizing As Range
Dim iCol As Integer, cCol As Integer
Dim iRow As Integer, cRow As Integer, iRowVis As Integer
Dim timerStart As Single, timerRow As Single, timerNow As Single

Dim cCalc As Integer

    timerStart = Timer

''    On Error GoTo NoRange

    Set rngUpdate = ValidateRange("BattleDisplay")
    Set rngTeam = ValidateRange("TeamList").EntireRow
    Set rngTeamDisplay = ValidateRange("TeamDisplay")
    Set rngMeta = ValidateRange("MetaList")
    
    Call GetBattleDisplayMode(bdm)

    If bdm.rngDataBattleLeague Is Nothing Then
        MsgBox "Text in BattleLeague cell is invalid"
        GoTo NoRange
    End If
    
    With rngUpdate
    
    cCol = .Columns.count
    cRow = .Rows.count
    
'First and last column, and first and last row are to be left blank and painted using neutral gray color.
'This serves as a border which helps people determine which color is gray (battle neutral) in different lighting,
'It also lets people add and delete pokemon rows and columns without accidentally distorting the named range or cell border lines.

    Call EraseBorders(CombineRanges(rngTeamDisplay.Rows(1), rngUpdate.Rows(cRow)))
    Call DrawOutsideBorder(rngUpdate, color_BattleTableBorder)
    Call PaintNeutral(.Columns(1).Cells)
    Call PaintNeutral(.Columns(cCol).Cells)
    Call PaintNeutral(.Rows(1).Cells)
    Call PaintNeutral(.Rows(cRow).Cells)
    
    Call DeleteBestCounterIcons

    
    Set selSave = RangeSelection()
    
    Call TimeCheck(timerStart, 1)

    ' we will be drawing a border around every-other pokemon group.
    Set rngBlockBorders = CombineRanges(rngTeamDisplay.Rows(1), rngUpdate.Rows(cRow - 1))
    
    For iRow = 2 To cRow - 1
        Dim rngMetaData As Range, rngTeamData As Range
        Dim pkTeam As Pokemon, pkMetaRow As Pokemon, pkMeta As Pokemon
        Dim csv As String, csvIV As String
        Dim bmaInit As BestMoveAnalysis, bma As BestMoveAnalysis
        
        timerRow = Timer
        
        If .Rows(iRow).EntireRow.Hidden Then
            ' erase hidden rows if we are actually recalculating while they are hidden.
                If Not fScoresFromComments Then Call PaintNeutral(.Rows(iRow).Cells)
            GoTo NextRow
        End If

        iRowVis = iRowVis + 1
        
        Set rngMetaData = Application.Intersect(.Rows(iRow).EntireRow, rngMeta)
        csv = Trim(rngMetaData.Cells(1, 1).Text)
        
        If rngMetaData.count > 1 Then csvIV = rngMetaData.Cells(1, 2).Text Else csvIV = ""
        If csvIV = "" Then csvIV = "13,13,13"  ' default csvIV
        
        Call InitPokemon(pkMetaRow, csv)
        Call QualifyPokemon(pkMetaRow, csvIV, bdm.rngDataBattleLeague, True)
        
        bma = bmaInit  'Move analysis starts over each row
                
        Call TimeCheck(timerRow)
        
        For iCol = 2 To cCol - 1
            Dim valScore As Single, valScoreMod As Single
            Dim cellUpd As Range
            
            Set cellUpd = .Cells(iRow, iCol)
            
            If .Columns(iCol).EntireColumn.Hidden Then
                ' erase hidden columns if we are actually recalculating while they are hidden.
                If Not fScoresFromComments Then Call PaintNeutral(cellUpd)
                GoTo NextCol
            End If
            
            Set rngTeamData = Application.Intersect(.Columns(iCol).EntireColumn, rngTeam)
            csv = Trim(rngTeamData.Cells(1, 1).Text)
            
            If rngTeamData.count > 1 Then csvIV = rngTeamData.Cells(2, 1).Text Else csvIV = ""
            If csvIV = "" Then csvIV = "13,13,13"  ' default csvIV
            
            Call InitPokemon(pkTeam, csv)
            
            Call QualifyPokemon(pkTeam, csvIV, bdm.rngDataBattleLeague, False)
            
            pkMeta = pkMetaRow ' forget previous battle data.
            
            bma.strSpecialInstruction = ""
            If rngTeamData.count > 2 Then bma.strSpecialInstruction = rngTeamData.Cells(3, 1).Text
            bma.fPokemonBenched = bma.strSpecialInstruction = "Benched"
            
            Call MarkCellPokemonBenched(cellUpd, bma.fPokemonBenched)
            
            timerNow = Timer - timerStart
            If False Then
'            If timerNow > 30 Then
                Dim su As Boolean
                
                su = Application.ScreenUpdating
                Application.ScreenUpdating = True
                
                If MsgBox("Calculating row " & iRow & ". " & Chr(10) & _
                RoundDown(timerNow) & " seconds to do " & cCalc & " simulations." & Chr(10) & _
                "This is taking too long.  Do you want to continue?", vbYesNo, "Battle Calculation") <> vbYes Then Exit Sub
                
                timerStart = Timer: timerRow = timerStart
                cCalc = 0
                Application.ScreenUpdating = su
            End If
            
            If selSave.Address = cellUpd.Address Then
                breakPoint = 0
            End If
        
            If fScoresFromComments Then
                ' We're doing an incremental update. If there is valid information in stored in the cell comments, use it.
                ' If the comment is missing or doesn't match, run a new battle.
            
                Dim pkTeamInComment As Pokemon, pkMetaInComment As Pokemon
                pkTeamInComment = pkTeam: pkMetaInComment = pkMeta
            
                valScore = BattleScoreFromComment(cellUpd, pkTeamInComment, pkMetaInComment)
                If valScore = -1 Or pkTeam.csv <> pkTeamInComment.csv Or pkMeta.csv <> pkMetaInComment.csv Or _
                    pkTeam.ivs.csvIV <> pkTeamInComment.ivs.csvIV Or pkMeta.ivs.csvIV <> pkMetaInComment.ivs.csvIV Then
                    GoTo NewScore
                End If
                    
                pkTeam = pkTeamInComment: pkMeta = pkMetaInComment
                
                Call TrimCommentAtSymbol(cellUpd, "~") ' remove Best Counter info from comment, since the winner may change.
            Else
NewScore:
                valScore = BattleScoreCore(pkTeam, pkMeta)
                
                If valScore = -1 Then
                    Call DeleteComment(cellUpd)
                Else
                    Call AddBattleScoreComment(cellUpd, valScore, pkTeam, pkMeta, bdm)
                    rngTeamData.Cells(1.1) = pkTeam.csv ' Beautify the CSV so the user can't make it ugly and non-standard.
                    rngMetaData.Cells(1, 1) = pkMeta.csv
                End If
                
                cCalc = cCalc + 1
            End If
            
            Call SetColorSmart(cellUpd, ColorFromScore(valScore))
            
            ' NOTE:  It is vital to change the value of cellUpd to ensure that formulas, _
                        which depend only on the content of comments, update.
                        
            If valScore = -1 Then
                Dim strError1 As String, strError2 As String
                
                strError1 = StrValidatePk(pkTeam)
                strError2 = StrValidatePk(pkMeta)
                
                If strError1 <> "" Then
                    cellUpd.value = strError1
                ElseIf strError2 <> "" Then
                    cellUpd.value = strError2
                ElseIf Not pkTeam.fQualified Or Not pkMeta.fQualified Then
                    cellUpd.value = "Disqualified"
                Else
                    cellUpd.value = "No Score"
                End If
            Else
                Select Case bdm.strDisplayMethod
                    Case "Team Scores"
                        cellUpd.value = valScore
                    Case "Team Charge Moves"
                        cellUpd.value = pkTeam.cm.strMove
                    Case "Team Moves If Multiple"
                        If pkTeam.fMultipleChargeMoves Then cellUpd.value = pkTeam.cm.strMove Else cellUpd.value = pkMeta.strName
                    Case "Team Pokemon"
                        cellUpd.value = pkTeam.strName
                    Case "Meta Scores"
                        cellUpd.value = 1000 - valScore ' score of oponent
                    Case "Meta Charge Moves"
                        cellUpd.value = pkMeta.cm.strMove
                    Case "Meta Pokemon"
                        cellUpd.value = pkMeta.strName
                End Select
            End If
            
            If bdm.fChargeMoveAlerts And Not bma.fPokemonBenched Then
                Call DisplayChargeMoveAlerts(cellUpd, cellForSizing, iRowVis, bdm, pkTeam, pkMeta)
            End If
            
            ' Best Move Analysis across entire row
            If Not bma.fPokemonBenched Then
                valScoreMod = valScore - 3 * pkMeta.cmBest.threat.threatLevel ' valScoreMod tries to recommend a team member not needing shields.
            
                If valScoreMod > bma.valScoreBest Then
                    If bma.strPokemonBest <> pkTeam.strName Then
                        bma.valScoreSecondBest = bma.valScoreBest
                        bma.iColSecondBest = bma.iColBest
                    End If
                    
                    bma.valScoreBest = valScoreMod
                    bma.iColBest = iCol
                    bma.strPokemonBest = pkTeam.strName
                ElseIf valScoreMod > bma.valScoreSecondBest And pkTeam.strName <> bma.strPokemonBest Then
                    bma.valScoreSecondBest = valScoreMod
                    bma.iColSecondBest = iCol
                End If
            End If
            
            ' Best Move Analysis and border drawing for a block of columns representing move options for a single pokemon.
            Call MarkPokemonBlock(bma, pkTeam.strName, rngUpdate, iRow, iCol, valScore, rngBlockBorders)
            
NextCol:
            Call TimeCheck(timerRow, 0.1 * iCol)
        Next iCol
        
        Call MarkPokemonBlock(bma, "", rngUpdate, iRow, cCol, 0, rngBlockBorders)
        Set rngBlockBorders = Nothing  ' only display the pokemon block borders once.
        
        If bdm.fBestCounterIcons Then
            If bma.iColBest > 0 Then
                Call DisplayBestCounterIcon(.Cells(iRow, bma.iColBest), cellForSizing, iRowVis, pkMeta, bdm, False)
            End If
            
            Dim deltaSecond As Single
            deltaSecond = 30 + (Abs(bma.valScoreBest - 500) / 8) ' the acceptible difference between first and second best grows with distance from a tie.
            
            If bma.iColSecondBest > 0 And bma.valScoreSecondBest > bma.valScoreBest - deltaSecond Then
                Call DisplayBestCounterIcon(.Cells(iRow, bma.iColSecondBest), cellForSizing, iRowVis, pkMeta, bdm, bma.valScoreSecondBest < bma.valScoreBest)
            End If
        End If
        
        Call TimeCheck(timerRow, 0.1 * iCol)
        
NextRow:
    Next iRow
    End With
    
NoRange:

    Call SelectRange(selSave)

End Sub

Function RangeSelection() As Range

    On Error GoTo NoRange
    Set RangeSelection = Selection
    Exit Function
    
NoRange:
    Set RangeSelection = Nothing
End Function

Sub SelectRange(rng As Range)

    On Error GoTo WTF

    If Not (rng Is Nothing) Then rng.Select
    
WTF:

End Sub


Sub DisplayPokemonName(cellUpd As Range, csvPokemon As String)
    Dim strPokemon As String, strSymbols As String
    
    strPokemon = ParsePokemonName(csvPokemon)
    strSymbols = SymbolsForPokemon(strPokemon)
    
    Call DisplayTextAndSymbols(cellUpd, strPokemon, strSymbols)

End Sub

Function StrStatVal(val As Single, Optional cChar As Integer = 4) As String
    ' rounded number expressed with leading spaces to make up 4 characters. Right aligned, sort well.
    StrStatVal = Right("    " & Round(val, 0), cChar)

End Function

Sub DisplayQuickMoveName(cellUpd As Range, csvPokemon As String, iMove As Integer)
    Dim strMove As String, strSymbols As String
    
    strMove = ParseMoveName(csvPokemon, iMove)
    strSymbols = SymbolForType(TypeOfQuickMove(strMove))
    
    Call DisplayTextAndSymbols(cellUpd, strMove, strSymbols)
End Sub

Sub DisplayChargeMoveName(cellUpd As Range, csvPokemon As String, iMove As Integer)
    Dim strMove As String, strSymbols As String
    
    strMove = ParseMoveName(csvPokemon, iMove)
    strSymbols = SymbolForType(TypeOfChargeMove(strMove))
    
    Call DisplayTextAndSymbols(cellUpd, strMove, strSymbols)
End Sub



Sub DisplayTextAndSymbols(cellUpd As Range, strText As String, strSymbols As String)
    Dim iSymbol As Integer, lenSymbol As Integer
    
    On Error GoTo DisplayError
    
    iSymbol = Len(strText) + 1
    lenSymbol = Len(strSymbols) + 1 ' adding one seems to be necessary to include last symbol.
    
    Call SetCellText(cellUpd, strText & " " & strSymbols)
    
    cellUpd.Font.Name = font_Text
    cellUpd.Characters(Start:=iSymbol, Length:=lenSymbol).Font.Name = font_Symbol
    
DisplayError:

End Sub


Private Function ColorFromScore(score As Single) As Double

Dim iColor As Integer
Dim rng As Range
Dim colorSet As Double

    'red blue and green should be values between 0 and 255 based on the score.
    'low scores have high blue, high scores have high red.
    'scores should be between 0 and 1000.
    
    On Error GoTo NoColor
    
    If score <> -1 Then
        iColor = score / 10 + 1
        If iColor < 1 Then iColor = 1 Else If iColor > 100 Then iColor = 100
        
        
        Set rng = Range("_ColorGradientTable")
        
        colorSet = rng.Cells(iColor, 1)
        
        ColorFromScore = colorSet
        Exit Function
    End If
    
NoColor:
    ColorFromScore = color_ShadeError

End Function


Function ScoreFromColor(color As Double) As Single

Dim iColor As Integer
Dim rng As Range
    
    On Error GoTo NoColor
    
    Set rng = Range("_ColorGradientTable")
    
    iColor = Application.Match(color, rng.Columns(1), 0) - 1
    
    ScoreFromColor = 10 * iColor
    Exit Function
    
NoColor:
    ScoreFromColor = -1

End Function

Sub PaintNeutral(rng As Range)
    Dim cell As Range
    
    For Each cell In rng
        If cell.Interior.color <> color_NeutralScore Then cell.Interior.color = color_NeutralScore
        If cell.Text <> "" Then cell.value = ""
        If Not (cell.Comment Is Nothing) Then cell.Comment.Delete
    Next cell
End Sub


Private Sub SetColorSmart(cellUpd As Range, color As Double)
    'white is just no fill in our case.
    If color = &HFFFFFF Then
        If cellUpd.Interior.ColorIndex <> xlColorIndexNone Then cellUpd.Interior.ColorIndex = xlColorIndexNone
    Else
        If cellUpd.Interior.color <> color Then cellUpd.Interior.color = color
    End If
    
End Sub

Function IsColumnVisible(rng As Range) As Boolean
    
    On Error GoTo IsVisible
    IsColumnVisible = rng.EntireColumn.Hidden <> True
    Exit Function
    
IsVisible:
    IsColumnVisible = True

End Function

Function IsRowVisible(rng As Range) As Boolean
    
    On Error GoTo IsVisible
    IsRowVisible = rng.EntireRow.Hidden <> True
    Exit Function
    
IsVisible:
    IsRowVisible = True

End Function




Sub DisplayBestCounterIcon(ByVal cell As Range, ByRef cellForSizing As Range, ByRef iRowVis As Integer, pkMeta As Pokemon, bdm As BattleDisplayMode, fSecondBest As Boolean)
    Dim sFound As Shape, sNew As Shape
    Dim strPokemon As String
    
    On Error GoTo NoIcon
    
    If fSecondBest Then
        ' only display second best icon if backup winner is desired.
        If Not bdm.fBestCounterIcons_CloseSecond Then
            Exit Sub
        End If
        Call AppendToComment(cell, "~Close Second Best Counter")
    Else
        Call AppendToComment(cell, "~Best Counter")
    End If
    
    Set sFound = FindPokemonShape(pkMeta.strNameData)
    
    If Not (sFound Is Nothing) Then
        Dim xCenter As Single
        Dim areaBasedRatio As Single
        
        sFound.Copy
        cell.Select
        ActiveSheet.Paste Link:=False
    
        Set sNew = Selection.ShapeRange(1)
        sNew.AlternativeText = "Best Counter"
        
        If cellForSizing Is Nothing Then Set cellForSizing = cell  ' size all icons consistent with the first one.
        
        If bdm.fBestCounterIcons_Large Then
            ' large icons proportional to the area of the cell.
            areaBasedRatio = 0.8 * Sqr((cellForSizing.Height * cellForSizing.width) / (sNew.Height * sNew.width))
        Else
            ' small icons based on the square of height of the cell instead of the whole area of the cell.
            areaBasedRatio = Sqr((cellForSizing.Height * cellForSizing.Height) / (sNew.Height * sNew.width))
        End If

        sNew.ScaleHeight areaBasedRatio, msoFalse
        
        If iRowVis Mod 2 <> 0 Then
            sNew.left = cell.left - 3
        Else
            sNew.left = cell.left + cell.width - sNew.width + 3
        End If
        
        sNew.top = sNew.top - (sNew.Height - cell.Height) / 2
        
        If fSecondBest Then sNew.Fill.PictureEffects.Insert msoEffectPhotocopy

    End If

NoIcon:

End Sub

Sub DisplayChargeMoveAlerts(cell As Range, ByRef cellForSizing As Range, ByRef iRowVis As Integer, bdm As BattleDisplayMode, pkTeam As Pokemon, pkMeta As Pokemon)
    Dim sFound As Shape, sNew As Shape
    Dim areaBasedRatio As Single, scaleFactor As Single
    Dim widthSoFar As Single
    Dim left As Single
    
    On Error GoTo NoImage
    
    If cellForSizing Is Nothing Then Set cellForSizing = cell  ' size all icons consistent with the first one.
    
    If bdm.fChargeMoveAlerts_Shields And pkMeta.cmBest.threat.threatLevel > 3 Then
        Set sFound = FindPokemonShape("Shield")
        If Not (sFound Is Nothing) Then
            sFound.Copy
            cell.Select
            ActiveSheet.Paste Link:=False
            
            Set sNew = Selection.ShapeRange(1)
            sNew.AlternativeText = "Move Alert"
            
            scaleFactor = Min(1.5, Sqr(pkMeta.cmBest.threat.threatLevel) / 5)
            
            areaBasedRatio = scaleFactor * Sqr((cellForSizing.Height * cellForSizing.Height) / (sNew.Height * sNew.width))
            sNew.ScaleHeight areaBasedRatio, msoFalse
            
            widthSoFar = widthSoFar + sNew.width
            
            If iRowVis Mod 2 <> 0 Then
                sNew.left = cell.left + cell.width - widthSoFar - 3
            Else
                sNew.left = cell.left + 3
            End If
            
            sNew.top = sNew.top - (sNew.Height - cell.Height) / 2
        End If
    End If
    
    If bdm.fChargeMoveAlerts_Bolts And pkTeam.cmStrongest.threat.threatLevel > 7 Then
    
If False Then 'Not Ready to do this because we don't have an accurate row position.  cell.Top is wrong.
        cell.Select
        scaleFactor = Min(0.05, Sqr(pkTeam.cmStrongest.threat.threatLevel) / 6)
        
        If iRowVis Mod 2 <> 0 Then
            left = cell.left + cell.width - widthSoFar - (216 * scaleFactor) - 3
        Else
            left = cell.left + widthSoFar + 3
        End If
        
        Set sNew = DrawLightning(left, Selection.top, scaleFactor)
        
        sNew.AlternativeText = "Move Alert"
End If

        Set sFound = FindPokemonShape("Mortal Attack")
        If Not (sFound Is Nothing) Then
            sFound.Copy
            cell.Select
            ActiveSheet.Paste Link:=False
            
            Set sNew = Selection.ShapeRange(1)
            sNew.AlternativeText = "Move Alert"
            
            scaleFactor = Min(1.2, Sqr(pkTeam.cmStrongest.threat.threatLevel) / 6)
            
            areaBasedRatio = scaleFactor * Sqr((cellForSizing.Height * cellForSizing.Height) / (sNew.Height * sNew.width))
            sNew.ScaleHeight areaBasedRatio, msoFalse
            sNew.ScaleWidth areaBasedRatio, msoFalse  'REVIEW - If using a bitmap image, do NOT call ScaleWidth too.
            
            If iRowVis Mod 2 <> 0 Then
                sNew.left = cell.left + cell.width - widthSoFar - sNew.width - 3
            Else
                sNew.left = cell.left + widthSoFar + 3
            End If
            
            sNew.top = sNew.top - (sNew.Height - cell.Height) / 2
        End If

    End If
    
NoImage:

End Sub

Function DrawLightning(posX As Single, posY As Single, sf As Single) As Shape

    Dim sNew As Shape
    
    With ActiveSheet.Shapes.BuildFreeform(msoEditingAuto, 20 * sf + posX, 124 * sf + posY)
        .AddNodes msoSegmentLine, msoEditingAuto, 220 * sf + posX, 0 * sf + posY
        .AddNodes msoSegmentLine, msoEditingAuto, 120 * sf + posX, 103 * sf + posY
        .AddNodes msoSegmentLine, msoEditingAuto, 182 * sf + posX, 105 * sf + posY
        .AddNodes msoSegmentLine, msoEditingAuto, 0 * sf + posX, 210 * sf + posY
        .AddNodes msoSegmentLine, msoEditingAuto, 70 * sf + posX, 120 * sf + posY
        .AddNodes msoSegmentLine, msoEditingAuto, 20 * sf + posX, 124 * sf + posY
        .ConvertToShape.Select
    End With
    
    Set sNew = Selection.ShapeRange(1)
    
    sNew.Line.Visible = msoFalse
    
    With sNew.Fill
        .Visible = msoTrue
        .TwoColorGradient msoGradientHorizontal, 1
        .ForeColor.RGB = RGB(255, 255, 0)
        .BackColor.RGB = RGB(255, 100, 0)
    End With

If False Then
    With sNew.Shadow
        .Type = msoShadow21

        .Style = msoShadowStyleOuterShadow
        .Blur = 2
        .OffsetX = 2
        .OffsetY = 2
        .RotateWithShape = msoFalse
        .ForeColor.RGB = RGB(0, 0, 0)
        .Transparency = 0.8
        .Size = 100
        .Visible = msoTrue
    End With
End If
    
    Set DrawLightning = sNew

End Function

Sub DeleteBestCounterIcons()

    Dim sh As Shape

    For Each sh In ActiveSheet.Shapes
        If sh.AlternativeText = "Best Counter" Or sh.AlternativeText = "Move Alert" Then sh.Delete
    Next sh
    
End Sub

Function BestCounterIconMode() As String
Dim rng As Range

    On Error GoTo NoRange
    
    BestCounterIconMode = "Large Icons" ' default
    BestCounterIconMode = Range("BestCounterIconMode").value ' Add cell on sheet named BestCounterIconMode containing values from _BestCounterIconModes
       
NoRange:
End Function

Sub GetBattleDisplayMode(ByRef bdm As BattleDisplayMode)
    Dim strMode As String, bdmInit As BattleDisplayMode
    
    bdm = bdmInit ' Clear

    bdm.strDisplayMethod = ValidateRange("DisplayMethod").Text
    bdm.strBattleLeague = ValidateRange("BattleLeague").Text
    Set bdm.rngDataBattleLeague = GetBattleLeagueData(bdm.strBattleLeague)
    
    On Error GoTo NoRange1
    
    ' Add cell on sheet named BestCounterIconMode containing values from _BestCounterIconModes
    strMode = Range("BestCounterIconMode").value
    
    Select Case strMode
        Case "Small Icons"
            bdm.fBestCounterIcons = True
        Case "Small Icons w Close Second"
            bdm.fBestCounterIcons = True
            bdm.fBestCounterIcons_CloseSecond = True
        Case "Large Icons"
            bdm.fBestCounterIcons = True
            bdm.fBestCounterIcons_Large = True
        Case "Large Icons w Close Second"
            bdm.fBestCounterIcons = True
            bdm.fBestCounterIcons_Large = True
            bdm.fBestCounterIcons_CloseSecond = True
    End Select
    
NoRange1:
    
    On Error GoTo NoRange2
    
    ' Add cell on sheet named ChargeMoveAlertMode containing values from _ChargeMoveAlertModes
    strMode = Range("ChargeMoveAlertMode").value
    
    Select Case strMode
        Case "Shields Only"
            bdm.fChargeMoveAlerts = True
            bdm.fChargeMoveAlerts_Shields = True
        Case "Bolts Only"
            bdm.fChargeMoveAlerts = True
            bdm.fChargeMoveAlerts_Bolts = True
        Case "Shields and Bolts"
            bdm.fChargeMoveAlerts = True
            bdm.fChargeMoveAlerts_Shields = True
            bdm.fChargeMoveAlerts_Bolts = True
    End Select
       
NoRange2:

End Sub

Sub AddBattleScoreComment(cell As Range, valScore As Single, pkTeam As Pokemon, pkMeta As Pokemon, bdm As BattleDisplayMode)
    
    Dim strText As String
    Dim fTypeEffectivenessBattle As Boolean
    
    ' The comment consists of 4 text blocks, and there must be a blank line between blocks.
    ' BattleScoreFromComment needs to be able to parse this text in order to update the screen without running new battle simulations.
    
    On Error GoTo NoComment
    
    ' If this is changed, must update BattleScoreFromComment
    
    'Text Block 1:  Battle Results
    strText = "Battle Score: " & CStr(valScore) & Chr(10)
    
    If valScore < 400 Then
        strText = strText & pkMeta.strName & " strongly favored in "
    ElseIf valScore < 490 Then
        strText = strText & pkMeta.strName & " favored in "
    ElseIf valScore > 600 Then
        strText = strText & pkTeam.strName & " strongly favored in "
    ElseIf valScore > 510 Then
        strText = strText & pkTeam.strName & " favored in "
    Else
        strText = strText & "Virtual Toss-Up in "
    End If
    
    If fTypeEffectivenessBattle Then
        strText = strText & "Type Effectiveness Battle." & Chr(10)
    Else
        strText = strText & bdm.strBattleLeague & " Battle." & Chr(10)
    End If
    
    strText = strText & Chr(10)
    
    'Text Block 2:  Team Pokemon Performance

    With pkTeam.cm
        strText = strText & pkTeam.strName & str_MoveRecommend & .strMove
        If .strBuffSymbols <> "" Then strText = strText & " (" & .strBuffSymbols & ")"
        strText = strText & Chr(10)
    End With
        
    With pkTeam.cmBest
        If .threat.pctDamage > 0 Then
            strText = strText & "Opportunity Level " & .threat.threatLevel & ":  " & .threat.pctDamage & "% damage by " & .strMove & _
                " after " & .cTurnsToCharge & " turns."
                
            strText = strText & " (" & -.hpPerCharge & " hp)"
                
            If .cTurnsToCharge = pkMeta.cmBest.cTurnsToCharge Then
                If pkTeam.bstat.attCMP > pkMeta.bstat.attCMP + 15 Then
                    strText = strText & " (Attacks First)"
                ElseIf pkTeam.bstat.attCMP > pkMeta.bstat.attCMP Then
                    strText = strText & " (Likely Attacks First)"
                ElseIf pkMeta.bstat.attCMP > pkTeam.bstat.attCMP + 15 Then
                    strText = strText & " (Attacks Second)"
                ElseIf pkMeta.bstat.attCMP > pkTeam.bstat.attCMP Then
                    strText = strText & " (Likely Attacks Second)"
                ElseIf pkTeam.strName = pkMeta.strName Then
                    strText = strText & " (Mirror Match)"
                Else
                    strText = strText & " (CMP winner indeterminate)"
                End If
            End If
                
            strText = strText & Chr(10)
        End If
    End With
    
    With pkTeam.cmStrongest
        If .threat.pctDamage > 0 Then
            strText = strText & "Strongest Opportunity Level " & .threat.threatLevel & ":  " & .threat.pctDamage & "% damage by " & .strMove & _
                " after " & .cTurnsToCharge & " turns."
                
            strText = strText & " (" & -.hpPerCharge & " hp)" & Chr(10)
        End If
    End With
    
    With pkTeam.cmQuickest
        If .threat.pctDamage > 0 Then
            strText = strText & "Fastest Threat Level " & .threat.threatLevel & ":  " & .threat.pctDamage & "% damage by " & .strMove & _
                " after " & .cTurnsToCharge & " turns."
                
            strText = strText & " (" & -.hpPerCharge & " hp)" & Chr(10)
        End If
    End With
    
    With pkTeam.cmBestBuff
        If .threat.pctDamage > 0 Then
            strText = strText & "Alternative Threat Level " & .threat.threatLevel & ":  " & .threat.pctDamage & "% damage by " & .strMove & _
                " after " & .cTurnsToCharge & " turns."
                
            strText = strText & " (" & -.hpPerCharge & " hp)" & Chr(10)
        End If
    End With
    
    If pkTeam.qm.cTurnsToVictory > 0 Then
        strText = strText & "Quick Move: " & pkTeam.qm.strMove & " alone can win the battle in " & pkTeam.qm.cTurnsToVictory & " moves."
        strText = strText & " (" & -pkTeam.qm.hpPerQuick & " hp)" & Chr(10)
    End If
    
    strText = strText & Chr(10)
    
    'Text Block 3:  Meta Pokemon Performance
    
    With pkMeta.cm
        strText = strText & pkMeta.strName & str_MoveRecommend & .strMove
        If .strBuffSymbols <> "" Then strText = strText & " (" & .strBuffSymbols & ")"
        strText = strText & Chr(10)
    End With
    
    With pkMeta.cmBest
        If .threat.pctDamage > 0 Then
            strText = strText & "Threat Level " & .threat.threatLevel & ":  " & .threat.pctDamage & "% damage by " & .strMove & _
                " after " & .cTurnsToCharge & " turns."
                
            strText = strText & " (" & -.hpPerCharge & " hp)" & Chr(10)
        End If
    End With
    
    With pkMeta.cmStrongest
        If .threat.pctDamage > 0 Then
            strText = strText & "Strongest Threat Level " & .threat.threatLevel & ":  " & .threat.pctDamage & "% damage by " & .strMove & _
                " after " & .cTurnsToCharge & " turns."
                
            strText = strText & " (" & -.hpPerCharge & " hp)" & Chr(10)
        End If
    End With
    
    With pkMeta.cmQuickest
        If .threat.pctDamage > 0 Then
            strText = strText & "Fastest Threat Level " & .threat.threatLevel & ":  " & .threat.pctDamage & "% damage by " & .strMove & _
                " after " & .cTurnsToCharge & " turns."
                
            strText = strText & " (" & -.hpPerCharge & " hp)" & Chr(10)
        End If
    End With
    
    With pkMeta.cmBestBuff
        If .threat.pctDamage > 0 Then
            strText = strText & "Alternative Threat Level " & .threat.threatLevel & ":  " & .threat.pctDamage & "% damage by " & .strMove & _
                " after " & .cTurnsToCharge & " turns."
                
            strText = strText & " (" & -.hpPerCharge & " hp)" & Chr(10)
        End If
    End With
    
    If pkMeta.qm.cTurnsToVictory > 0 Then
        strText = strText & "Quick Move: " & pkMeta.qm.strMove & " alone can win the battle in " & pkMeta.qm.cTurnsToVictory & " moves."
        strText = strText & " (" & -pkMeta.qm.hpPerQuick & " hp)" & Chr(10)
    End If
    
    strText = strText & Chr(10)
    
    'Text Block 4:  CSV and Additional Information
    
    strText = strText & pkTeam.csv
    If pkTeam.ivs.csvIV <> "" Then strText = strText & " [" & pkTeam.ivs.csvIV & "]"
    
    strText = strText & Chr(10)
    
    With pkTeam.bstat
    If .level > 0 Then
        strText = strText & "Level: " & .level & ", CP: " & .cp & _
            ", CMP: " & RoundDown(.attCMP, 1) & ", Attack: " & RoundDown(.attInit, 1) & ", Defense: " & RoundDown(.defInit, 1) & ", HP: " & .hp _
            & ", SProd: " & RoundDown((.attInit * .defInit * .hp) / 1000000, 2) & "M"
    End If
    End With
    
    strText = strText & Chr(10) & Chr(10)
    
    strText = strText & pkMeta.csv
    If pkMeta.ivs.csvIV <> "" Then strText = strText & " [" & pkMeta.ivs.csvIV & "]"
    
    strText = strText & Chr(10)
    
    With pkMeta.bstat
    If .level > 0 Then
        strText = strText & "Level: " & .level & ", CP: " & .cp & _
            ", CMP: " & RoundDown(.attCMP, 1) & ", Attack: " & RoundDown(.attInit, 1) & ", Defense: " & RoundDown(.defInit, 1) & ", HP: " & .hp _
            & ", SProd: " & RoundDown((.attInit * .defInit * .hp) / 1000000, 2) & "M"
    End If
    End With
    
    strText = strText & Chr(10) & Chr(10)
    
    If (Not fTypeEffectivenessBattle) Then
        ' add hyperlink to pvpoke battle simulation
        Dim cpMaxBattleLeague As Integer
        cpMaxBattleLeague = bdm.rngDataBattleLeague.Cells(1, blData_MaxCp)
        
        If cpMaxBattleLeague > 0 Then
            strText = strText & "See Also: https://pvpoke.com/battle/" & cpMaxBattleLeague & "/" & StrForPvPoke(pkTeam.strName) & "/" & StrForPvPoke(pkMeta.strName) & "/11/" & Chr(10)
        End If
        
        strText = strText & Chr(10)
    End If
    
    With cell
    
    If Not (.CommentThreaded Is Nothing) Then .CommentThreaded.Delete
    
    If .Comment Is Nothing Then
        .AddComment strText
        .Comment.Shape.TextFrame.AutoSize = True
    Else
        .Comment.Text strText
    End If

    End With
    
    Exit Sub
    
NoComment:
    breakPoint = 0

End Sub

Function BattleScoreFromComment(cellUpd As Range, pkTeam As Pokemon, pkMeta As Pokemon) As Single
' Get the values back out of a comment

    Dim strText As String, strExtra As String
    Dim strGroup1 As String, strGroup2 As String, strGroup3 As String, strGroup4 As String, strGroup5 As String
    Dim strScoreParse As String, strTeamParse As String, strMetaParse As String
    Dim strTeamNumbersParse As String, strMetaNumbersParse As String
    Dim strTeamNumbersParse2 As String, strMetaNumbersParse2 As String, strGroupSep As String
    Dim strCsv As String
    
    On Error GoTo NoComment
    
    strText = cellUpd.Comment.Text
    strGroupSep = Chr(10) & Chr(10)
    
    Call Parse5Substrings(strText, strGroupSep, strGroup1, strGroup2, strGroup3, strGroup4, strGroup5)
    
    strScoreParse = StrTrimBeforeAndAfter(strGroup1, ":", Chr(10))
    BattleScoreFromComment = CDec(strScoreParse)
    
    Call Parse4Substrings(strGroup2, Chr(10), strTeamParse, strTeamNumbersParse, strTeamNumbersParse2, strExtra)
    Call Parse4Substrings(strGroup3, Chr(10), strMetaParse, strMetaNumbersParse, strMetaNumbersParse2, strExtra)
    
    strCsv = ParseSubstring(strGroup4, 1, Chr(10))
    pkTeam.csv = ParseSubstring(strCsv, 1, " [")
    pkTeam.ivs.csvIV = ParseSubstringBetween(strCsv, " [", "]")
    
    strCsv = ParseSubstring(strGroup5, 1, Chr(10))
    pkMeta.csv = ParseSubstring(strCsv, 1, " [")
    pkMeta.ivs.csvIV = ParseSubstringBetween(strCsv, " [", "]")

    pkTeam.strName = StrTrimBeforeAndAfter(strTeamParse, "", str_MoveRecommend)
    pkTeam.cm.strMove = StrTrimBeforeAndAfter(strTeamParse, str_MoveRecommend, " (")
    pkMeta.strName = StrTrimBeforeAndAfter(strMetaParse, "", str_MoveRecommend)
    pkMeta.cm.strMove = StrTrimBeforeAndAfter(strMetaParse, str_MoveRecommend, " (")

    If strTeamNumbersParse <> "" Then
        strTeamNumbersParse = StrTrimBeforeAndAfter(strTeamNumbersParse, "Level ", ":")
        pkTeam.cmBest.threat.threatLevel = CDec(strTeamNumbersParse)
        pkTeam.cmStrongest.threat.threatLevel = pkTeam.cmBest.threat.threatLevel ' the same if next insruction is missing.
    End If
    
    If left(strTeamNumbersParse2, 9) = "Strongest" Then
        strTeamNumbersParse2 = StrTrimBeforeAndAfter(strTeamNumbersParse2, "Level ", ":")
        pkTeam.cmStrongest.threat.threatLevel = CDec(strTeamNumbersParse2)
    End If
    
    If strMetaNumbersParse <> "" Then
        strMetaNumbersParse = StrTrimBeforeAndAfter(strMetaNumbersParse, "Level ", ":")
        pkMeta.cmBest.threat.threatLevel = CDec(strMetaNumbersParse)
        pkMeta.cmStrongest.threat.threatLevel = pkMeta.cmBest.threat.threatLevel ' the same if next insruction is missing.
    End If
    
    If left(strMetaNumbersParse2, 9) = "Strongest" Then
        strMetaNumbersParse2 = StrTrimBeforeAndAfter(strMetaNumbersParse2, "Level ", ":")
        pkMeta.cmStrongest.threat.threatLevel = CDec(strMetaNumbersParse2)
    End If
    
    Exit Function

NoComment:
    BattleScoreFromComment = -1

End Function
Sub AppendToComment(cellUpd As Range, strAppend As String)
    
    Dim strComment As String
    On Error GoTo NoComment
    
    With cellUpd
    
    If Not (.Comment Is Nothing) Then
        .Comment.Text .Comment.Text & strAppend & Chr(10)
    End If
    
    End With
    
NoComment:
End Sub

Sub DeleteComment(cellUpd As Range)
    
    If Not (cellUpd.Comment Is Nothing) Then cellUpd.Comment.Delete
 
End Sub

Sub TrimCommentAtSymbol(cellUpd As Range, strSymbol As String)
    Dim strText As String, iSymbol As Integer
    
    If cellUpd.Comment Is Nothing Then Exit Sub
    On Error GoTo NoComment
    
    strText = cellUpd.Comment.Text
    iSymbol = InStr(strText, strSymbol)
    
    If iSymbol > 1 Then
        strText = left(strText, iSymbol - 1)
        cellUpd.Comment.Text strText
    End If
    
NoComment:
End Sub



Sub MarkPokemonBlock(ByRef bma As BestMoveAnalysis, ByVal strTeam As String, ByVal rngUpdate As Range, ByVal iRow As Integer, _
ByVal iCol As Integer, ByVal valScore As Single, ByVal rngBlockBorders)
 
    If strTeam = bma.strSamePokemon Then
        If valScore > bma.valScoreBestSamePokemon Then
            bma.valScoreBestSamePokemon = valScore
            bma.iColBestSamePokemon = iCol
        End If
        
        If valScore < bma.valScoreWorstSamePokemon Then
            bma.valScoreWorstSamePokemon = valScore
        End If
        
        bma.iColLastSamePokemon = iCol
        
    Else
        If bma.iColSamePokemon > 0 Then
        
            ' We are done processing a block of columns representing a single pokemon
            
            bma.cSamePokemonBlocks = bma.cSamePokemonBlocks + 1
            If bma.cSamePokemonBlocks Mod 2 = 0 And Not (rngBlockBorders Is Nothing) Then
                Dim rng As Range
                ' Draw a border around every other pokemon group.
                
                Set rng = CombineRanges(rngBlockBorders.Columns(bma.iColSamePokemon), rngBlockBorders.Columns(bma.iColLastSamePokemon))
                Call DrawOutsideBorder(rng, color_PokemonGroupBorder)
                
            End If
            
            If bma.iColLastSamePokemon = bma.iColSamePokemon Then
                Call MarkCellBestScore(rngUpdate.Cells(iRow, bma.iColSamePokemon), False)
            Else
                While bma.iColSamePokemon <= bma.iColLastSamePokemon
                    Call MarkCellBestScore(rngUpdate.Cells(iRow, bma.iColSamePokemon), _
                        bma.iColSamePokemon = bma.iColBestSamePokemon And bma.valScoreBestSamePokemon > bma.valScoreWorstSamePokemon)
                    bma.iColSamePokemon = bma.iColSamePokemon + 1
                Wend
            End If
        End If
        
        bma.iColSamePokemon = iCol
        bma.iColBestSamePokemon = iCol
        bma.iColLastSamePokemon = iCol
        
        bma.strSamePokemon = strTeam
        bma.valScoreBestSamePokemon = valScore
        bma.valScoreWorstSamePokemon = valScore
    End If

End Sub
            

Sub MarkCellBestScore(cell As Range, fBest As Boolean)
    If cell.Font.Italic <> fBest Then cell.Font.Italic = fBest
End Sub

Sub MarkCellPokemonBenched(cell As Range, fPokemonBenched As Boolean)
    If cell.Font.Strikethrough <> fPokemonBenched Then cell.Font.Strikethrough = fPokemonBenched
End Sub

Sub DrawOutsideBorder(rng As Range, colorSet As Double)

    On Error GoTo NoBorder

    With rng.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .color = colorSet
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With rng.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .color = colorSet
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With rng.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .color = colorSet
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With rng.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .color = colorSet
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    
NoBorder:

End Sub

Sub EraseBorders(rng As Range)
    On Error GoTo NoBorder

    rng.Borders(xlDiagonalDown).LineStyle = xlNone
    rng.Borders(xlDiagonalUp).LineStyle = xlNone
    rng.Borders(xlEdgeLeft).LineStyle = xlNone
    rng.Borders(xlEdgeTop).LineStyle = xlNone
    rng.Borders(xlEdgeBottom).LineStyle = xlNone
    rng.Borders(xlEdgeRight).LineStyle = xlNone
    rng.Borders(xlInsideVertical).LineStyle = xlNone
    rng.Borders(xlInsideHorizontal).LineStyle = xlNone
    
NoBorder:
End Sub

Function CombineRanges(rng1 As Range, rng2 As Range) As Range
    Dim strAddress As String
    
    Set CombineRanges = Nothing
    On Error GoTo NoRange
    
    strAddress = rng1.Address & ":" & rng2.Address
    Set CombineRanges = Range(strAddress)

NoRange:
End Function

Sub SetCellText(cell As Range, strText As String)
'    If cell.Text <> strText Then cell.value = strText
    cell.value = strText ' must force stats recalc for Best Counter, etc.
End Sub

Sub SetCellValue(cell As Range, value As Single)
'    If cell.value <> value Then cell.value = value
    cell.value = value
End Sub

Sub TimeCheck(ByRef timerStart As Single, Optional tooLong As Single = 0.05)
    Dim timerNow As Single
    
    timerNow = Timer - timerStart
    
    If timerNow > 15 Then Exit Sub ' we probably broke into the debugger.
    
    If timerNow > tooLong Then
    'uncomment when profiling
'        MsgBox ("Too Slow!  Press Ctrl+Break to Debug.  Comment out MsgBox if TimeCheck undesired.")
        timerStart = Timer
    End If
    
End Sub

Function StrForPvPoke(strName As String) As String
    Dim ichParen As Integer
    Dim strPvPoke As String
    
    strPvPoke = StrConv(strName, vbLowerCase)
    
    ichParen = InStr(strName, " (")
    If ichParen > 0 Then
        strPvPoke = left(strPvPoke, ichParen - 1) & "_" & StrTrimBeforeAndAfter(strPvPoke, "(", ")")
    End If
    
    StrForPvPoke = strPvPoke

End Function



