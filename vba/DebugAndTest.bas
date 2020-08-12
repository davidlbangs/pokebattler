Attribute VB_Name = "DebugAndTest"
Option Explicit
' Macros not used in the project at all, but which can come in handy during the development process to develop and test code or spreadsheets.


Function FindPokemonShape(strPokemon As String) As Shape
    Dim s As Shape
    Dim ws As Worksheet
    
On Error GoTo NoShape

    Set ws = Application.Worksheets("Pokemon Data")

On Error GoTo LookAtAltText

    If strPokemon = "" Then GoTo NoShape    'make sure we don't get wrong shape by being vague.
    
    Set FindPokemonShape = ws.Shapes(strPokemon)
    Exit Function
    
LookAtAltText:
On Error GoTo NoShape
    
    ' look the hard way.
    For Each s In ws.Shapes
        If s.AlternativeText = strPokemon Then
            Set FindPokemonShape = s
            Exit Function
        End If
    Next s
        
NoShape:
    Set FindPokemonShape = Nothing
    
End Function

Sub DeleteHyper()
    Dim s As Shape
    
On Error GoTo NoShape

    For Each s In ActiveSheet
        If s.Hyperlink.Address <> "" Then s.Hyperlink.Delete
    Next s
    
NoShape:
    
End Sub

Sub InsertPokemonShape()
    Dim sFound As Shape, sNew As Shape
    Set sFound = FindPokemonShape(Selection.Text)
    
    If Not (sFound Is Nothing) Then
        Set sNew = sFound.Duplicate
        sNew.AlternativeText = "Duplicate"
        sNew.left = Selection.left + 50
        sNew.top = Selection.top + (Selection.Height - sNew.Height) / 2
    End If
    
End Sub

Sub SelectPokemonShape()
    Dim sFound As Shape, sNew As Shape
    Set sFound = FindPokemonShape(Selection.Text)
    
    If IsObject(sFound) Then
        sFound.Select
        ActiveWindow.ScrollIntoView sFound.left, sFound.top, sFound.width, sFound.Height, True
    End If
    
End Sub

Function FindLostPokemonShape(strPokemon As String, iFind As Integer) As Shape
    Dim s As Shape
    Dim ws As Worksheet
    
On Error GoTo NoShape
    Set ws = Application.Worksheets("Pokemon Data")

    For Each s In ws.Shapes
        If InStr(s.Name, strPokemon) > 0 Then
            If iFind <= 1 Then
                Set FindLostPokemonShape = s
                Exit Function
            Else
                iFind = iFind - 1
            End If
        End If
    Next s
    
NoShape:
    
    Set FindLostPokemonShape = Nothing
    
End Function

Sub RetreivePokemonShape()
    Dim sFound As Shape, sNew As Shape
    Dim leftSave As Double, topSave As Double
    Set sFound = FindPokemonShape(Selection.Text)
    
    If Not (sFound Is Nothing) Then
        leftSave = sFound.left
        topSave = sFound.top
        
        sFound.left = Selection.left
        sFound.top = Selection.top
        
        'Edit object manually then put it back where it was
        
        sFound.left = leftSave
        sFound.top = topSave
    End If
    
End Sub

Function LacksShape(str As String) As Boolean
    Dim sFound As Shape
    Set sFound = FindPokemonShape(str)
    
    LacksShape = (sFound Is Nothing)

End Function

Sub PastePokemonShapes()
    Dim sFound As Shape, sNew As Shape
    Dim selSave As Range, cell As Range
    
    Set selSave = Selection.Cells
    
    For Each cell In selSave
    
        Set sFound = FindPokemonShape(cell.Text)
        
        If Not (sFound Is Nothing) Then
            sFound.Copy
            ActiveSheet.Paste Link:=False
            
            Set sNew = Selection.ShapeRange(1)
            sNew.AlternativeText = "Duplicate"
            sNew.ScaleHeight cell.Height / sNew.Height, msoFalse
            
            sNew.left = cell.left + cell.width - sNew.width
            sNew.top = cell.top + (cell.Height - sNew.Height) / 2
        End If
    Next cell
    
End Sub

Sub ReplaceImage()
Attribute ReplaceImage.VB_ProcData.VB_Invoke_Func = "d\n14"
    Dim sFound As Shape, sNew As Shape
    Dim dest As Range, cellRef As Range, cellsVisible As Range
    Dim areaBasedRatio As Single
    Dim top As Single, left As Single, width As Single, heigth As Single
    Dim strName As String, strAlt As String
    
    ' Select Old Object
    
    Set sFound = Selection.ShapeRange(1)
    top = sFound.top
    left = sFound.left
    strAlt = sFound.AlternativeText
    strName = sFound.Name
    sFound.Delete
    
    ' Select New Object
    
    Set cellRef = Range("I1")
    Set cellsVisible = ActiveWindow.VisibleRange
    Set dest = Range("I1")
    
    Set dest = dest.Columns(1).EntireColumn
    Set dest = dest.Cells(cellsVisible.Row + 10, 1)
    
    dest.Select
    ActiveSheet.Paste Link:=False
    Set sNew = FindLostPokemonShape("Picture", 1)
    
    'select and crop
    
    sNew.AlternativeText = strAlt
    sNew.Name = strName
    areaBasedRatio = Sqr((cellRef.Height * cellRef.width) / (sNew.Height * sNew.width)) * 0.7
    sNew.ScaleHeight areaBasedRatio, msoFalse
    sNew.top = top
    sNew.left = left
    


End Sub

Sub DeleteDuplicateImages()

    Dim sh As Shape

    For Each sh In ActiveSheet.Shapes
        If sh.AlternativeText = "Duplicate" Then sh.Delete
    Next sh
    
End Sub

Sub ResetImagesX()

    Dim sh As Shape

    For Each sh In ActiveSheet.Shapes
        If sh.AlternativeText <> "" Then
            sh.Name = sh.AlternativeText
            sh.ScaleHeight 1, msoTrue
            sh.ScaleWidth 1, msoTrue
        End If
    Next sh
    
End Sub

Sub ResetImagesX1()
    Dim str As String

    Dim sh As Shape

    For Each sh In ActiveSheet.Shapes
        str = left(sh.AlternativeText, 5)
        If str = "Updat" Then
            sh.Delete
        End If
    Next sh
    
End Sub



Sub BringHome()

    Dim sh As Shape

    For Each sh In ActiveSheet.Shapes
        If sh.left > 2500 Or sh.top > 13000 Then
            sh.left = 100
            sh.Right = 100
        End If
    Next sh
    
End Sub

Sub ArrangeImages()
    Dim sFound As Shape
    Dim rngTable As Range
    Dim cell As Range
    Dim xCenter As Single
    Dim areaBasedRatio As Single
    Dim iRow As Integer, cRow As Integer
    Dim dest As Range
    Dim breakPoint As Integer
    
    
    Set rngTable = Range("_PokemonTable")
    Set dest = Range("PokemonImages")
    
    For Each cell In rngTable.Columns(1).Cells
        Set sFound = FindPokemonShape(cell.Text)
        
        If (cell.Text = "Genesect (Shock)") Then
            breakPoint = 0
        End If
        
        If Not (sFound Is Nothing) Then
        
            areaBasedRatio = Sqr((dest.Height * dest.width) / (sFound.Height * sFound.width)) * 0.7
            sFound.ScaleHeight areaBasedRatio, msoFalse
            
            xCenter = dest.left + ((cell.Row Mod 8) + 1) * (dest.width / 9)
            
            
            sFound.left = xCenter - sFound.width / 2
            sFound.top = cell.top - (sFound.Height - cell.Height) / 2

        End If
        
    Next cell
    
    'Arrange pokemon representing types
    
    Set rngTable = Range("J4:J21")
    Set dest = Range("L1")
    
    For Each cell In rngTable.Columns(1).Cells
        Set sFound = FindPokemonShape(cell.Text)
        
        If (sFound Is Nothing) Then
            breakPoint = 0
        End If
        
        If Not (sFound Is Nothing) Then
        
            areaBasedRatio = Sqr((dest.Height * dest.width) / (sFound.Height * sFound.width)) * 0.7
            sFound.ScaleHeight areaBasedRatio, msoFalse
            
            xCenter = dest.left + ((cell.Row Mod 8) + 1) * (dest.width / 9)
            
            
            sFound.left = xCenter - sFound.width / 2
            sFound.top = cell.top - (sFound.Height - cell.Height) / 2

        End If
        
    Next cell

    ' break point
    
    Set dest = Range("L1")
    
End Sub


Sub BeautifySelCsv()
    Dim strUgly As String, strPretty As String
    Dim iNextMove As Integer, strNextMove As String
    Dim cell As Range
    
    For Each cell In Selection
        strUgly = cell.Text
        
        strPretty = ParsePokemonName(strUgly) + ", " + ParseMoveName(strUgly, 1) + ", " + ParseMoveName(strUgly, 2)
        
        If Len(strPretty) < 10 Then Exit Sub ' not real csv
        
        iNextMove = 3
        strNextMove = ParseMoveName(strUgly, iNextMove)
        
        While strNextMove <> ""
            strPretty = strPretty & ", " & strNextMove
            iNextMove = iNextMove + 1
            strNextMove = ParseMoveName(strUgly, iNextMove)
        Wend
        
        cell.value = strPretty
        
    Next cell

End Sub


Sub ResetComments()
'Update 20141110
Dim pComment As Comment
For Each pComment In Application.ActiveSheet.Comments
   pComment.Shape.top = pComment.Parent.top + pComment.Parent.Height + 5
   pComment.Shape.left = pComment.Parent.left + 5
Next
End Sub

Private Sub RecalcNow()
    If (Application.Calculation = xlCalculationManual) Then
        ActiveSheet.Calculate
    End If
End Sub

'In case Excel modes accidentally got left off during debugging.
Private Sub FixUpdateMode()
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
End Sub


Sub FillSelectionColor()
Dim cell As Range
    For Each cell In Selection
        cell.Interior.color = cell.Text
    Next cell
End Sub

Function StringNBS(str As String) As String
    StringNBS = Replace(str, " ", Chr(160))
End Function

Function StrMoveAlert(levelThreat As Integer, levelOpportunity As Integer) As String
    Dim rngSymbols As Range, strSymbols As String
    
    Set rngSymbols = Range("_NumberSymbols")
    
    If Not (rngSymbols Is Nothing) Then
        strSymbols = rngSymbols.Text
        
        StrMoveAlert = Mid(strSymbols, MinMaxI(levelThreat, 1, 10), 1) + Mid(strSymbols, MinMaxI(levelOpportunity, 1, 10) + 10, 1)
    End If

End Function

Function StrFormatShieldText(str As String, levelThreat As Integer, levelOpportunity As Integer, fRight As Boolean) As String
    Dim strText As String

    If fRight Then
        strText = Chr(160) & " " & StringNBS(str) & " " & StrMoveAlert(levelThreat, levelOpportunity)
    Else
        strText = StrMoveAlert(levelThreat, levelOpportunity) & " " & StringNBS(str) & " " & Chr(160)
    End If
    
    StrFormatShieldText = strText

End Function


Sub TestMoveAlert()
    Dim sFound As Shape, sNew As Shape
    Dim areaBasedRatio As Single, scaleFactor As Single
    Dim widthSoFar As Single
    Dim left As Single
        
        scaleFactor = Min(0.05, Sqr(10) / 6)
        
        Set sNew = DrawLightning(Selection.left + 3, Selection.Cells(1, 1).top, 1)
        
        sNew.AlternativeText = "Move Alert"

End Sub


