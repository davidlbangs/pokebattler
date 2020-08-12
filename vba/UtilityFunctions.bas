Attribute VB_Name = "UtilityFunctions"
Option Explicit

'Generally useful functions which could be used in any macro project.  Not related to Pokemon at all.

Function RoundDown(val As Single, Optional cDigit As Integer = 0) As Single
    RoundDown = Application.WorksheetFunction.RoundDown(val, cDigit)
End Function

Function RoundUp(val As Single, Optional cDigit As Integer = 0) As Single
    RoundUp = Application.WorksheetFunction.RoundUp(val, cDigit)
End Function

Function Round(val As Single, Optional cDigit As Integer = 0) As Single
    Round = Application.WorksheetFunction.Round(val, cDigit)
End Function

Function Min(val1 As Single, val2 As Single) As Single
    If val1 > val2 Then Min = val2 Else Min = val1
End Function

Function Max(val1 As Single, val2 As Single) As Single
    If val1 > val2 Then Max = val1 Else Max = val2
End Function

Function MinI(val1 As Integer, val2 As Integer) As Integer
    If val1 > val2 Then MinI = val2 Else MinI = val1
End Function

Function MaxI(val1 As Integer, val2 As Integer) As Integer
    If val1 > val2 Then MaxI = val1 Else MaxI = val2
End Function

Function MinMax(val1 As Single, valMin As Single, valMax As Single) As Single
    If val1 > valMax Then
        MinMax = valMax
    ElseIf val1 < valMin Then
        MinMax = valMin
    Else
        MinMax = val1
    End If
End Function

Function MinMaxI(val1 As Integer, valMin As Integer, valMax As Integer) As Integer
    If val1 > valMax Then
        MinMaxI = valMax
    ElseIf val1 < valMin Then
        MinMaxI = valMin
    Else
        MinMaxI = val1
    End If
End Function

Function Average(val1 As Single, val2 As Single) As Single
    Average = (val1 + val2) / 2
End Function

Function WeightedAverage(stat1 As Single, stat2 As Single, weightStat1 As Single) As Single
' weight should be a value between zero and one, inlusive.  0.5 makes it an average, 1 makes it all stat1, 0 makes it all stat2
    WeightedAverage = stat1 * weightStat1 + stat2 * (1 - weightStat1)
End Function

Function NotNothing(object As Variant) As Boolean
    NotNothing = Not object Is Nothing
End Function


Private Function CellInNamedRange(ByRef cellTarget As Range, strNamedRange As String) As Boolean
    Dim rangeNamedRange As Range
    
    On Error GoTo Finished
    
    CellInNamedRange = False
    Set rangeNamedRange = cellTarget.Worksheet.Range(strNamedRange)
    
    If RangesIntersect(cellTarget, rangeNamedRange) Then
        CellInNamedRange = True
    End If
Finished:
End Function

Function RangesIntersect(range1 As Range, range2 As Range) As Boolean
    On Error GoTo NoRanges
    
    RangesIntersect = Not (Intersect(range1, range2) Is Nothing)
    Exit Function

NoRanges:
    RangesIntersect = False
End Function

Function ValidateRange(strRange As String, Optional strMessage As String = "Set the named range %") As Range
' Do not call this function in any loop. Nobody likes repeated error messages.
Dim rng As Range

On Error GoTo NoRange
    Set ValidateRange = Range(strRange)
    Exit Function
    
NoRange:
    MsgBox Replace(strMessage, "%", strRange)
    Exit Function

End Function


Function ParseSubstring(ByVal str As String, ByVal iSubstring As Integer, ByVal strSep As String) As String

    Dim ichSeparator As Integer, ichSubstring As Integer, strSubstring As String
    
    ' strSubstring = ""
    ichSubstring = 1
    ichSeparator = InStr(str, strSep)
    
    While iSubstring > 1
        If ichSeparator = 0 Then GoTo Done ' not found
        
        ichSubstring = ichSeparator + Len(strSep)
        ichSeparator = InStr(ichSubstring, str, strSep)
        iSubstring = iSubstring - 1
    Wend
    
    If ichSeparator > 0 Then strSubstring = left(str, ichSeparator - 1) Else strSubstring = str
    If ichSubstring > 1 Then strSubstring = Mid(strSubstring, ichSubstring)
    
Done:
    
    ParseSubstring = Trim(strSubstring)

End Function

Sub Parse4Substrings(ByVal str As String, ByVal strSep As String, ByRef str1 As String, ByRef str2 As String, ByRef str3 As String, ByRef str4 As String)
    Dim ichSep As Integer

    str1 = str
    ichSep = InStr(str1, strSep)
    If ichSep = 0 Then Exit Sub
    
    str2 = Mid(str1, ichSep + Len(strSep))
    str1 = left(str1, ichSep - 1)
    
    ichSep = InStr(str2, strSep)
    If ichSep = 0 Then Exit Sub
    
    str3 = Mid(str2, ichSep + Len(strSep))
    str2 = left(str2, ichSep - 1)
    
    ichSep = InStr(str3, strSep)
    If ichSep = 0 Then Exit Sub
    
    str4 = Mid(str3, ichSep + Len(strSep))
    str3 = left(str3, ichSep - 1)
    
Done:
    
End Sub

Sub Parse5Substrings(ByVal str As String, ByVal strSep As String, ByRef str1 As String, ByRef str2 As String, ByRef str3 As String, ByRef str4 As String, ByRef str5 As String)
    Dim ichSep As Integer

    str1 = str
    ichSep = InStr(str1, strSep)
    If ichSep = 0 Then Exit Sub
    
    str2 = Mid(str1, ichSep + Len(strSep))
    str1 = left(str1, ichSep - 1)
    
    ichSep = InStr(str2, strSep)
    If ichSep = 0 Then Exit Sub
    
    str3 = Mid(str2, ichSep + Len(strSep))
    str2 = left(str2, ichSep - 1)
    
    ichSep = InStr(str3, strSep)
    If ichSep = 0 Then Exit Sub
    
    str4 = Mid(str3, ichSep + Len(strSep))
    str3 = left(str3, ichSep - 1)
    
    ichSep = InStr(str4, strSep)
    If ichSep = 0 Then Exit Sub
    
    str5 = Mid(str4, ichSep + Len(strSep))
    str4 = left(str4, ichSep - 1)
    
Done:
    
End Sub

Function ParseSubstringBetween(ByVal str As String, ByVal strBefore As String, ByVal strAfter As String) As String
    Dim ichSubstring As Integer, cchSubstring As Integer
    Dim strSubstring
    
    'strBefore and strAfter both must be present in the string in order to find a substring.
    
    ParseSubstringBetween = ""
    
    ichSubstring = InStr(str, strBefore)
    If ichSubstring > 0 Then
        ichSubstring = ichSubstring + Len(strBefore)
        cchSubstring = InStr(ichSubstring, str, strAfter) - ichSubstring
        If cchSubstring > 0 Then
            ParseSubstringBetween = Trim(Mid(str, ichSubstring, cchSubstring))
        End If
    End If
End Function

Function StrTrimBeforeAndAfter(ByVal str As String, ByVal strBefore As String, ByVal strAfter As String) As String
    Dim iBefore As Integer, iAfter As Integer
    Dim strTrim
    
    ' Neither strBefore nor strAfter are required delimeters.  If present, trim the beginning or the end to these limits.
    
    strTrim = str
    
    If strBefore <> "" Then
        iBefore = InStr(strTrim, strBefore)
        If iBefore > 0 Then strTrim = Mid(strTrim, iBefore + Len(strBefore))
    End If
    
    If strAfter <> "" Then
        iAfter = InStr(strTrim, strAfter)
        If iAfter > 0 Then strTrim = left(strTrim, iAfter - 1)
    End If
    
    StrTrimBeforeAndAfter = Trim(strTrim)
    
End Function



