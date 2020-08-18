Attribute VB_Name = "BattleData"
' Pokemon Go Battle Planner
' (c) 2020 David Bangs.  All rights reserved
'
' A tool to help PVP particants plan an ideal team and to guide them through battles by providing a heads up view of
' best moves and matchups.

Option Explicit
Option Compare Text

'_PokemonTable
Public Const pkData_Pokemon As Integer = 1
Public Const pkData_Type1 As Integer = 2
Public Const pkData_Type2 As Integer = 3
Public Const pkData_Category As Integer = 4
Public Const pkData_Attack As Integer = 5
Public Const pkData_Defense As Integer = 6
Public Const pkData_Stamina As Integer = 7
Public Const pkData_MaxCp As Integer = 8

'_BattleLeagueTable
Public Const blData_League As Integer = 1
Public Const blData_MaxCp As Integer = 2
Public Const blData_Restriction As Integer = 3

Public Const factor_PVP = (0.5) * (1.3)

Public Type QuickMove
    strMove As String
    strType As String
    rngData As Range
    
    ' battle specific info from BattleScoreCore
    factorMult As Single
    cTurnsToQuick As Single
    cTurnsToVictory As Single ' number of turns to defeat this oponent using only quick moves
    dmgQuick As Single ' Damage per move, always pre-buff
    hpPerQuick As Integer 'Oponent hp taken per charge move, always pre-buff. An integer
    
    dptQuick As Single  ' damage per turn to specific opponent.  May be modified by stat changing attacks.
    hpptQuick As Single ' hp points taken per turn rounded such that each attack takes an integer number of hp
    dptQuickInit As Single ' original value before any stat changing attacks.
    
End Type

Public Type MoveThreat
    pctDamage As Integer ' % of necessary damage done per charge move
    threatLevel As Integer ' Assessment of risk of this attack from 1 to 10 for analysis.
End Type

Public Type ChargeMove
    strMove As String
    strType As String
    rngData As Range
    
    ' Calculations from CalChargeMoveStats.  Specific to battle oponent.
    factorMult As Single
    factorTime As Single
    factorBuff As Single
    dmgCharge As Single ' Damage per move, always pre-buff and pre-time factor.
    hpPerCharge As Integer 'Oponent hp taken per charge move, always pre-buff and pre-time factor. An integer
    
    dptCharge As Single
    hpptCharge As Single ' hp points taken per turn rounded such that each attack takes an integer number of hp
    dptChargeInit As Single 'Damage dealt by charge move undiluted by time factor or any stat changing buff or debuff.
    
    

    cTurnsToCharge As Single
    cTurnsToVictory As Single   ' estimate of number of moves to win battle using this move.  Used for buff move planning.
    
    hpptFirstMoveAdvantage As Single ' Portion of hpptCharge that was awarded by RewardFirstMoveAdvantage
    
    'Info from AdjustForBuff
    strBuffSymbols As String
    
    'Calculation from DetermineMoveThreat
    threat As MoveThreat

End Type

Public Type IndividualValues ' Known as the pokemons iv's.
    csvIV As String    ' string that was used to set ivs differently from default.
    Attack As Integer ' ivs have values from 0 to 15 for Attack, Defense and Stamina
    Defense As Integer
    Stamina As Integer
    levelMax As Single  ' maximum level assignable. 40, or 41 for best friend, or other if level is known or level 40 isn't likely.
End Type

Public Type BattleStats
    cp As Single
    level As Single
    
    attCMP As Single    'attack stat before Shadow boost was applied, used for determining CMP winner.
    attInit As Single   'attack stat at start of battle.
    att As Single       'attack stat as modified by stat changing attacks
    
    defInit As Single   'defense stat at start of battle.
    def As Single       'defense stat as modified by stat changing attacks
    hp As Integer       'stamina stat, must be a whole number.
End Type

Public Type Pokemon
    ' fields filled in by InitPokemon
    csv As String
    strName As String       'version of name for display
    strNameData As String   'version of name needed to look up data in data table
    rngData As Range 'range reference to _PokemonDataTable
    
    fShadow As Boolean     'is this a shadow Pokemon?
    strType1 As String
    strType2 As String
    qm As QuickMove
    
    fInvalid As Boolean ' errors prevent a battle simulation
    fInvalidPokemon As Boolean
    fInvalidQuickMove As Boolean
    fInvalidChargeMove As Boolean
    fMultipleChargeMoves As Boolean
    fLegendaryOrMythical As Boolean
    
    ' fields filled in by QualifyPokemon
    
    fQualified As Boolean ' Qualified for battle.
    ivs As IndividualValues
    bstat As BattleStats
    
    ' fields filled in by or under BattleScoreCore

    ' BattleScoreCore / DetermineBestChargeMove
    
    cmBest As ChargeMove
    cmBestBuff As ChargeMove
    cmStrongest As ChargeMove ' Charge most worthy of blocking.
    cmQuickest As ChargeMove ' First to fire
    
    ' BattleScoreCore / AdjustForBuff
    
    cm As ChargeMove ' Charge move that reflects battle score.  May be a composit move, with a list of moves in the name, _
            and stats reflect buffs in battle. For real initial stats of real moves, see above.

End Type

Function GetQuickMoveData(strQuickMove As String) As Range
    Dim rngTable As Range
    Dim iQuickMove As Integer
    
    If strQuickMove <> "" Then
        On Error GoTo NoTable
        
        Set rngTable = Range("_QuickMoveTable")
        iQuickMove = Application.Match(strQuickMove, rngTable.Columns(1), 0)
        
        If iQuickMove > 0 Then
            Set GetQuickMoveData = rngTable.Rows(iQuickMove)
            Exit Function
        End If
    End If
    
NoTable:
    Set GetQuickMoveData = Nothing

End Function

Function GetChargeMoveData(strChargeMove As String) As Range
    Dim rngTable As Range
    Dim iChargeMove As Integer
    
    If strChargeMove <> "" Then
        On Error GoTo NoTable
        
        Set rngTable = Range("_ChargeMoveTable")
        iChargeMove = Application.Match(strChargeMove, rngTable.Columns(1), 0)
        
        If iChargeMove > 0 Then
            Set GetChargeMoveData = rngTable.Rows(iChargeMove)
            Exit Function
        End If
    End If
    
NoTable:
    Set GetChargeMoveData = Nothing

End Function

Function GetPokemonData(strNameData As String) As Range
    Dim rngTable As Range
    Dim iPokemon As Integer

    On Error GoTo NoTable
    
    Set rngTable = Range("_PokemonTable")
    iPokemon = Application.Match(strNameData, rngTable.Columns(1), 0)
    
    If iPokemon > 0 Then
        Set GetPokemonData = rngTable.Rows(iPokemon)
        Exit Function
    End If
    
NoTable:
    Set GetPokemonData = Nothing

End Function


Function ParsePokemonName(csvPokemon As String, Optional fDataName As Boolean = False) As String

    Dim iComma As Integer, iUnderscore As Integer
    Dim strPokemon As String, strParen As String
    
    'Pull the name of the Pokemon from PVPOKE description and parse out Pokemon name.
    'Correct case and put underscore qualifiers in parentheses for beauty.
    
    strPokemon = Application.WorksheetFunction.Proper(csvPokemon)
    
    iComma = InStr(strPokemon, ",")
    If iComma > 0 Then strPokemon = left(strPokemon, iComma - 1)
    
    iUnderscore = InStr(strPokemon, "_")
    If iUnderscore > 0 Then
        'Replace underscore with parenthesis.  Was necessary to Parse csv imported from pvpoke.com.
        strParen = Application.WorksheetFunction.Proper(Mid(strPokemon, iUnderscore + 1))
        strPokemon = left(strPokemon, iUnderscore - 1) & " (" & strParen & ")"
    End If
    
    If fDataName Then
        If InStr(strPokemon, "(shadow") > 0 Then
            strPokemon = Replace(strPokemon, "(shadow)", "")
        End If
    End If
    
    ParsePokemonName = Trim(strPokemon)

End Function

Function ParseMoveName(ByVal csvPokemon As String, ByVal iMove As Integer) As String
    
    ParseMoveName = Application.WorksheetFunction.Proper(ParseSubstring(csvPokemon, iMove + 1, ","))
    
    If InStr(ParseMoveName, "_") <> 0 Then
        ' convert csv imported from PvPoke.com
        ParseMoveName = Replace(ParseMoveName, "_", " ")
    End If
    
End Function

Function GetPkString(pk As Pokemon, iProp As Integer) As String
    If pk.rngData Is Nothing Then GetPkString = "" Else GetPkString = pk.rngData.Cells(1, iProp)
End Function

Function GetPkValue(pk As Pokemon, iProp As Integer) As Single
    If pk.rngData Is Nothing Then GetPkValue = 0 Else GetPkValue = pk.rngData.Cells(1, iProp)
End Function

Function GetDataString(rngData As Range, iProp As Integer) As String
    If rngData Is Nothing Then GetDataString = "" Else GetDataString = rngData.Cells(1, iProp)
End Function

Function GetDataValue(rngData As Range, iProp As Integer) As Single
    If rngData Is Nothing Then GetDataValue = 0 Else GetDataValue = rngData.Cells(1, iProp)
End Function

Function GetDataInteger(rngData As Range, iProp As Integer) As Integer
    If rngData Is Nothing Then GetDataInteger = 0 Else GetDataInteger = CInt(rngData.Cells(1, iProp))
End Function


Function GetBattleLeagueData(strBattleLeague As String) As Range
    Dim rngTable As Range
    Dim iBattleLeague As Integer

    On Error GoTo NoTable
    
    Set rngTable = Range("_BattleLeagueTable")
    iBattleLeague = Application.Match(strBattleLeague, rngTable.Columns(1), 0)
    
    If iBattleLeague > 0 Then
        Set GetBattleLeagueData = rngTable.Rows(iBattleLeague)
        Exit Function
    End If
    
NoTable:
    Set GetBattleLeagueData = Nothing

End Function

Function StatMultiplier(level As Single) As Single
    Dim rngTable As Range
    
    'Stat multiplier to convert a base stat to a stat that applies at a particular level of pokemon.
    'Explained here:  https://www.dragonflycave.com/pokemon-go/stats

    On Error GoTo NoCP
    
    Set rngTable = Range("_StatMultiplierTable")
    
    StatMultiplier = Application.VLookup(level, rngTable, 2, True)
    Exit Function
    
NoCP:
    StatMultiplier = 1

End Function

Function GetDmgQuickMove(qm As QuickMove) As Single
    
    GetDmgQuickMove = qm.rngData.Cells(1, 3).value

End Function

Function GetDptQuickMove(qm As QuickMove) As Single
    
    GetDptQuickMove = qm.rngData.Cells(1, 6).value

End Function

Function GetEptQuickMove(qm As QuickMove) As Single

    GetEptQuickMove = qm.rngData.Cells(1, 7).value

End Function

Function IsStatAlteringChargeMove(cm As ChargeMove) As Boolean

    IsStatAlteringChargeMove = cm.rngData.Cells(1, 6).value <> 0

End Function

Function GetEnergyChargeMove(cm As ChargeMove) As Single

    GetEnergyChargeMove = cm.rngData.Cells(1, 4).value
      
End Function

Function GetDmgChargeMove(cm As ChargeMove) As Single

    GetDmgChargeMove = cm.rngData.Cells(1, 3).value
      
End Function

Function SymbolForType(strType As String) As String

    Dim rngTable As Range
    
    If strType = "" Then
        SymbolForType = ""
    Else
        On Error GoTo NoTypeSymbol
        
        Set rngTable = Range("_TypeSymbolTable")
        
        ' Use the table for the table_array parameter
        
        SymbolForType = Application.VLookup(strType, rngTable, 2, False)
    End If
    
    Exit Function
    
NoTypeSymbol:
    SymbolForType = "?"

End Function

Function GetSpecialEffectSymbols(cm As ChargeMove) As String
    Dim rngTable As Range
    
    If cm.rngData Is Nothing Then
        GetSpecialEffectSymbols = "?"
    Else
        GetSpecialEffectSymbols = cm.rngData.Cells(1, 11)
    End If
    
NoSpecialEffect:
End Function

Function SpecialEffectSymbols(strMove As String) As String
    Dim rngTable As Range
    
    On Error GoTo NoSpecialEffect
    
    SpecialEffectSymbols = ""
    
    Set rngTable = Range("_ChargeMoveTable")
    
    ' Use the table for the table_array parameter
    
    SpecialEffectSymbols = Application.VLookup(strMove, rngTable, 11, False)
    
NoSpecialEffect:
End Function


Function SpecialEffectWarning(strMove As String) As String
    Dim rngTable As Range
    
    On Error GoTo NoSpecialEffect
    
    SpecialEffectWarning = ""
    
    Set rngTable = Range("_ChargeMoveTable")
    
    ' Use the table for the table_array parameter
    
    SpecialEffectWarning = Application.VLookup(strMove, rngTable, 12, False)
    
NoSpecialEffect:
End Function

Function ChargeMoveAbbreviation(strMove As String) As String
    Dim rngTable As Range
    
    On Error GoTo NoAbbreviation
    
    Set rngTable = Range("_ChargeMoveTable")
    ChargeMoveAbbreviation = Application.VLookup(strMove, rngTable, 13, False)
    
    If ChargeMoveAbbreviation = "" Then ChargeMoveAbbreviation = strMove
    Exit Function
    
NoAbbreviation:
    ChargeMoveAbbreviation = strMove
End Function

Function TypeOfQuickMove(strMove As String) As String
' Return "" if no attack, a string describing the attack type, or "Unknown" if an attack not in the table.

    Dim rngTable As Range
    
    If strMove = "" Then
        TypeOfQuickMove = ""
        Exit Function
    End If
    
    On Error GoTo NoMoveType
    
    Set rngTable = Range("_QuickMoveTable")
    
    ' Use the table for the table_array parameter
    
    TypeOfQuickMove = Application.VLookup(strMove, rngTable, 2, False)
    Exit Function
    
NoMoveType:
    TypeOfQuickMove = "Unknown"

End Function

Function TypeOfChargeMove(strMove As String) As String
' Return "" if no attack, a string describing the attack type, or "Unknown" if an attack not in the table.

    Dim rngTable As Range
    
    If strMove = "" Then
        TypeOfChargeMove = ""
        Exit Function
    End If
    
    On Error GoTo NoMoveType
    
    Set rngTable = Range("_ChargeMoveTable")
    
    ' Use the table for the table_array parameter
    
    TypeOfChargeMove = Application.VLookup(strMove, rngTable, 2, False)
    Exit Function
    
NoMoveType:
    TypeOfChargeMove = "Unknown"

End Function

Function SymbolsForPokemon(csv As String) As String
    Dim strPokemon As String
    Dim strType1 As String, strType2 As String
    Dim strSymbol1 As String, strSymbol2 As String
    Dim rngData As Range
    
    'On Windows, consider formatting output string with font Segoe UI Symbol.  Some fonts do not have all the symbols
    
    strPokemon = ParsePokemonName(csv, True)
    
    If strPokemon = "" Then
        SymbolsForPokemon = ""
        Exit Function
    End If
    
    Set rngData = GetPokemonData(strPokemon)
    
    If rngData Is Nothing Then
        SymbolsForPokemon = SymbolForType(strPokemon)
    Else
        strType1 = rngData.Cells(1, pkData_Type1)
        strType2 = rngData.Cells(1, pkData_Type2)
        
        strSymbol1 = SymbolForType(strType1)
        strSymbol2 = SymbolForType(strType2)
        
        SymbolsForPokemon = strSymbol1 & strSymbol2
    End If

End Function
    


Function SymbolsForMoveset(csv As String) As String
    Dim pk As Pokemon
    Dim strSymbols As String, strNextMove As String, iNextMove As Integer
    Dim strQuickMove As String
    
    'On Windows, consider formatting output string with font Segoe UI Symbol.  Some fonts do not have all the symbols.
    
    strQuickMove = ParseMoveName(csv, 1)
    strNextMove = ParseMoveName(csv, 2)
    
    ' Require at least the quickmove and one charge move.
    strSymbols = SymbolForType(TypeOfQuickMove(strQuickMove)) & SymbolForType(TypeOfChargeMove(strNextMove))
    
    iNextMove = 3
    strNextMove = ParseMoveName(csv, iNextMove)
    
    While strNextMove <> ""
        strSymbols = strSymbols & SymbolForType(TypeOfChargeMove(strNextMove))
        iNextMove = iNextMove + 1
        strNextMove = ParseMoveName(csv, iNextMove)
    Wend
    
    SymbolsForMoveset = strSymbols
End Function

Function SpecialEffectsForMoveset(csv As String) As String
    Dim pk As Pokemon
    Dim strSymbols As String, strNextMove As String, strNextSymbols As String, iNextMove As Integer
    Dim strError As String
    
    'On Windows, special effects symbols look good in Verdana, and better at larger sizes of verdana.
    
    Call InitPokemon(pk, csv)
    strError = StrValidatePk(pk)
    
    If strError <> "" Then
        SpecialEffectsForMoveset = strError
        Exit Function
    End If

' Now the special effects symbol.
    
    ' Require at least the quickmove and one charge move.
    strSymbols = SpecialEffectSymbols(ParseMoveName(csv, 2))
    
    iNextMove = 3
    strNextMove = ParseMoveName(csv, iNextMove)
    
    While strNextMove <> ""
        strNextSymbols = SpecialEffectSymbols(strNextMove)
        
        If strNextSymbols <> "" Then
            If strSymbols = "" Then
                strSymbols = strNextSymbols
            Else
                strSymbols = strSymbols & "," & strNextSymbols
            End If
        End If
        
        iNextMove = iNextMove + 1
        strNextMove = ParseMoveName(csv, iNextMove)
    Wend
    
    If strSymbols = "" Then strSymbols = "OK"
    SpecialEffectsForMoveset = strSymbols

End Function

Sub InitQuickMove(qm As QuickMove, strQuickMove As String)
    Dim qmInit As QuickMove
    
    qm = qmInit ' clear data
    qm.strMove = strQuickMove
    
    If qm.strMove = "" Then
        qm.strType = ""
    Else
        Set qm.rngData = GetQuickMoveData(qm.strMove)
    
        If qm.rngData Is Nothing Then
            qm.strType = "Unknown"
        Else
            qm.strType = qm.rngData.Cells(1, 2).value
            qm.cTurnsToQuick = GetDataValue(qm.rngData, 5)
        End If
    End If

End Sub

Sub InitChargeMove(cm As ChargeMove, qm As QuickMove, strChargeMove As String)
    Dim cmInit As ChargeMove
    
    cm = cmInit ' clear data
    cm.strMove = strChargeMove
    
    If cm.strMove = "" Then
        cm.strType = ""
    Else
        Set cm.rngData = GetChargeMoveData(cm.strMove)
    
        If cm.rngData Is Nothing Then
            cm.strType = "Unknown"
        Else
            cm.strType = cm.rngData.Cells(1, 2).value
            cm.cTurnsToCharge = CTurnsToChargeMove(cm, qm)
        End If
    End If

End Sub

Sub InitPokemon(ByRef pk As Pokemon, csv As String)
    Dim pkInit As Pokemon
    Dim iNextMove As Integer, strNextMove As String, strQuickMove As String, strChargeMove As String
    Dim strCategory As String
    Dim strType As String

    
    pk = pkInit ' empty
    
    pk.strName = ParsePokemonName(csv)
    pk.csv = pk.strName  ' we will rebuild this to beautify and standardize
    
    strQuickMove = ParseMoveName(csv, 1)
    If strQuickMove <> "" Then pk.csv = pk.csv & ", " & strQuickMove
    
    strChargeMove = ParseMoveName(csv, 2)
    If strChargeMove <> "" Then pk.csv = pk.csv & ", " & strChargeMove
    
    iNextMove = 3
    strNextMove = ParseMoveName(csv, iNextMove)
    
    While strNextMove <> ""
        strType = TypeOfChargeMove(strNextMove)
        If strType = "Unknown" Then
            pk.fInvalid = True
            pk.fInvalidChargeMove = True
        End If
        pk.fMultipleChargeMoves = True
        
        pk.csv = pk.csv & ", " & strNextMove
        iNextMove = iNextMove + 1
        strNextMove = ParseMoveName(csv, iNextMove)
    Wend
    
    ' pk.csv is now a normalized version of csv, with consistent spacing and capitalization.

    pk.strNameData = pk.strName

    If InStr(pk.strName, "(shadow") > 0 Then
        pk.strNameData = Trim(Replace(pk.strName, "(shadow)", ""))
        pk.fShadow = True
    End If
    
    Set pk.rngData = GetPokemonData(pk.strNameData)
    
    If pk.rngData Is Nothing Then
        pk.fInvalid = True
        pk.fInvalidPokemon = True
    Else
        pk.strType1 = pk.rngData.Cells(1, pkData_Type1)
        pk.strType2 = pk.rngData.Cells(1, pkData_Type2)
        
        Call InitQuickMove(pk.qm, strQuickMove)
        
        If pk.qm.rngData Is Nothing Then
            pk.fInvalid = True
            pk.fInvalidQuickMove = True
        Else
            Call InitChargeMove(pk.cm, pk.qm, strChargeMove)
    
            If pk.cm.rngData Is Nothing Then
                pk.fInvalid = True
                pk.fInvalidChargeMove = True
            End If
        End If
    End If

    strCategory = GetPkString(pk, pkData_Category)
    If strCategory = "M" Or strCategory = "L" Then pk.fLegendaryOrMythical = True
    
End Sub

Function StrValidatePk(pk As Pokemon) As String
    StrValidatePk = ""
    If pk.fInvalid Then
        If pk.fInvalidPokemon Then
            StrValidatePk = "Not A Pokemon"
        Else
            If pk.qm.strMove = "" Then
                StrValidatePk = "Missing Quick Move"
            ElseIf pk.fInvalidQuickMove Then
                StrValidatePk = "Bad Quick Move"
            ElseIf pk.cm.strMove = "" Then
                StrValidatePk = "Missing Charge Move"
            ElseIf pk.fInvalidChargeMove Then
                StrValidatePk = "Bad Charge Move"
            End If
        End If
    End If
End Function

' Qualify a Pokemon for battle.

Sub QualifyPokemon(pk As Pokemon, csvIV As String, ByVal rngDataBattleLeague As Range, fTypeMuseOK As Boolean)
    Dim cpMaxBattleLeague As Integer

    If pk.fInvalid Then Exit Sub
    
    cpMaxBattleLeague = GetDataInteger(rngDataBattleLeague, blData_MaxCp)
    
    ' let's use 13, 13, 13 as our default ivs
    pk.ivs.Attack = 13: pk.ivs.Defense = 13: pk.ivs.Stamina = 13: pk.ivs.levelMax = 40

    If csvIV <> "" Then
        Dim str1 As String, str2 As String, str3 As String, str4 As String
        
        Call Parse4Substrings(csvIV, ",", str1, str2, str3, str4)
        If IsNumeric(str1) Then pk.ivs.Attack = MinMaxI(CInt(str1), 1, 15)
        If IsNumeric(str2) Then pk.ivs.Defense = MinMaxI(CInt(str2), 1, 15)
        If IsNumeric(str3) Then pk.ivs.Stamina = MinMaxI(CInt(str3), 1, 15)
        If IsNumeric(str4) Then pk.ivs.levelMax = MinMax(CDec(str4), 1, 41)
        pk.ivs.csvIV = csvIV
    End If
    
    
    pk.fQualified = True
    
    ' Qualified unless the League has restrictions.
    
    Select Case GetDataString(rngDataBattleLeague, blData_Restriction)
    
    Case "Premier"
        If pk.fLegendaryOrMythical Then pk.fQualified = False
        
    Case "Flying"
        If pk.strType1 <> "Flying" And pk.strType2 <> "Flying" Then pk.fQualified = False
    
    End Select
    
    Call CalcPokemonStats(pk, cpMaxBattleLeague)

End Sub



