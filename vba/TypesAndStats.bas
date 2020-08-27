Attribute VB_Name = "TypesAndStats"
Option Explicit

Function TypeEffectivenessMultiplier(strMoveType As String, pk As Pokemon, pkDefender As Pokemon) As Single
    Dim effectiveness As Integer
    Dim mult As Single
    
    ' These scores are the attack strength multipliers used by Pokemon Go.
    ' The highest would be 3.072, if the move is doubly effective against the defender and has 20% STAB.

    effectiveness = TypeEffectiveness(strMoveType, pkDefender.strType1) + TypeEffectiveness(strMoveType, pkDefender.strType2)
    
    mult = 1.6 ^ effectiveness
    ' Extreme case is Tropius, where value ranges from 0.244 to 2.56
    
    If strMoveType = pk.strType1 Or strMoveType = pk.strType2 Then
        mult = mult * 1.2   ' add stab
    End If
    
    ' Including stab, maximum possible value is 3.072
    
    TypeEffectivenessMultiplier = mult

End Function

Function TypeEffectiveness(strTypeAttack As String, strTypeDefend As String) As Integer

    'TypeEffectiveness returns +1, 0, -1, or -2 in all cases.
    '+1 signifies super effective. 0 signifies neutral.  -1 not very effective.  -2 doubly ineffective.
    'Add the TypeEffectiveness against the defender's primary and secondary types.

    TypeEffectiveness = 0
    
    If strTypeAttack = "" Or strTypeAttack = "Unknown" Or strTypeDefend = "" Or strTypeDefend = "Unknown" Then Exit Function
    
    Select Case strTypeAttack
    
    Case "Bug"
        Select Case strTypeDefend
        Case "Dark", "Grass", "Psychic"
            TypeEffectiveness = 1
        Case "Fairy", "Fighting", "Fire", "Flying", "Ghost", "Poison", "Steel"
            TypeEffectiveness = -1
        End Select
    
    Case "Dark"
        Select Case strTypeDefend
        Case "Ghost", "Psychic"
            TypeEffectiveness = 1
        Case "Dark", "Fairy", "Fighting"
            TypeEffectiveness = -1
        End Select
    
    Case "Dragon"
        Select Case strTypeDefend
        Case "Dragon"
            TypeEffectiveness = 1
        Case "Steel"
            TypeEffectiveness = -1
        Case "Fairy"
            TypeEffectiveness = -2
        End Select
    
    Case "Electric"
        Select Case strTypeDefend
        Case "Flying", "Water"
            TypeEffectiveness = 1
        Case "Dragon", "Electric", "Grass"
            TypeEffectiveness = -1
        Case "Ground"
            TypeEffectiveness = -2
        End Select
    Case "Fairy"
        Select Case strTypeDefend
        Case "Dark", "Dragon", "Fighting"
            TypeEffectiveness = 1
        Case "Fire", "Poison", "Steel"
            TypeEffectiveness = -1
        End Select
    Case "Fighting"
        Select Case strTypeDefend
        Case "Dark", "Ice", "Normal", "Rock", "Steel"
            TypeEffectiveness = 1
        Case "Bug", "Fairy", "Flying", "Poison", "Psychic"
            TypeEffectiveness = -1
        Case "Ghost"
            TypeEffectiveness = -2
        End Select
    Case "Fire"
        Select Case strTypeDefend
        Case "Bug", "Grass", "Ice", "Steel"
            TypeEffectiveness = 1
        Case "Dragon", "Fire", "Rock", "Water"
            TypeEffectiveness = -1
        End Select
    Case "Flying"
        Select Case strTypeDefend
        Case "Bug", "Fighting", "Grass"
            TypeEffectiveness = 1
        Case "Electric", "Rock", "Steel"
            TypeEffectiveness = -1
        End Select
    Case "Ghost"
        Select Case strTypeDefend
        Case "Ghost", "Psychic"
            TypeEffectiveness = 1
        Case "Dark"
            TypeEffectiveness = -1
        Case "Normal"
            TypeEffectiveness = -2
        End Select
    Case "Grass"
        Select Case strTypeDefend
        Case "Ground", "Rock", "Water"
            TypeEffectiveness = 1
        Case "Bug", "Dragon", "Fire", "Flying", "Grass", "Poison", "Steel"
            TypeEffectiveness = -1
        End Select
    Case "Ground"
        Select Case strTypeDefend
        Case "Electric", "Fire", "Poison", "Rock", "Steel"
            TypeEffectiveness = 1
        Case "Bug", "Grass"
            TypeEffectiveness = -1
        Case "Flying"
            TypeEffectiveness = -2
        End Select
    Case "Ice"
        Select Case strTypeDefend
        Case "Dragon", "Flying", "Grass", "Ground"
            TypeEffectiveness = 1
        Case "Fire", "Ice", "Steel", "Water"
            TypeEffectiveness = -1
        End Select
    Case "Normal"
        Select Case strTypeDefend
        Case "Rock", "Steel"
            TypeEffectiveness = -1
        Case "Ghost"
            TypeEffectiveness = -2
        End Select
    Case "Poison"
        Select Case strTypeDefend
        Case "Fairy", "Grass"
            TypeEffectiveness = 1
        Case "Ghost", "Ground", "Poison", "Rock"
            TypeEffectiveness = -1
        Case "Steel"
            TypeEffectiveness = -2
        End Select
    Case "Psychic"
        Select Case strTypeDefend
        Case "Fighting", "Poison"
            TypeEffectiveness = 1
        Case "Psychic", "Steel"
            TypeEffectiveness = -1
        Case "Dark"
            TypeEffectiveness = -2
        End Select
    Case "Rock"
        Select Case strTypeDefend
        Case "Bug", "Fire", "Flying", "Ice"
            TypeEffectiveness = 1
        Case "Fighting", "Ground", "Steel"
            TypeEffectiveness = -1
        End Select
    Case "Steel"
        Select Case strTypeDefend
        Case "Fairy", "Ice", "Rock"
            TypeEffectiveness = 1
        Case "Electric", "Fire", "Steel", "Water"
            TypeEffectiveness = -1
        End Select
    Case "Water"
        Select Case strTypeDefend
        Case "Fire", "Ground", "Rock"
            TypeEffectiveness = 1
        Case "Dragon", "Grass", "Water"
            TypeEffectiveness = -1
        End Select
    End Select

End Function

Sub CalcQuickMoveStats(pk As Pokemon, pkDefender As Pokemon)
    With pk.qm
        .factorMult = TypeEffectivenessMultiplier(.strType, pk, pkDefender)
        
        ' https://gamepress.gg/pokemongo/damage-mechanics
        ' https://www.reddit.com/r/TheSilphRoad/comments/i2mvde/calculating_damage_done_by_an_attack_in_pogo/
        
        .dmgQuick = GetDmgQuickMove(pk.qm) * .factorMult * pk.bstat.att * factor_PVP
        .hpPerQuick = RoundDown(.dmgQuick / pkDefender.bstat.def) + 1 ' Initial, pre-buff hp taken per quick move.
        
        .dptQuick = .dmgQuick / .cTurnsToQuick
        .hpptQuick = CSng(.hpPerQuick) / CSng(.cTurnsToQuick) ' ensure floating point math.
        .dptQuickInit = .dptQuick
        .cTurnsToVictory = CTurnsToVictoryQm(pk.qm, pkDefender.bstat.hp, pkDefender.bstat.def)
    End With
End Sub

Sub CalcChargeMoveStats(cm As ChargeMove, pk As Pokemon, pkDefender As Pokemon, cTurnsMaxBattle As Single)

    With cm
        .factorMult = TypeEffectivenessMultiplier(.strType, pk, pkDefender)
        .factorTime = TimeFactor(cm, cTurnsMaxBattle)
        
        ' https://gamepress.gg/pokemongo/damage-mechanics
        ' https://www.reddit.com/r/TheSilphRoad/comments/i2mvde/calculating_damage_done_by_an_attack_in_pogo/
        
        .dmgCharge = GetDmgChargeMove(cm) * .factorMult * pk.bstat.att * factor_PVP
        .hpPerCharge = RoundDown(.dmgCharge / pkDefender.bstat.def) + 1 ' Initial , pre-buff, pre-time factor hp taken per charge move.
        
        .dptChargeInit = .dmgCharge / .cTurnsToCharge
        .dptCharge = .dptChargeInit * .factorTime
        .hpptCharge = HpPerTurnCm(cm, pkDefender.bstat.def)
        
        ' number of turns in battle , estimate if half of charge move damage is blocked or surplus.  For planning, not scoring.
        .cTurnsToVictory = RoundUpTurnsQm(pkDefender.bstat.hp / (pk.qm.hpptQuick + .hpptCharge / 2), pk.qm)
        
    End With

End Sub


Sub CalcPokemonStats(pk As Pokemon, cpMax As Integer)
    Dim valAttack As Integer, valDefense As Integer, valStamina As Integer, statMult As Single
    
    ' level and stats of each pokemon
    ' Initial stats are preserved in cases where stat changing attacks may change the stats during battle.
    
    valAttack = CInt(pk.rngData.Cells(1, pkData_Attack)) + pk.ivs.Attack
    valDefense = CInt(pk.rngData.Cells(1, pkData_Defense)) + pk.ivs.Defense
    valStamina = CInt(pk.rngData.Cells(1, pkData_Stamina)) + pk.ivs.Stamina
    
    With pk.bstat
    
    For .level = pk.ivs.levelMax To 0 Step -0.5
        .cp = CpAtLevelFromStats(.level, valAttack, valDefense, valStamina)
        If .cp <= cpMax Then Exit For
    Next .level
    
    statMult = StatMultiplier(.level)
    
    .attCMP = statMult * valAttack      ' attack stat for Charge Move priority.
    .attInit = .attCMP                  ' start battling with .attInit
    
    .defInit = statMult * valDefense    ' start battling with .defInit
    
    If pk.fShadow Then                  ' alter .attInit and .defInit symetrically for Shadow Pokemon
        .attInit = (6 / 5) * .attInit
        .defInit = (5 / 6) * .defInit
    End If
    
    .att = .attInit   ' .att starts with .attInit, but can be changed with stat changing moves.
    .def = .defInit   ' .def starts with .defInit, but can be changed with stat changing moves.
    
    .hp = RoundDown(statMult * valStamina)
    
    End With

End Sub


Function LevelAtCpForPokemon(pk As Pokemon, cpMax As Integer, levelMax As Single)
    Dim valAttack As Integer, valDefense As Integer, valStamina As Integer
    
    valAttack = pk.rngData.Cells(1, pkData_Attack) + pk.ivs.Attack
    valDefense = pk.rngData.Cells(1, pkData_Defense) + pk.ivs.Defense
    valStamina = pk.rngData.Cells(1, pkData_Stamina) + pk.ivs.Stamina
    LevelAtCpForPokemon = LevelAtCpFromStats(cpMax, levelMax, valAttack, valDefense, valStamina)

End Function


Function LevelAtCpFromStats(cpMax As Integer, levelMax As Single, valAttack As Integer, valDefense As Integer, valStamina As Integer) As Single
    Dim level As Single
    
    For level = levelMax To 0 Step -0.5
        If CpAtLevelFromStats(level, valAttack, valDefense, valStamina) <= cpMax Then Exit For
    Next level
    
    LevelAtCpFromStats = level
End Function

Function CpAtLevelFromStats(level As Single, valAttack As Integer, valDefense As Integer, valStamina As Integer) As Single
' see https://gamepress.gg/pokemongo/pokemon-stats-advanced#:~:text=Calculating%20CP,*%20CP_Multiplier%5E2)%20%2F%2010
    Dim mult As Single
    
    'valAttack, valDefense, valStamina are the pokemon's base stat plus the related IV of the pokemon.
    'For a perfect pokemon, add 15 to each base stat.
    
    mult = StatMultiplier(level)
    CpAtLevelFromStats = RoundDown((valAttack * Sqr(valDefense) * Sqr(valStamina) * mult * mult) / 10)
    
End Function


Function CTurnsToChargeMove(cm As ChargeMove, qm As QuickMove) As Single

    Dim cTurnsToCharge As Single
    
    If cm.rngData Is Nothing Or qm.rngData Is Nothing Then
        CTurnsToChargeMove = 0 ' test case: charge move is valid but quick move is invalid during InitChargeMove
    Else
        cTurnsToCharge = GetEnergyChargeMove(cm) / GetEptQuickMove(qm)
        
        ' The answer must be an integer and a multiple of the duration of the quick move
        CTurnsToChargeMove = RoundUpTurnsQm(cTurnsToCharge, qm)
    End If

End Function

Function CTurnsToVictoryQm(qm As QuickMove, hpDefender As Integer, defDefender As Single) As Single
    Dim cTurns As Single, dhpPerQuickMove As Single

    dhpPerQuickMove = Max(RoundDown(qm.dptQuick * qm.cTurnsToQuick / defDefender), 1)
    
    cTurns = (hpDefender / dhpPerQuickMove) * qm.cTurnsToQuick
    CTurnsToVictoryQm = Application.WorksheetFunction.RoundUp(cTurns / qm.cTurnsToQuick, 0) * qm.cTurnsToQuick
    
End Function

Function HpPerTurnQm(qm As QuickMove, defDefender As Single) As Single

    Dim cTurns As Single, dhpPerQuickMove As Single
    
    ' https://www.reddit.com/r/TheSilphRoad/comments/i2mvde/calculating_damage_done_by_an_attack_in_pogo/

    dhpPerQuickMove = RoundDown(qm.dptQuick * qm.cTurnsToQuick / defDefender) + 1
    
    HpPerTurnQm = dhpPerQuickMove / qm.cTurnsToQuick

End Function

Function HpPerTurnCm(cm As ChargeMove, defDefender As Single) As Single
    Dim dhpPerChargeMove
    
    ' https://www.reddit.com/r/TheSilphRoad/comments/i2mvde/calculating_damage_done_by_an_attack_in_pogo/

    dhpPerChargeMove = RoundDown(cm.dptCharge * cm.cTurnsToCharge / defDefender) + 1
    
    HpPerTurnCm = dhpPerChargeMove / cm.cTurnsToCharge

End Function

Function RoundUpTurnsQm(cTurns As Single, qm As QuickMove) As Single
    
    RoundUpTurnsQm = Application.WorksheetFunction.RoundUp(cTurns / qm.cTurnsToQuick, 0) * qm.cTurnsToQuick
    
End Function

Function TimeFactor(cm As ChargeMove, cTurnsMaxBattle As Single) As Single
    Dim cTurnsCharge As Single
    
    TimeFactor = 1  ' Default answer indicating no time factor impact, or no battle details available.
    
    If cTurnsMaxBattle > 0 Then
        ' the longer a charge attack takes to fire, the less valuable it is.
        ' consider valuing an attack at half value if it takes a whole expected battle to fire.
        ' notice that quicker attacks are still discounted, but less so exponentially.
        
        'if cTurnsMaxBattle is infinite, TimeFactor is just 1.
        'The default cTurnsMaxBattle is 30 turns, but BattleScore provides a custom estimate.
        
        TimeFactor = (1 / (2 ^ (cm.cTurnsToCharge * 1.6 / cTurnsMaxBattle)))
        
    End If
    
End Function

' This code not used, as shield factor is not implemented.

Function ShieldFactor(cm As ChargeMove, cTurnsInBattle As Single, cShields As Single) As Single
    Dim cTurnsCharge As Single
    
    ShieldFactor = 1
    
    If cTurnsInBattle > 0 Then

        cTurnsCharge = cm.cTurnsToCharge
        
        If cShields > 0 Then
            Dim cChargeMovesInBattle As Single, factorShield As Single
    
            cChargeMovesInBattle = cTurnsInBattle / cTurnsCharge
            
            If cShields >= cChargeMovesInBattle Then
                factorShield = 0
            Else
                factorShield = (cChargeMovesInBattle - cShields) / cChargeMovesInBattle
            End If
            
            ShieldFactor = factorShield
        End If
    End If
    
End Function
