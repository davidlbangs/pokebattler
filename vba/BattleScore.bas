Attribute VB_Name = "BattleScore"
' Pokemon Go Battle Planner
' (c) 2020 David Bangs.  All rights reserved
'
' A tool to help PVP particants plan an ideal team and to guide them through battles by providing a heads up view of
' best moves and matchups.

Option Explicit
Option Compare Text

Dim breakPoint As Integer

Function BattleScoreX(ByVal csvAttacker As String, ByVal csvDefender As String, ByVal strBattleLeague As String) As Integer

    ' Simple function that can be called from a spreadsheet cell, using cells present on our spreadsheet pages as the only inputs.
    ' This is only for documentation and testing purposes.  The X in its name discourages its use.

    Dim pk1 As Pokemon, pk2 As Pokemon, rngDataBattleLeague As Range
    
    Call InitPokemon(pk1, csvAttacker)
    Call InitPokemon(pk2, csvDefender)
    
    Set rngDataBattleLeague = GetBattleLeagueData(strBattleLeague)
    
    Call QualifyPokemon(pk1, "13,13,13,40", rngDataBattleLeague, True)
    Call QualifyPokemon(pk2, "13,13,13,40", rngDataBattleLeague, True)
    
    BattleScoreX = BattleScoreCore(pk1, pk2)

End Function

Function BattleScoreCore(ByRef pk1 As Pokemon, ByRef pk2 As Pokemon) As Integer

' BattleScore gives a Pokémon an effectiveness score against another Pokémon,
' based on the types of both Pokémon and the types of their quick and charge moves.
'
' The score ranges from 0 to 1000.  500 is a tie.
' The scores are symmetrical so Score (A vs B )+ Score (B vs A) = 1000.

' The method of reporting the score is the same as used by the battle simulation web site pvpoke.com.
' But this is NOT an attack by attack battle simulation!
'
' Winning a duel is NOT just winning three battles between pairs of pokemon.  Pokemon are
' changed out mid battle, so any simulation that assumes a pair of pokemon starts fresh together and use a pre-determined
' number of shields is not  useful in decided whether to switch out pokemon mid-battle.

' Rather than modelling damage per quick and charge attack, we model it per turn, or per 0.5 second interval.
' This is not factually accurate in that it spreads the impact of the charge attack evenly accross all turns.  However, actual timing
' is unknown in a situation where pokemon are being changed mid-battle. Shield use is also unpredictable in real battle.

' Note that slow moves are penalized heavily, first by a function called TimeFactor(), which returns a lower factor for slower moves.
' TrimWastedDamage tries to keep us from getting overly excited about a move that can kill the oponent 5 times over.  Once is just as good.
' RewardFirstMoveAdvantage is a nod to the fact that the pokemon that can launch its attack first might just win more.
'
' Again, this has been tuned for the purposes of helping a player decide in real time which pokemon to bring in against each oponent, and which
' charge moves to launch.
'
' Battles are either based on TypeEffectiveness or BattleLeague. BattleLeague is used way more, but early support for TypeEffectiveness calculations
' which ignore pokemon and move stats was retained because it may be useful to battle a pokemon against a "Type Muse", so that types may be present
' in the battle table as a stand-in for when the oposing pokemon is not in the table but is of known type.

Dim breakPoint As Integer

    BattleScoreCore = -1 'error
    
    'We need two qualified pokemon!
    
    If Not (pk1.fQualified And pk2.fQualified) Then Exit Function
        
    If pk1.fTypeEffectivenessBattle Or pk2.fTypeEffectivenessBattle Then
        'Type Effectivness Battle
        
        'Calculations based ONLY on Type Effectiveness Only.
        
        Dim damageByPk1 As Single, damageByPk2 As Single
        
        pk1.qm.factorMult = TypeEffectivenessMultiplier(pk1.qm.strType, pk1, pk2)
        pk2.qm.factorMult = TypeEffectivenessMultiplier(pk2.qm.strType, pk2, pk1)
        
        pk1.cm.factorMult = TypeEffectivenessMultiplier(pk1.cm.strType, pk1, pk2)
        pk2.cm.factorMult = TypeEffectivenessMultiplier(pk2.cm.strType, pk2, pk1)

        ' These functions will switch out the charge move IF the pokemon has multiple charge moves and there is a more effective move
        ' in the list.
        
        Call DetermineBestChargeMoveByType(pk1, pk2)
        Call DetermineBestChargeMoveByType(pk2, pk1)
        
        pk1.cm.strMove = "[" & pk1.cm.strType & "]" 'ByRef
        pk2.cm.strMove = "[" & pk2.cm.strType & "]" 'ByRef
    
        ' considering only type effectiveness and stab, assuming quick and charge moves have equal damage potential.
        damageByPk1 = pk1.qm.factorMult + pk1.cm.factorMult
        damageByPk2 = pk2.qm.factorMult + pk2.cm.factorMult
        
        BattleScoreCore = 1000 * (damageByPk1 / (damageByPk1 + damageByPk2))
    Else
        Dim cTurnsMaxBattle As Single
        
        'League Battle
        
        Call CalcQuickMoveStats(pk1, pk2)
        Call CalcQuickMoveStats(pk2, pk1)
        
        ' number of turns in battle if only quick moves were used
        cTurnsMaxBattle = Min(pk1.qm.cTurnsToVictory, pk2.qm.cTurnsToVictory)
            
        ' consider best charge move using detailed pokemon and battle knowledge.
        
        Call CalcChargeMoveStats(pk1.cm, pk1, pk2, cTurnsMaxBattle)
        Call CalcChargeMoveStats(pk2.cm, pk2, pk1, cTurnsMaxBattle)

        Call DetermineBestChargeMoves(pk1, pk2, cTurnsMaxBattle)
        Call DetermineBestChargeMoves(pk2, pk1, cTurnsMaxBattle)
        
        If pk1.cmBestBuff.strMove <> pk1.cm.strMove Or IsStatAlteringChargeMove(pk1.cm) Or _
            pk2.cmBestBuff.strMove <> pk2.cm.strMove Or IsStatAlteringChargeMove(pk2.cm) Then
            
            Dim cTurnsInBattle As Single
            
            ' One or more stat altering moves are in play!  Let them do their work before calculating the final score.
            ' Note that AdjustForBuff is actually very clever and may COMBINE the buff and charge move into a one/two punch.
            
            cTurnsInBattle = Min(pk1.cm.cTurnsToVictory, pk2.cm.cTurnsToVictory)
        
            Call AdjustForBuff(pk1.cm, pk1.cmBestBuff, cTurnsMaxBattle, cTurnsInBattle, _
                pk1.qm.dptQuick, pk1.cm.dptCharge, pk1.bstat.def, _
                pk2.qm.dptQuick, pk2.cm.dptCharge, pk2.bstat.def)

            Call AdjustForBuff(pk2.cm, pk2.cmBestBuff, cTurnsMaxBattle, cTurnsInBattle, _
                pk2.qm.dptQuick, pk2.cm.dptCharge, pk2.bstat.def, _
                pk1.qm.dptQuick, pk1.cm.dptCharge, pk1.bstat.def)
                
            ' Requantize - So that each descrete move takes an integer number of hp points post buff.
            pk1.qm.hpptQuick = HpPerTurnQm(pk1.qm, pk2.bstat.def)
            pk1.cm.hpptCharge = HpPerTurnCm(pk1.cm, pk2.bstat.def)
            pk2.qm.hpptQuick = HpPerTurnQm(pk2.qm, pk1.bstat.def)
            pk2.cm.hpptCharge = HpPerTurnCm(pk2.cm, pk1.bstat.def)
        End If
        
        ' Limit cm.hpptCharge to reduce extreme wasted energy caused by overkill from distorting recommendations.
        Call TrimWastedDamage(pk1, pk2)
        Call TrimWastedDamage(pk2, pk1)
        
        Call RewardFirstMoveAdvantage(pk1, pk2)
        
        ' Final calculations and score
        
        pk1.cm.cTurnsToVictory = RoundUpTurnsQm(pk2.bstat.hp / (pk1.qm.hpptQuick + pk1.cm.hpptCharge), pk1.qm)
        pk2.cm.cTurnsToVictory = RoundUpTurnsQm(pk1.bstat.hp / (pk2.qm.hpptQuick + pk2.cm.hpptCharge), pk2.qm)
        
        'Score is a per-thousand ratio ratio of damage by attacker to combined damage. A score of 500 denotes a tie.
        
        BattleScoreCore = 1000 * (pk2.cm.cTurnsToVictory / (pk1.cm.cTurnsToVictory + pk2.cm.cTurnsToVictory))
        
    End If
    
End Function

Sub DetermineBestChargeMoves(pk As Pokemon, pkDefender As Pokemon, ByVal cTurnsMaxBattle As Single)

    pk.cmBest = pk.cm
    pk.cmBestBuff = pk.cm
    pk.cmStrongest = pk.cm
    pk.cmQuickest = pk.cm

    If pk.fMultipleChargeMoves Then
        Dim cmNext As ChargeMove, strNext As String, iNext As Integer
                
        pk.cmBest.factorBuff = BuffFactor(pk.cmBest, pkDefender.cm, cTurnsMaxBattle)
        pk.cmBestBuff.factorBuff = pk.cmBest.factorBuff
        
        iNext = 3
        strNext = ParseMoveName(pk.csv, iNext)
            
        While strNext <> ""
            Call InitChargeMove(cmNext, pk.qm, strNext)
            Call CalcChargeMoveStats(cmNext, pk, pkDefender, cTurnsMaxBattle)
            cmNext.factorBuff = BuffFactor(cmNext, pkDefender.cm, cTurnsMaxBattle)
    
            If (cmNext.dptCharge + pk.qm.dptQuick) * cmNext.factorBuff > _
                (pk.cmBest.dptCharge + pk.qm.dptQuick) * pk.cmBest.factorBuff Then
                
                pk.cmBest = cmNext
                
                ' best move yet.  If it is tied to be best buff or strongest or quickest attack, break the tie in its favor.
                
                If cmNext.dmgCharge >= pk.cmStrongest.dmgCharge Then
                    pk.cmStrongest = cmNext
                End If
                
                If cmNext.factorBuff >= pk.cmBestBuff.factorBuff Then 'buff's being equal, the best move wins the buff contest.
                    pk.cmBestBuff = cmNext
                End If
                
                If cmNext.cTurnsToCharge <= pk.cmBestBuff.cTurnsToCharge Then
                    pk.cmQuickest = cmNext
                End If
            Else
                ' not best move yet, but. . .
                
                If cmNext.dmgCharge > pk.cmStrongest.dmgCharge Then
                    pk.cmStrongest = cmNext
                End If
                
                If cmNext.factorBuff > pk.cmBestBuff.factorBuff Then
                    pk.cmBestBuff = cmNext
                End If
                
                If cmNext.cTurnsToCharge < pk.cmBestBuff.cTurnsToCharge Then
                    pk.cmQuickest = cmNext
                End If
            End If
            
            iNext = iNext + 1
            strNext = ParseMoveName(pk.csv, iNext)
    
        Wend
        
        pk.cm = pk.cmBest ' Use the best move
    
    End If
    
    ' determine threats , to be used in reporting.
    Call DetermineMoveThreat(pk.cmBest, pkDefender)
    If pk.cmStrongest.strMove <> pk.cmBest.strMove Then Call DetermineMoveThreat(pk.cmStrongest, pkDefender)
    
    With pk.cmQuickest
    If .strMove <> pk.cmBest.strMove And .strMove <> pk.cmStrongest.strMove Then Call DetermineMoveThreat(pk.cmQuickest, pkDefender)
    End With
    
    With pk.cmBestBuff
    If .strMove <> pk.cmBest.strMove And .strMove <> pk.cmStrongest.strMove And .strMove <> pk.cmQuickest.strMove Then Call DetermineMoveThreat(pk.cmBestBuff, pkDefender)
    End With
    
End Sub

Sub DetermineBestChargeMoveByType(pk As Pokemon, pkDefender As Pokemon)

    If pk.fMultipleChargeMoves Then
   
        Dim cmNext As ChargeMove, strNext As String, iNext As Integer
        
        iNext = 3
        cmNext.strMove = ParseMoveName(pk.csv, iNext)
            
        While cmNext.strMove <> ""
            cmNext.strType = TypeOfChargeMove(cmNext.strMove, pk.strType1)
            cmNext.factorMult = TypeEffectivenessMultiplier(cmNext.strType, pk, pkDefender)

            If cmNext.factorMult > pk.cm.factorMult Then
                pk.cm = cmNext
            End If
            
            iNext = iNext + 1
            cmNext.strMove = ParseMoveName(pk.csv, iNext)
        Wend
        
    End If
    
    pk.cmBest = pk.cm
    pk.cmBestBuff = pk.cm
    pk.cmStrongest = pk.cm
    pk.cmQuickest = pk.cm

End Sub

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

Sub AdjustForBuff(cm As ChargeMove, cmBuff As ChargeMove, _
ByVal cTurnsMaxBattle As Single, ByVal cTurnsInBattle As Single, _
ByRef dptAttackerQuick As Single, ByRef dptAttackerCharge As Single, ByRef defAttacker As Single, _
ByRef dptDefenderQuick As Single, ByRef dptDefenderCharge As Single, ByRef defDefender As Single)

    Dim valChanceOfBuff As Single, valChanceOfBuff_BuffMove As Single
    Dim cStagesAttackerAttack As Single, cStagesAttackerDefense As Single, cStagesDefenderAttack As Single, cStagesDefenderDefense As Single
    Dim buffAttackerAttack As Single, buffAttackerDefense As Single, buffDefenderAttack As Single, buffDefenderDefense As Single
    
    With cm.rngData
    
    valChanceOfBuff = .Cells(1, 6).value  'Percentage chance firing the charge move will cause a buff.
    
    If valChanceOfBuff > 0 Or cmBuff.strMove <> cm.strMove Then
        Dim cTurnsAfterBuff As Single
        Dim score As Single, scoreAlt As Single
        
        cStagesAttackerAttack = .Cells(1, 7) * valChanceOfBuff
        cStagesDefenderAttack = .Cells(1, 8) * valChanceOfBuff
        cStagesAttackerDefense = .Cells(1, 9) * valChanceOfBuff
        cStagesDefenderDefense = .Cells(1, 10) * valChanceOfBuff
        
        cm.strBuffSymbols = GetSpecialEffectSymbols(cm)
    
        'we have work to do!  A kind of turn by turn battle simulation is needed to understand buff effects.
        'values passed in to this subroutine ByRef will be adjusted proportiately to reflect nerfs happening during the battle.
    
        'Forgive the digression as we consider an alternative scenario using local data.
        
        scoreAlt = 0
        If cmBuff.strMove <> cm.strMove Then
            'If there is charge move available which has a 100% reliable buff, model the case where attacker uses a single
            'buff move first before switching to the charge move.
            
            'OR, if the strongest charge move has a 100% chance of a negative buff, such as Wild Charge or overheat, but there is another
            'move available which does not, explore using the non-self-destructing move first.
            
            'Model this using local data and calculate scoreAlt.  If scoreAlt is higher than the score would otherwise be,
            'commit to using this scenario by returning the modelled numbers ByRef to the caller.
            
            valChanceOfBuff_BuffMove = cmBuff.rngData.Cells(1, 6).value
        
            If valChanceOfBuff_BuffMove = 1 Or valChanceOfBuff = 1 Then
                ' strMoveBuff , by definition , has a BETTER Buff than cm.strMove.
                ' But, it might just be that cm.strMove has a Bad Buff, such as Overheat and Wild Charge.
                
                ' If valChanceOfBuff_BuffMove  = 1, there is a 100% chance that strBuffMove will provide a positive buff.
                ' If valChanceOfBuff = 1, there is a 100% chance that strBuffMove could DELAY the negative buff of cm.strMove.
                ' Either scenario is worth some time to explore.
                
                ' Actually, this code could handle any scenario, but we are restricting to most promising cases right now.
                
                If (cmBuff.cTurnsToCharge + cm.cTurnsToCharge < cTurnsInBattle) Then ' There is time to simulate a buff move.
                
                    Dim cStagesAttackerAttackBuff As Single, cStagesAttackerDefenseBuff As Single, cStagesDefenderAttackBuff As Single, cStagesDefenderDefenseBuff As Single
                    Dim dptAttackerQuick1 As Single, dptAttackerCharge1 As Single, defAttacker1 As Single
                    Dim dptDefenderQuick1 As Single, dptDefenderCharge1 As Single, defDefender1 As Single

                    cStagesAttackerAttackBuff = cmBuff.rngData.Cells(1, 7) * valChanceOfBuff_BuffMove
                    cStagesDefenderAttackBuff = cmBuff.rngData.Cells(1, 8) * valChanceOfBuff_BuffMove
                    cStagesAttackerDefenseBuff = cmBuff.rngData.Cells(1, 9) * valChanceOfBuff_BuffMove
                    cStagesDefenderDefenseBuff = cmBuff.rngData.Cells(1, 10) * valChanceOfBuff_BuffMove
                    
                    'blending in the buff attack is a bit tricker. Battle has two parts.
                    cTurnsAfterBuff = cTurnsInBattle - cmBuff.cTurnsToCharge
                        
                    'The rest of the stats get buffed by the buff move THEN by any buffs by the regular charge move
                    'applied proportionately as the battle progresses.
                    
                    dptAttackerQuick1 = dptAttackerQuick: dptAttackerCharge1 = dptAttackerCharge: defAttacker1 = defAttacker
                    dptDefenderQuick1 = dptDefenderQuick: dptDefenderCharge1 = dptDefenderCharge: defDefender1 = defDefender
                    buffAttackerAttack = 1: buffAttackerDefense = 1: buffDefenderAttack = 1: buffDefenderDefense = 1
                    
                    ' we used to just call AdjustStatForBuff for each stat. This code is trying to avoid unecessary calls.
                     If cStagesAttackerAttack <> 0 Or cStagesAttackerAttackBuff <> 0 Then
                        'first, the buff move buffs the charge move once.
                        dptAttackerCharge1 = dptAttackerCharge1 * BuffForStage(cStagesAttackerAttackBuff)
                        
                        'during the rest of the battle, the charge move may buff itself. Note we are using cTurnsAfterBuff and no buff move.
                        Call AdjustStatForBuff(dptAttackerCharge1, cStagesAttackerAttack, cm.cTurnsToCharge, 0, 0, cTurnsAfterBuff, True)
                        
                        ' quick attack adjustment is different in this case, since it is not being changed.
                        Call AdjustStatForBuff(buffAttackerAttack, cStagesAttackerAttack, cm.cTurnsToCharge, _
                            cStagesAttackerAttackBuff, cmBuff.cTurnsToCharge, cTurnsInBattle, True)
                        dptAttackerQuick1 = dptAttackerQuick1 * buffAttackerAttack
                    End If
                    
                    If cStagesAttackerDefense <> 0 Or cStagesAttackerDefenseBuff <> 0 Then
                        If cStagesAttackerDefense = cStagesAttackerAttack And cStagesAttackerDefenseBuff = cStagesAttackerAttackBuff Then
                            buffAttackerDefense = buffAttackerAttack
                        Else
                            Call AdjustStatForBuff(buffAttackerDefense, cStagesAttackerDefense, cm.cTurnsToCharge, _
                                cStagesAttackerDefenseBuff, cmBuff.cTurnsToCharge, cTurnsInBattle, True)
                        End If
                        
                        defAttacker1 = defAttacker1 * buffAttackerDefense
                    End If
                    
                    If cStagesDefenderAttack <> 0 Or cStagesDefenderAttackBuff Then
                        Call AdjustStatForBuff(buffDefenderAttack, cStagesDefenderAttack, cm.cTurnsToCharge, _
                            cStagesDefenderAttackBuff, cmBuff.cTurnsToCharge, cTurnsInBattle, False)
                        dptDefenderQuick1 = dptDefenderQuick1 * buffDefenderAttack
                        dptDefenderCharge1 = dptDefenderCharge1 * buffDefenderAttack
                    End If
                    
                    If cStagesDefenderDefense <> 0 Or cStagesDefenderDefenseBuff Then
                        If cStagesDefenderDefense = cStagesDefenderAttack And cStagesDefenderDefenseBuff = cStagesDefenderAttackBuff Then
                            buffDefenderDefense = buffDefenderAttack
                        Else
                            Call AdjustStatForBuff(buffDefenderDefense, cStagesDefenderDefense, cm.cTurnsToCharge, _
                                cStagesDefenderDefenseBuff, cmBuff.cTurnsToCharge, cTurnsInBattle, False)
                        End If
                        
                        defDefender1 = defDefender1 * buffDefenderDefense
                    End If
                    
                    dptAttackerCharge1 = WeightedAverage(cmBuff.dptCharge, dptAttackerCharge1, cmBuff.cTurnsToCharge / cTurnsInBattle)
                        
                    'Did the nerf move improve the attacker's score?  If so, keep it.
                    Dim damageByAttacker1 As Single, damageByDefender1 As Single
                    
                    damageByAttacker1 = (dptAttackerQuick1 + dptAttackerCharge1) / defDefender1
                    damageByDefender1 = (dptDefenderQuick1 + dptDefenderCharge1) / defAttacker1
                    scoreAlt = 1000 * (damageByAttacker1 / (damageByAttacker1 + damageByDefender1))
                End If
            End If

        End If
        
        'Main case - just step through the battle modifying ByRef values based on charge move buffs.
        
        If valChanceOfBuff > 0 Then
        
            buffAttackerAttack = 1: buffAttackerDefense = 1: buffDefenderAttack = 1: buffDefenderDefense = 1
            
            ' we used to just call AdjustStatForBuff 6 times, one for each stat. This code is trying to avoid unecessary calls.
            
            If cStagesAttackerAttack <> 0 Then
                Call AdjustStatForBuff(buffAttackerAttack, cStagesAttackerAttack, cm.cTurnsToCharge, _
                    0, 0, cTurnsInBattle, True)
                dptAttackerQuick = dptAttackerQuick * buffAttackerAttack
                dptAttackerCharge = dptAttackerCharge * buffAttackerAttack
            End If
            
            If cStagesAttackerDefense <> 0 Then
                If cStagesAttackerDefense = cStagesAttackerAttack Then
                    buffAttackerDefense = buffAttackerAttack
                Else
                    Call AdjustStatForBuff(buffAttackerDefense, cStagesAttackerDefense, cm.cTurnsToCharge, _
                        0, 0, cTurnsInBattle, True)
                End If
                
                defAttacker = defAttacker * buffAttackerDefense
            End If
            
            If cStagesDefenderAttack <> 0 Then
                Call AdjustStatForBuff(buffDefenderAttack, cStagesDefenderAttack, cm.cTurnsToCharge, _
                    0, 0, cTurnsInBattle, False)
                dptDefenderQuick = dptDefenderQuick * buffDefenderAttack
                dptDefenderCharge = dptDefenderCharge * buffDefenderAttack
            End If
            
            If cStagesDefenderDefense <> 0 Then
                If cStagesDefenderDefense = cStagesDefenderAttack Then
                    buffDefenderDefense = buffDefenderAttack
                Else
                    Call AdjustStatForBuff(buffDefenderDefense, cStagesDefenderDefense, cm.cTurnsToCharge, _
                        0, 0, cTurnsInBattle, False)
                End If
                
                defDefender = defDefender * buffDefenderDefense
            End If
        End If
        
        If scoreAlt > 0 Then
            'We have modelled an alternate scenario.
            'If this would improve our score, use this scenario.
            
            Dim damageByAttacker As Single, damageByDefender As Single
            
            damageByAttacker = (dptAttackerQuick + dptAttackerCharge) / defDefender
            damageByDefender = (dptDefenderQuick + dptDefenderCharge) / defAttacker
            score = 1000 * (damageByAttacker / (damageByAttacker + damageByDefender))
                        
            If scoreAlt > score Then
                'The nerf move was beneficial!  Keep it.
                
                dptAttackerQuick = dptAttackerQuick1: dptAttackerCharge = dptAttackerCharge1: defAttacker = defAttacker1
                dptDefenderQuick = dptDefenderQuick1: dptDefenderCharge = dptDefenderCharge1: defDefender = defDefender1
                
                cm.strMove = ChargeMoveAbbreviation(cmBuff.strMove) & "+" & ChargeMoveAbbreviation(cm.strMove)
                
                If valChanceOfBuff_BuffMove > 0 Then
                    cm.strBuffSymbols = GetSpecialEffectSymbols(cmBuff)
                    If valChanceOfBuff > 0 Then cm.strBuffSymbols = cm.strBuffSymbols & " + " & GetSpecialEffectSymbols(cm)
                End If
            End If
        End If
    End If
    
    End With
    
End Sub

Sub AdjustStatForBuff(ByRef stat As Single, _
    cStagesPerChargeMove As Single, cTurnsPerChargeMove As Single, _
    cStagesPerBuffMove As Single, cTurnsPerBuffMove As Single, _
    cTurnsInBattle As Single, fAttackerStat As Boolean)

    If cStagesPerChargeMove <> 0 Or cStagesPerBuffMove <> 0 Then
        Dim currentBuff As Single, newBuff As Single
        Dim valStage As Single
        Dim cTurnsSoFar As Single
        Dim statSav As Single
        
        statSav = stat
    
        valStage = 0
        currentBuff = 1
        
        cTurnsSoFar = cTurnsPerBuffMove
        
        If cStagesPerBuffMove <> 0 And cTurnsInBattle > cTurnsSoFar Then
            ' Insert a Single Buff Move
        
            valStage = cStagesPerBuffMove
            newBuff = BuffForStage(valStage)
            
            ' weighted average ratioing buff move buff only to turns after it occurs
            stat = WeightedAverage(stat, stat * newBuff, cTurnsSoFar / cTurnsInBattle)
            
            currentBuff = newBuff
        End If
        
        ' problem with parameters getting past in wrong order.  Detect if it happens again.
        If cTurnsPerChargeMove <= 1 Then
            MsgBox "cTurnsPerChargeMove too small."
            Exit Sub
        End If
        
        If cStagesPerChargeMove <> 0 Then
            While (cTurnsInBattle - cTurnsSoFar > cTurnsPerChargeMove)
        
                valStage = valStage + cStagesPerChargeMove
                newBuff = BuffForStage(valStage)
                
                If newBuff = currentBuff Then Exit Sub ' no more buffing.
            
                cTurnsSoFar = cTurnsSoFar + cTurnsPerChargeMove
                
                ' weighted average.
                stat = WeightedAverage(stat, stat * newBuff / currentBuff, cTurnsSoFar / cTurnsInBattle)
                
                currentBuff = newBuff

            Wend
        End If
        
        If fAttackerStat And cStagesPerChargeMove < 0 Then
            Dim statAlternativeDebuff As Single
        
            ' To discourage Pokemon from using moves that harm their own stats, apply at least a minimum debuff,
            ' as if the move were used once in the second half of the battle.
            
            'Pokemon - Please consider that if you use Wild Charge, Overheat, Draco Meteor to win this battle, you will be weak for the next oponent.
            'Also, if you weaken yourself  and the oponent blocks it, you will be sorry.
            
 '           statAlternativeDebuff = statSav * BuffForStage(cStagesPerBuff / 2)
            statAlternativeDebuff = 0.6 * statSav + 0.4 * (statSav * BuffForStage(cStagesPerChargeMove))
            If statAlternativeDebuff < stat Then
                 stat = statAlternativeDebuff
            End If
        End If
        
    End If

End Sub


Function BuffForStage(valStage As Single) As Single

' "minimumStatStage": -4,
' "maximumStatStage": 4,
' "attackBuffMultiplier": [0.33, 0.4, 0.5, 0.67, 1.0, 1.5, 2.0, 2.5, 3.0],
' "defenseBuffMultiplier": [0.33, 0.4, 0.5, 0.67, 1.0, 1.5, 2.0, 2.5, 3.0]

'Negative values are called Nerfs.

    If valStage >= 4 Then
        BuffForStage = 3
    ElseIf valStage <= -4 Then
        BuffForStage = 1 / 3
    Else
        Dim wholeStage As Single, fractStage As Single
        
        wholeStage = RoundDown(valStage)
        fractStage = valStage - wholeStage
    
        Select Case wholeStage
            Case -3:
                BuffForStage = 0.4 + fractStage * 0.066 'fractStage is negative for negative values
            Case -2:
                BuffForStage = 0.5 + fractStage * 0.1
            Case -1:
                BuffForStage = 2 / 3 + fractStage * (1 / 6)
            Case 0:
                If valStage < 0 Then
                    BuffForStage = 1 + fractStage * (1 / 3)
                Else
                    BuffForStage = 1 + fractStage * 0.5
                End If
            Case 1:
                BuffForStage = 1.5 + fractStage * 0.5
            Case 2:
                BuffForStage = 2 + fractStage * 0.5
            Case 3:
                BuffForStage = 2.5 + fractStage * 0.5
        End Select
    End If
                
End Function

Sub TrimWastedDamage(pk As Pokemon, pkDefender As Pokemon)

    ' Trim excess damage from the last instance of the charge move.  This greatly effects the score created by a charge move that overkills.
    
    ' Note that this is in addition to TimeFactor, which penalized ALL charge moves by a fixed factor And a time factor to account
    ' for the huge chance the move would get blocked or match would end before firing.  This is specifically for overkill wasted damage.
    
    Dim hpChargePerAttack As Single, hpChargeNeededLastAttack As Single
    Dim cTurnsToCharge As Single, cCompletedAttacks As Single
    
    With pk.cm
    
    hpChargePerAttack = .hpptCharge * .cTurnsToCharge
    
    cCompletedAttacks = RoundDown(pkDefender.bstat.hp / hpChargePerAttack)
    hpChargeNeededLastAttack = pkDefender.bstat.hp - cCompletedAttacks * hpChargePerAttack
    .hpptCharge = WeightedAverage(.hpptCharge, hpChargeNeededLastAttack / .cTurnsToCharge, cCompletedAttacks / (cCompletedAttacks + 1))
    
    End With
    
If False Then

    ' Since you can't do a charge attack without quick attacks along the way, so why not trim more completely?
    ' In practice this could distort results. You CAN do a charge attack without quick attacks if you come into a battle already
    ' having energy, and the more severe trimming exagerates small differences in timing.
    
    Dim hpQuickPerAttack As Single, hpPerAttack As Single

    With pk.cm
    
    hpQuickPerAttack = pk.qm.hpptQuick * .cTurnsToCharge
    hpChargePerAttack = .hpptCharge * .cTurnsToCharge
    hpPerAttack = hpQuickPerAttack + hpChargePerAttack
    
    cCompletedAttacks = RoundDown(pkDefender.bstat.hp / hpPerAttack)
    hpChargeNeededLastAttack = Max(pkDefender.bstat.hp - cCompletedAttacks * hpPerAttack - hpQuickPerAttack, 0)
    .hpptCharge = WeightedAverage(.hpptCharge, hpChargeNeededLastAttack / .cTurnsToCharge, cCompletedAttacks / (cCompletedAttacks + 1))
    
    End With
End If
End Sub

Sub RewardFirstMoveAdvantage(pk1 As Pokemon, pk2 As Pokemon)

    If pk1.cm.cTurnsToCharge = pk2.cmQuickest.cTurnsToCharge Then
        If pk1.bstat.attCMP > pk2.bstat.attCMP + 15 Then GoTo BonusForPk1:
    ElseIf pk1.cm.cTurnsToCharge < pk2.cmQuickest.cTurnsToCharge Then
BonusForPk1:
        With pk1.cm
            .hpptFirstMoveAdvantage = ((.hpptCharge * .cTurnsToCharge / 4) / .cTurnsToVictory)
            .hpptCharge = .hpptCharge + .hpptFirstMoveAdvantage
        End With
    ElseIf pk2.cm.cTurnsToCharge = pk1.cmQuickest.cTurnsToCharge Then
        If pk2.bstat.attCMP > pk1.bstat.attCMP + 15 Then GoTo BonusForPk2:
    ElseIf pk2.cm.cTurnsToCharge < pk1.cmQuickest.cTurnsToCharge Then
BonusForPk2:
        With pk2.cm
            .hpptFirstMoveAdvantage = ((.hpptCharge * .cTurnsToCharge / 4) / .cTurnsToVictory)
            .hpptCharge = .hpptCharge + .hpptFirstMoveAdvantage
        End With
    End If
        
End Sub

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
    

Sub DetermineMoveThreat(cm As ChargeMove, pkDefender As Pokemon)
    ' Important Statistics for User to Understand Battle Strategy
    Dim pctAdjusted As Integer, timeFactorAdjusted As Single

    cm.threat.pctDamage = 100 * CSng(cm.hpPerCharge) / CSng(pkDefender.bstat.hp) ' force floating point math.
    
    timeFactorAdjusted = Average(cm.factorTime, 1)  ' compromise:  Reduce by time factor, but not as much. Slow moves are still dangerous.
    pctAdjusted = MinI(100, cm.threat.pctDamage)     ' killing me more than once doesn't increase the threat.
    
    ' now boil this down to a number between 1 and 10 that represents the need to use a shield.
    cm.threat.threatLevel = RoundDown(MinMax((pctAdjusted * timeFactorAdjusted / 8), 1, 10))
        
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
        If SymbolForType(pk.strNameData) <> "?" Then
            pk.strType1 = pk.strNameData  ' Pokemon name IS a type, like Fighting.  Just return it so people can treat types as Pokemon.
        Else
            pk.strType1 = "Unknown"
        End If
        
        pk.strType2 = ""
    Else
        pk.strType1 = pk.rngData.Cells(1, pkData_Type1)
        pk.strType2 = pk.rngData.Cells(1, pkData_Type2)
    End If

    Call InitQuickMove(pk.qm, strQuickMove, pk.strType1)
    Call InitChargeMove(pk.cm, pk.qm, strChargeMove, pk.strType1)
    
    If pk.strType1 = "Unknown" Then
        pk.fInvalid = True
        pk.fInvalidPokemon = True
    ElseIf pk.qm.strType = "Unknown" Then
        pk.fInvalid = True
        pk.fInvalidQuickMove = True
    ElseIf pk.rngData Is Nothing Then
        pk.fTypeMuse = True
    ElseIf pk.qm.rngData Is Nothing Then
        pk.fInvalid = True
        pk.fInvalidQuickMove = True
    ElseIf pk.cm.rngData Is Nothing Then
        pk.fInvalid = True
        pk.fInvalidChargeMove = True
    End If
    
    strCategory = GetPkString(pk, pkData_Category)
    If strCategory = "M" Or strCategory = "L" Then pk.fLegendaryOrMythical = True
    
End Sub

' Qualify a Pokemon for battle.

Sub QualifyPokemon(pk As Pokemon, csvIV As String, ByVal rngDataBattleLeague As Range, fTypeMuseOK As Boolean)
    Dim cpMaxBattleLeague As Integer

    If pk.fInvalid Then Exit Sub
    
    cpMaxBattleLeague = GetDataInteger(rngDataBattleLeague, blData_MaxCp)
    If cpMaxBattleLeague = 0 Then
        ' Nothing to do to qualify for Type Effeciveness Battle.
        pk.fTypeEffectivenessBattle = True
        pk.fQualified = True
        Exit Sub
    End If
    
    If pk.fTypeMuse Then
        ' The pokemon is a Type Muse. How would our Pokemon do against a generic pokemon of a specific type with specific type moves?
        ' Generally, we allow type muses in the meta but not on our team.
        
        pk.fQualified = fTypeMuseOK  ' Team members can't be type prototypes, but meta muses can be
        pk.fTypeEffectivenessBattle = True
        Exit Sub
    End If
    
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

Function StrValidatePk(pk As Pokemon) As String
    StrValidatePk = ""
    If pk.fInvalid Then
        If pk.fInvalidPokemon Then
            StrValidatePk = "Not A Pokemon"
        ElseIf Not pk.fTypeMuse Then
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

Sub InitQuickMove(qm As QuickMove, strQuickMove As String, Optional strDefaultType As String = "")
    Dim qmInit As QuickMove
    
    qm = qmInit ' clear data
    qm.strMove = strQuickMove
    
    If qm.strMove = "" Then
        qm.strType = strDefaultType
    Else
        Set qm.rngData = GetQuickMoveData(qm.strMove)
    
        If qm.rngData Is Nothing Then
            If SymbolForType(qm.strMove) <> "?" Then
                qm.strType = qm.strMove  ' Attack name IS a type, like Fighting.  Just return it so people can use pseudo-attacks
            Else
                qm.strType = "Unknown"
            End If
        Else
            qm.strType = qm.rngData.Cells(1, 2).value
            qm.cTurnsToQuick = GetDataValue(qm.rngData, 5)
        End If
    End If

End Sub

Sub InitChargeMove(cm As ChargeMove, qm As QuickMove, strChargeMove As String, Optional strDefaultType As String = "")
    Dim cmInit As ChargeMove
    
    cm = cmInit ' clear data
    cm.strMove = strChargeMove
    
    If cm.strMove = "" Then
        cm.strType = strDefaultType
    Else
        Set cm.rngData = GetChargeMoveData(cm.strMove)
    
        If cm.rngData Is Nothing Then
            If SymbolForType(cm.strMove) <> "?" Then
                cm.strType = cm.strMove  ' Attack name IS a type, like Fighting.  Just return it so people can use pseudo-attacks
            Else
                cm.strType = "Unknown"
            End If
        Else
            cm.strType = cm.rngData.Cells(1, 2).value
            cm.cTurnsToCharge = CTurnsToChargeMove(cm, qm)
        End If
    End If

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

Function BuffFactor(cm As ChargeMove, cmDefender As ChargeMove, ByVal cTurnsMaxBattle As Single) As Single
    Dim cTurnsInBattle As Single
    Dim dptAttackerQuick As Single, dptAttackerCharge As Single, defAttacker As Single
    Dim dptDefenderQuick As Single, dptDefenderCharge As Single, defDefender As Single
    Dim breakPoint As Integer
    
    BuffFactor = 1 ' Default answer indicating no time factor impact, or no battle details available.
    
    If cTurnsMaxBattle > 0 And IsStatAlteringChargeMove(cm) Then
        dptAttackerQuick = 1: dptAttackerCharge = 1: defAttacker = 1
        dptDefenderQuick = 1: dptDefenderCharge = 1: defDefender = 1
        
        cTurnsInBattle = Min(cm.cTurnsToVictory, cmDefender.cTurnsToVictory)

        Call AdjustForBuff(cm, cm, cTurnsMaxBattle, cTurnsInBattle, _
            dptAttackerQuick, dptAttackerCharge, defAttacker, dptDefenderQuick, dptDefenderCharge, defDefender)
                        
        BuffFactor = ((dptAttackerQuick + dptAttackerCharge) * defAttacker) / ((dptDefenderQuick + dptDefenderCharge) * defDefender)
    End If

End Function







