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


    Dim cTurnsMaxBattle As Single

    BattleScoreCore = -1 'error
    
    'We need two qualified pokemon!
    
    If Not (pk1.fQualified And pk2.fQualified) Then Exit Function
        

    
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
        
        ' Apply opponent buff first so our own decision about whether to combine moves can take that into account.
        Call AdjustForBuff(pk2, pk1, cTurnsMaxBattle, cTurnsInBattle)
    
        Call AdjustForBuff(pk1, pk2, cTurnsMaxBattle, cTurnsInBattle)

    End If
    
    If False Then ' This is tending to make all moves seem the same. Good moves are penalized more than bad ones to equalize them.
    
        ' Limit cm.hpptCharge to reduce extreme wasted energy caused by overkill from distorting recommendations.
        Call TrimWastedDamage(pk1, pk2)
        Call TrimWastedDamage(pk2, pk1)
        
    End If
    
    Call RewardFirstMoveAdvantage(pk1, pk2)
    
    ' Final calculations and score
    
    pk1.cm.cTurnsToVictory = RoundUpTurnsQm(pk2.bstat.hp / (pk1.qm.hpptQuick + pk1.cm.hpptCharge), pk1.qm)
    pk2.cm.cTurnsToVictory = RoundUpTurnsQm(pk1.bstat.hp / (pk2.qm.hpptQuick + pk2.cm.hpptCharge), pk2.qm)
    
    'Score is a per-thousand ratio ratio of damage by attacker to combined damage. A score of 500 denotes a tie.
    
    BattleScoreCore = 1000 * (pk2.cm.cTurnsToVictory / (pk1.cm.cTurnsToVictory + pk2.cm.cTurnsToVictory))
        
    
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


Sub AdjustForBuff(pk As Pokemon, pkDefender As Pokemon, ByVal cTurnsMaxBattle As Single, ByVal cTurnsInBattle As Single)

    Dim valChanceOfBuff As Single, valChanceOfBuff_BuffMove As Single
    Dim cStagesAttackerAttack As Single, cStagesAttackerDefense As Single, cStagesDefenderAttack As Single, cStagesDefenderDefense As Single
    Dim buffAttackerAttack As Single, buffAttackerDefense As Single, buffDefenderAttack As Single, buffDefenderDefense As Single
    Dim cm As ChargeMove, cmBuff As ChargeMove
    
    cm = pk.cm
    cmBuff = pk.cmBestBuff
    
    With cm.rngData
    
    valChanceOfBuff = .Cells(1, 6).value  'Percentage chance firing the charge move will cause a buff.
    
    If valChanceOfBuff > 0 Or cmBuff.strMove <> cm.strMove Then
        Dim cTurnsAfterBuff As Single
        Dim score As Integer, scoreAlt As Integer
        
        cStagesAttackerAttack = .Cells(1, 7) * valChanceOfBuff
        cStagesDefenderAttack = .Cells(1, 8) * valChanceOfBuff
        cStagesAttackerDefense = .Cells(1, 9) * valChanceOfBuff
        cStagesDefenderDefense = .Cells(1, 10) * valChanceOfBuff
        
        pk.cm.strBuffSymbols = GetSpecialEffectSymbols(cm)
    
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
                    Dim pkAlt As Pokemon, pkDefAlt As Pokemon

                    cStagesAttackerAttackBuff = cmBuff.rngData.Cells(1, 7) * valChanceOfBuff_BuffMove
                    cStagesDefenderAttackBuff = cmBuff.rngData.Cells(1, 8) * valChanceOfBuff_BuffMove
                    cStagesAttackerDefenseBuff = cmBuff.rngData.Cells(1, 9) * valChanceOfBuff_BuffMove
                    cStagesDefenderDefenseBuff = cmBuff.rngData.Cells(1, 10) * valChanceOfBuff_BuffMove
                    
                    'blending in the buff attack is a bit tricker. Battle has two parts.
                    cTurnsAfterBuff = cTurnsInBattle - cmBuff.cTurnsToCharge
                        
                    'The rest of the stats get buffed by the buff move THEN by any buffs by the regular charge move
                    'applied proportionately as the battle progresses.
                    
                    pkAlt = pk
                    pkDefAlt = pkDefender
                    
                    buffAttackerAttack = 1: buffAttackerDefense = 1: buffDefenderAttack = 1: buffDefenderDefense = 1
                    
                    ' we used to just call AdjustStatForBuff for each stat. This code is trying to avoid unecessary calls.
                     If cStagesAttackerAttack <> 0 Or cStagesAttackerAttackBuff <> 0 Then
                        'first, the buff move buffs the charge move once.
                        pkAlt.cm.dptCharge = pkAlt.cm.dptCharge * BuffForStage(cStagesAttackerAttackBuff)
                        
                        'during the rest of the battle, the charge move may buff itself. Note we are using cTurnsAfterBuff and no buff move.
                        Call AdjustStatForBuff(pkAlt.cm.dptCharge, cStagesAttackerAttack, cm.cTurnsToCharge, 0, 0, cTurnsAfterBuff, True)
                        
                        ' quick attack adjustment is different in this case, since it is not being changed.
                        Call AdjustStatForBuff(buffAttackerAttack, cStagesAttackerAttack, cm.cTurnsToCharge, _
                            cStagesAttackerAttackBuff, cmBuff.cTurnsToCharge, cTurnsInBattle, True)
                            
                        pkAlt.qm.dptQuick = pkAlt.qm.dptQuick * buffAttackerAttack
                    End If
                    
                    If cStagesAttackerDefense <> 0 Or cStagesAttackerDefenseBuff <> 0 Then
                        If cStagesAttackerDefense = cStagesAttackerAttack And cStagesAttackerDefenseBuff = cStagesAttackerAttackBuff Then
                            buffAttackerDefense = buffAttackerAttack
                        Else
                            Call AdjustStatForBuff(buffAttackerDefense, cStagesAttackerDefense, cm.cTurnsToCharge, _
                                cStagesAttackerDefenseBuff, cmBuff.cTurnsToCharge, cTurnsInBattle, True)
                        End If
                        
                        pkAlt.bstat.def = pkAlt.bstat.def * buffAttackerDefense
                    End If
                    
                    If cStagesDefenderAttack <> 0 Or cStagesDefenderAttackBuff Then
                        Call AdjustStatForBuff(buffDefenderAttack, cStagesDefenderAttack, cm.cTurnsToCharge, _
                            cStagesDefenderAttackBuff, cmBuff.cTurnsToCharge, cTurnsInBattle, False)
                        pkDefAlt.qm.dptQuick = pkDefAlt.qm.dptQuick * buffDefenderAttack
                        pkDefAlt.cm.dptCharge = pkDefAlt.cm.dptCharge * buffDefenderAttack
                    End If
                    
                    If cStagesDefenderDefense <> 0 Or cStagesDefenderDefenseBuff Then
                        If cStagesDefenderDefense = cStagesDefenderAttack And cStagesDefenderDefenseBuff = cStagesDefenderAttackBuff Then
                            buffDefenderDefense = buffDefenderAttack
                        Else
                            Call AdjustStatForBuff(buffDefenderDefense, cStagesDefenderDefense, cm.cTurnsToCharge, _
                                cStagesDefenderDefenseBuff, cmBuff.cTurnsToCharge, cTurnsInBattle, False)
                        End If
                        
                        pkDefAlt.bstat.def = pkDefAlt.bstat.def * buffDefenderDefense
                    End If
                    
                    pkAlt.cm.dptCharge = WeightedAverage(cmBuff.dptCharge, pkAlt.cm.dptCharge, cmBuff.cTurnsToCharge / cTurnsInBattle)
                    
                    ' Requantize - So that each descrete move takes an integer number of hp points post buff.
                    pkAlt.qm.hpptQuick = HpPerTurnQm(pkAlt.qm, pkDefAlt.bstat.def)
                    pkAlt.cm.hpptCharge = HpPerTurnCm(pkAlt.cm, pkDefAlt.bstat.def)
                    pkDefAlt.qm.hpptQuick = HpPerTurnQm(pkDefAlt.qm, pkAlt.bstat.def)
                    pkDefAlt.cm.hpptCharge = HpPerTurnCm(pkDefAlt.cm, pkAlt.bstat.def)
                    
                    ' Alternate score.  Only use if better .
                    
                    pkAlt.cm.cTurnsToVictory = RoundUpTurnsQm(pkDefAlt.bstat.hp / (pkAlt.qm.hpptQuick + pkAlt.cm.hpptCharge), pkAlt.qm)
                    pkDefAlt.cm.cTurnsToVictory = RoundUpTurnsQm(pkAlt.bstat.hp / (pkDefAlt.qm.hpptQuick + pkDefAlt.cm.hpptCharge), pkDefAlt.qm)
                    scoreAlt = 1000 * (pkDefAlt.cm.cTurnsToVictory / (pkAlt.cm.cTurnsToVictory + pkDefAlt.cm.cTurnsToVictory))
        
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
                pk.qm.dptQuick = pk.qm.dptQuick * buffAttackerAttack
                pk.cm.dptCharge = pk.cm.dptCharge * buffAttackerAttack
            End If
            
            If cStagesAttackerDefense <> 0 Then
                If cStagesAttackerDefense = cStagesAttackerAttack Then
                    buffAttackerDefense = buffAttackerAttack
                Else
                    Call AdjustStatForBuff(buffAttackerDefense, cStagesAttackerDefense, cm.cTurnsToCharge, _
                        0, 0, cTurnsInBattle, True)
                End If
                
                pk.bstat.def = pk.bstat.def * buffAttackerDefense
            End If
            
            If cStagesDefenderAttack <> 0 Then
                Call AdjustStatForBuff(buffDefenderAttack, cStagesDefenderAttack, cm.cTurnsToCharge, _
                    0, 0, cTurnsInBattle, False)
                pkDefender.qm.dptQuick = pkDefender.qm.dptQuick * buffDefenderAttack
                pkDefender.cm.dptCharge = pkDefender.cm.dptCharge * buffDefenderAttack
            End If
            
            If cStagesDefenderDefense <> 0 Then
                If cStagesDefenderDefense = cStagesDefenderAttack Then
                    buffDefenderDefense = buffDefenderAttack
                Else
                    Call AdjustStatForBuff(buffDefenderDefense, cStagesDefenderDefense, cm.cTurnsToCharge, _
                        0, 0, cTurnsInBattle, False)
                End If
                
                pkDefender.bstat.def = pkDefender.bstat.def * buffDefenderDefense
            End If
            
            ' Requantize - So that each descrete move takes an integer number of hp points post buff.
            pk.qm.hpptQuick = HpPerTurnQm(pk.qm, pkDefender.bstat.def)
            pk.cm.hpptCharge = HpPerTurnCm(pk.cm, pkDefender.bstat.def)
            pkDefender.qm.hpptQuick = HpPerTurnQm(pkDefender.qm, pk.bstat.def)
            pkDefender.cm.hpptCharge = HpPerTurnCm(pkDefender.cm, pk.bstat.def)
            
        End If
        
        If scoreAlt > 0 Then
            'We have modelled an alternate scenario.
            'If this would improve our score, use this scenario.
                    
            pk.cm.cTurnsToVictory = RoundUpTurnsQm(pkDefender.bstat.hp / (pk.qm.hpptQuick + pk.cm.hpptCharge), pk.qm)
            pkDefender.cm.cTurnsToVictory = RoundUpTurnsQm(pk.bstat.hp / (pkDefender.qm.hpptQuick + pkDefender.cm.hpptCharge), pkDefender.qm)
            score = 1000 * (pkDefender.cm.cTurnsToVictory / (pk.cm.cTurnsToVictory + pkDefender.cm.cTurnsToVictory))
                        
            If scoreAlt > score Then
                'The nerf move was beneficial!  Keep it.
                
                pk = pkAlt
                pkDefender = pkDefAlt
                
                pk.cm.strMove = ChargeMoveAbbreviation(cmBuff.strMove) & "+" & ChargeMoveAbbreviation(cm.strMove)
                
                If valChanceOfBuff_BuffMove > 0 Then
                    pk.cm.strBuffSymbols = GetSpecialEffectSymbols(cmBuff)
                    If valChanceOfBuff > 0 Then pk.cm.strBuffSymbols = pk.cm.strBuffSymbols & " + " & GetSpecialEffectSymbols(cm)
                End If
            End If
        End If
    End If
    
    End With
    
End Sub

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

        Call AdjustForBuffForFactor(cm, cTurnsMaxBattle, cTurnsInBattle, _
            dptAttackerQuick, dptAttackerCharge, defAttacker, dptDefenderQuick, dptDefenderCharge, defDefender)
                        
        BuffFactor = ((dptAttackerQuick + dptAttackerCharge) * defAttacker) / ((dptDefenderQuick + dptDefenderCharge) * defDefender)
    End If

End Function

Sub AdjustForBuffForFactor(cm As ChargeMove, _
ByVal cTurnsMaxBattle As Single, ByVal cTurnsInBattle As Single, _
ByRef dptAttackerQuick As Single, ByRef dptAttackerCharge As Single, ByRef defAttacker As Single, _
ByRef dptDefenderQuick As Single, ByRef dptDefenderCharge As Single, ByRef defDefender As Single)

    'simplified version of AdjustForBuff used in BuffFactor.  It was just to hard to maintain the main without breaking
    'BuffFactor's very different needs.

    Dim valChanceOfBuff As Single
    Dim cStagesAttackerAttack As Single, cStagesAttackerDefense As Single, cStagesDefenderAttack As Single, cStagesDefenderDefense As Single
    Dim buffAttackerAttack As Single, buffAttackerDefense As Single, buffDefenderAttack As Single, buffDefenderDefense As Single
    
    With cm.rngData
    
    valChanceOfBuff = .Cells(1, 6).value  'Percentage chance firing the charge move will cause a buff.
    
    If valChanceOfBuff > 0 Then
        Dim cTurnsAfterBuff As Single
        Dim score As Single, scoreAlt As Single
        
        cStagesAttackerAttack = .Cells(1, 7) * valChanceOfBuff
        cStagesDefenderAttack = .Cells(1, 8) * valChanceOfBuff
        cStagesAttackerDefense = .Cells(1, 9) * valChanceOfBuff
        cStagesDefenderDefense = .Cells(1, 10) * valChanceOfBuff
    
        'we have work to do!  A kind of turn by turn battle simulation is needed to understand buff effects.
        'values passed in to this subroutine ByRef will be adjusted proportiately to reflect nerfs happening during the battle.
        
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
    

Sub DetermineMoveThreat(cm As ChargeMove, pkDefender As Pokemon)
    ' Important Statistics for User to Understand Battle Strategy
    Dim pctAdjusted As Integer, timeFactorAdjusted As Single

    cm.threat.pctDamage = 100 * CSng(cm.hpPerCharge) / CSng(pkDefender.bstat.hp) ' force floating point math.
    
    timeFactorAdjusted = Average(cm.factorTime, 1)  ' compromise:  Reduce by time factor, but not as much. Slow moves are still dangerous.
    pctAdjusted = MinI(100, cm.threat.pctDamage)     ' killing me more than once doesn't increase the threat.
    
    ' now boil this down to a number between 1 and 10 that represents the need to use a shield.
    cm.threat.threatLevel = RoundDown(MinMax((pctAdjusted * timeFactorAdjusted / 8), 1, 10))
        
End Sub













