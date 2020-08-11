from pkb_types import *

def CalcPokemonStats(pk:Pokemon, cpMax):

    # level and stats of each pokemon
    # Initial stats are preserved in cases where stat changing attacks may change the stats during battle.

    valAttack = int(pk.tupleData[pkData_Attack] + pk.ivs.attack)
    valDefense = int(pk.tupleData[pkData_Defense] + pk.ivs.defense)
    valStamina = int(pk.tupleData[pkData_Stamina] + pk.ivs.stamina)

    bstat = BattleStats()

    bstat.level = pk.ivs.levelMax
    while bstat.level >= 1.5:
        bstat.cp = CpAtLevelFromStats(bstat.level, valAttack, valDefense, valStamina)
        if bstat.cp <= cpMax:
            break
        bstat.level -= 0.5

    statMult = StatMultiplier(bstat.level)

    bstat.attCMP = statMult * valAttack
    bstat.attInit = bstat.attCMP

    bstat.defInit = statMult * valDefense

    if pk.fShadow:
        bstat.attInit = (6 / 5) * bstat.attInit
        bstat.defInit = (5 / 6) * bstat.defInit

    bstat.attack = bstat.attInit
    bstat.defense = bstat.defInit

    bstat.hp = RoundDown(statMult * valStamina)

    pk.bstat = bstat

def StatMultiplier(level):
    tupleRow= TupleFromTable(xlf.tableStatMultiplier, level)
    return ValFromTuple(tupleRow, 1, 1)

def CpAtLevelFromStats(level, valAttack, valDefense, valStamina):
    # ' see https://gamepress.gg/pokemongo/pokemon-stats-advanced

    # valAttack, valDefense, valStamina are the pokemon's base stat plus the related IV of the pokemon.
    # For a perfect pokemon, add 15 to each base stat.

    mult = StatMultiplier(level)
    return RoundDown((valAttack * math.sqrt(valDefense) * math.sqrt(valStamina) * mult * mult) / 10)

class BattleLeague:
    def __init__(self, strBattleLeague):
        self.strBattleLeague = strBattleLeague

        tupleData = TupleFromTable(xlf.tableBattleLeagues, strBattleLeague)
        self.strBattleLeague = StrFromTuple(tupleData, 0)
        self.cpMax = ValFromTuple(tupleData, 1)
        self.strRestriction = StrFromTuple(tupleData, 2)

def QualifyPokemon(pk: Pokemon, bl: BattleLeague, fTypeMuseOK=False):

    if pk.fInvalid:
        return

    if bl.cpMax == 0:
        # Nothing to do to qualify for Type Effectiveness Battle.
        pk.fTypeEffectivenessBattle = True
        pk.fQualified = True
        return

    if pk.fTypeMuse:
        # The pokemon is a Type Muse. How would our Pokemon do against a generic pokemon of a
        # specific type with specific type moves?
        # Generally, we allow type muses in the meta but not on our team.

        pk.fQualified = fTypeMuseOK  # Team members can't be type prototypes, but meta muses can be
        pk.fTypeEffectivenessBattle = True
        return

    pk.fQualified = True

    # Qualified unless the League has restrictions

    if bl.strRestriction == "Premier":
        if pk.fLegendaryOrMythical:
            pk.fQualified = False
    elif bl.strRestriction == "Flying":
        if pk.strType1 != "Flying" and pk.strType2 != "Flying":
            pk.fQualified = False

    CalcPokemonStats(pk, bl.cpMax)



















