from pkb_stats import *

assert "Up" != "Down", 'We Use Asserts to Test Our Assumptions'

print("\nTypes and Symbols:")
print(TypeOfChargeMove("Dynamic Punch"))
print(TypeOfQuickMove("Charm"))
print(SymbolForType("Fighting"))
print(SymbolForType("Ground"))

print("\nQuickMove:")
qm = QuickMove("Charm")
print(qm.strMove)
print(qm.strType)
print(qm.valDamage)
print(qm.valEnergy)
print(qm.cTurnsToQuick)
print(qm.RoundUpTurns(5))

print("\nChargeMove:")
cm = ChargeMove("Dynamic Punch", qm)
print(cm.strMove)
print(cm.strType)
print(cm.valEnergy)
print(cm.cTurnsToCharge)

print("\nUtilities:")
print(MinMax(5, 10, 20))
print(WeightedAverage(1, 2, 2/3))
print("\nTypeEffectiveness:")
print(TypeEffectiveness("Fire", "Water"))
print(TypeEffectiveness("Ground", "Flying"))

print("\nPokemon:")
pk = Pokemon("Togekiss, Charm, ancient_POWER, Flamethrower", "15, 15, 12")
print("name: " + pk.strName)
print("types: %s , %s" % (pk.strType1,pk.strType2))
print("default moves: %s, %s" % (pk.qm.strMove, pk.cm.strMove))

print("\nBattleLeague:")
bl = BattleLeague("ML Flying Cup")
print("battle league: %s: " % bl.strBattleLeague)
print("cpMax: %d" % bl.cpMax)
print("restriction: %s" % bl.strRestriction)

print("\nQualifyPokemon:")
QualifyPokemon(pk, bl, False)
string = pk.strName + " is fighting in " + bl.strBattleLeague
print(string)
print("level: %.1f" % pk.bstat.level)
print("cp: %d" % pk.bstat.cp)
print("hp: %d" % pk.bstat.hp)

print("\nTypeEffectiveness: ")
pkDefender = Pokemon("Metagross, Bullet Punch, Meteor Mash")
mult = TypeEffectivenessMultiplier(pk.qm.strType, pk, pkDefender)
print("%s using %s has type effectiveness multiplier of %.2f against %s"
      % (pk.strName, pk.qm.strMove, mult, pkDefender.strName))
mult = TypeEffectivenessMultiplier(pk.cm.strType, pk, pkDefender)
print("%s using %s has type effectiveness multiplier of %.2f against %s"
      % (pk.strName, pk.cm.strMove, mult, pkDefender.strName))
mult = TypeEffectivenessMultiplier("Fire", pk, pkDefender)
print("%s using %s has type effectiveness multiplier of %.2f against %s"
      % (pk.strName, "Flamethrower", mult, pkDefender.strName))


