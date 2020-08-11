from pkb_pokemon import *

# This module implements Type Effectiveness

def TypeEffectivenessMultiplier(strMoveType, pk: Pokemon, pkDefender: Pokemon):

    # These scores are the attack strength multipliers used by Pokemon Go.
    # The highest would be 3.072, if the move is doubly effective against the defender and has 20% STAB.

    te= TypeEffectiveness(strMoveType, pkDefender.strType1) + TypeEffectiveness(strMoveType, pkDefender.strType2)

    mult = 1.6 ** float(te)
    # Extreme case is Tropius, where value ranges from 0.244 to 2.56

    if strMoveType == pk.strType1 or strMoveType == pk.strType2:
        mult = mult * 1.2   # add S.T.A.B

    # Including stab, maximum possible value is 3.072
    return mult


def TypeEffectiveness(strTypeAttack, strTypeDefend):
# TypeEffectiveness returns +1, 0, -1, or -2 in all cases.
# +1 signifies super effective. 0 signifies neutral.  -1 not very effective.  -2 doubly ineffective.
# Add the TypeEffectiveness against the defender's primary and secondary types.

    if strTypeAttack == "Bug":
        return TEBucket(strTypeDefend, "Dark Grass Psychic", "Fairy Fighting Fire Ghost Poison Steel")
    elif strTypeAttack == "Dark":
        return TEBucket(strTypeDefend, "Ghost Psychic", "Dark Fairy Fighting")
    elif strTypeAttack == "Dragon":
        return TEBucket(strTypeDefend, "Dragon", "Steel", "Fairy")
    elif strTypeAttack == "Electric":
        return TEBucket(strTypeDefend, "Dragon", "Steel", "Fairy")
    elif strTypeAttack == "Fairy":
        return TEBucket(strTypeDefend, "Dark Dragon Fighting", "Fire Poison Steel")
    elif strTypeAttack == "Fighting":
        return TEBucket(strTypeDefend, "Dark Ice Normal Rock Steel", "Bug Fairy Flying Poison Psychic", "Ghost")
    elif strTypeAttack == "Fire":
        return TEBucket(strTypeDefend, "Bug Grass Ice Steel", "Dragon Fire Rock Water")
    elif strTypeAttack == "Flying":
        return TEBucket(strTypeDefend, "Bug Fighting Grass", "Electric Rock Steel")
    elif strTypeAttack == "Ghost":
        return TEBucket(strTypeDefend, "Ghost Psychic", "Dark", "Normal")
    elif strTypeAttack == "Grass":
        return TEBucket(strTypeDefend, "Ground Rock Water", "Bug Dragon Fire Flying Grass Poison Steel")
    elif strTypeAttack == "Ground":
        return TEBucket(strTypeDefend, "Electric Fire Poison Rock Steel", "Bug Grass", "Flying")
    elif strTypeAttack == "Ice":
        return TEBucket(strTypeDefend, "Dragon Flying Grass Ground", "Fire Ice Steel Water")
    elif strTypeAttack == "Normal":
        return TEBucket(strTypeDefend, "", "Rock Steel", "Ghost")
    elif strTypeAttack == "Poison":
        return TEBucket(strTypeDefend, "Fairy Grass", "Ghost Ground Poison Rock", "Steel")
    elif strTypeAttack == "Psychic":
        return TEBucket(strTypeDefend, "Fighting Poison", "Psychic Steel", "Dark")
    elif strTypeAttack == "Rock":
        return TEBucket(strTypeDefend, "Bug Fire Flying Ice", "Fighting Ground Steel")
    elif strTypeAttack == "Steel":
        return TEBucket(strTypeDefend, "Fairy Ice Rock", "Electric Fire Steel Water")
    elif strTypeAttack == "Water":
        return TEBucket(strTypeDefend, "Fire Ground Rock", "Dragon Grass Water")
    return 0

def TEBucket(strTypeDefend, strPlus1, strNeg1, strNeg2=""):
    if strPlus1.find(strTypeDefend) >= 0:
        return 1
    elif strNeg1.find(strTypeDefend) >= 0:
        return -1
    elif strNeg2.find(strTypeDefend) >= 0:
        return -2
    return 0


