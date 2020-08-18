from pkb_moves import *

# This module implements the concepts and classes for Pokemon

pkData_Pokemon = 0
pkData_Type1 = 1
pkData_Type2 = 2
pkData_Category = 3
pkData_Attack = 4
pkData_Defense = 5
pkData_Stamina = 6
pkData_MaxCp = 7

def ParsePokemonName(csv, fDataName=False):
    # 'Pull the name of the Pokemon from PVPOKE description and parse out Pokemon name.
    #  'Correct case and put underscore qualifiers in parentheses for beauty.

    strPokemon = csv.title()

    ichComma = strPokemon.find(",")
    if ichComma >= 0:   strPokemon = strPokemon[:ichComma]

    ichUnderscore = strPokemon.find("_")
    if ichUnderscore >= 0:
        # Replace underscore with parenthesis.  Was necessary to Parse csv imported from pvpoke.com
        strParen = strPokemon[ichUnderscore + 1:].title()
        strPokemon = strPokemon[:ichUnderscore] + " (" + strParen + ")"

    if fDataName:
        if strPokemon.find("(Shadow)") >= 0:
            strPokemon = strPokemon.replace("(Shadow)", "")

    return strPokemon.strip()

#

class IndividualValues: # Known as the pokemon's iv's
    def __init__(self, csvIvs = ""):
        ' if no csvIvs is given, assume default values'
        self.attack = int(MinMax(ValFromCsv(csvIvs, 0, 13), 0, 15))
        self.defense = int(MinMax(ValFromCsv(csvIvs, 1, 13), 0, 15))
        self.stamina = int(MinMax(ValFromCsv(csvIvs, 2, 13), 0, 15))
        self.levelMax = int(MinMax(ValFromCsv(csvIvs, 3, 40), 1, 41) * 2) / 2 # must be multiple of 0.5 from 1 to 41
        self.csvIvs = csvIvs


class BattleStats:
    def __init__(self):
        self.cp = 0
        self.level = 0
        self.attCMP = 0      # attack stat before the Shadow boost was applied, used for determining CMP winner
        self.attInit = 0     # attack stat at start of battle
        self.attack  = 0        # attack stat as modified by stat changing attacks
        self.defInit = 0     # defense stat at start of battle.
        self.defense = 0         # defense stat as modified by stat changing attacks
        self.hp = 0          # stamina stat, must be a whole number

class Pokemon:

    def __init__(self, csv, csvIvs=""):
        
        self.strName = ParsePokemonName(csv)
        self.csv = self.strName # We will rebuild this to beautify and standardize

        self.strType1 = ""
        self.strType2 = ""

        self.tupleData = ()
        self.fInvalid = False
        self.fInvalidPokemon = False
        self.fInvalidQuickMove = False
        self.fInvalidChargeMove = False
        self.fMultipleChargeMoves = False
        self.fLegendaryOrMythical = False
        self.fShadow = False

        strQuickMove = ParseMoveName(csv, 0)
        if strQuickMove != "":  self.csv = self.csv + ", " + strQuickMove
        
        strChargeMove = ParseMoveName(csv, 1)
        if strChargeMove != "": self.csv = self.csv + ", " + strChargeMove
        
        iNextMove = 2
        strNextMove = ParseMoveName(csv, iNextMove)
        
        while strNextMove != "":
            strType = TypeOfChargeMove(strNextMove)
            if strType == "Unknown":
                self.fInvalid = True
                self.fInvalidChargeMove = True
            self.fMultipleChargeMoves = True
            self.csv = self.csv + ", " + strNextMove
            iNextMove = iNextMove + 1
            strNextMove = ParseMoveName(csv, iNextMove)
        
        # self.csv is now a normalized version of csv, with consistent spacing and capitalization.

        self.strNameData = self.strName
        if self.strName.find("(Shadow)") >= 0:
            self.strNameData = self.strName.replace(" (Shadow)","").strip()
            self.fShadow = True

        self.tupleData = TupleFromTable(xlf.tablePokemon, self.strNameData)

        if len(self.tupleData) == 0:
            self.fInvalid = True
            self.fInvalidPokemon = True
        else:
            self.strType1 = self.tupleData[pkData_Type1]
            self.strType2 = self.tupleData[pkData_Type2]

            self.qm = QuickMove(strQuickMove)

            if len(self.qm.tupleData) == 0:
                self.fInvalid = True
                self.fInvalidQuickMove = True
            else:
                self.cm = ChargeMove(strChargeMove, self.qm)

                if len(self.cm.tupleData) == 0:
                    self.fInvalid = True
                    self.fInvalidChargeMove = True

        self.ivs = IndividualValues(csvIvs)

        strCategory = StrFromTuple(self.tupleData, pkData_Category)
        if strCategory == "M" or strCategory == "L":
            self.fLegendaryOrMythical = True

        # fields filled in by QualifyPokemon
        self.fQualified = False
        self.fTypeEffectivenessBattle = False
        self.bstat = BattleStats()

        # end __init__








        
            
        
        

