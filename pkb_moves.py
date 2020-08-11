from pkb_xldata import *

def TypeOfQuickMove(strMove, strDefaultType = ""):
    # return the Type entry for a quick move in the charge move table.
    # Because of use in Excel Spreadsheet, return "" for ""
    # If strMove is the name of a type, return that type.
    # As a last resort, return "Unknown"

    if strMove == "":
        return strDefaultType

    rowData = TupleFromTable(xlf.tableQuickMoves, strMove)

    if len(rowData) == 0:
        if SymbolForType(strMove) != "?":
            return strMove  # if the move is the name of a type, just return it.

        return "Unknown"
    else:
        return rowData[1]

class QuickMove:
    def __init__(self, strMove, strDefaultType=""):
        self.strMove = strMove
        self.strType = strDefaultType
        self.cTurnsToQuick = 0
        self.tupleData = ()
        self.valDamage = 0
        self.valEnergy = 0
        self.cTurnsToQuick = 0
        self.valDamagePerTurns = 0
        self.valEnergyPerTurn = 0

        if strMove != "":
            self.tupleData = TupleFromTable(xlf.tableQuickMoves, strMove)

            if len(self.tupleData) == 0:
                if SymbolForType(strMove) != "?":
                    self.strType = strMove
                else:
                    self.strType = "Unknown"
            else:
                self.strType = self.tupleData[1]
                self.valDamage = self.tupleData[2]
                self.valEnergy = self.tupleData[3]
                self.cTurnsToQuick = self.tupleData[4]
                self.valDamagePerTurn = self.tupleData[5]
                self.valEnergyPerTurn = self.tupleData[6]

    def RoundUpTurns(self, cTurns):
        return RoundUp(cTurns / self.cTurnsToQuick) * self.cTurnsToQuick

def TypeOfChargeMove(strMove, strDefaultType = ""):
    # return the Type entry for a charge move in the charge move table.
    # Because of use in Excel Spreadsheet, return "" for ""
    # If strMove is the name of a type, return that type.
    # As a last resort, return "Unknown"

    if strMove == "":
        return strDefaultType

    tupleData = TupleFromTable(xlf.tableChargeMoves, strMove)

    if len(tupleData) == 0:
        if SymbolForType(strMove) != "?":
            return strMove  # if the move is the name of a type, just return it.

        return "Unknown"
    else:
        return tupleData[1]

class ChargeMove:
    def __init__(self, strMove, qm, strDefaultType=""):
        self.tupleData = ()
        self.strMove = strMove
        self.strType = strDefaultType
        self.valPower = 0
        self.valEnergy = 0
        self.valChanceOfBuff = 0
        self.cTurnsToCharge = 0

        if strMove != "":
            self.tupleData = TupleFromTable(xlf.tableChargeMoves, strMove)

            if len(self.tupleData) == 0:
                if SymbolForType(strMove) != "?":
                    self.strType = strMove
                else:
                    self.strType = "Unknown"
            else:
                self.strType = self.tupleData[1]
                self.valPower = self.tupleData[2]
                self.valEnergy = self.tupleData[3]
                self.valChanceOfBuff = self.tupleData[5]

                if qm.valEnergyPerTurn != 0: # quick move might be type only, or invalid
                    self.cTurnsToCharge = qm.RoundUpTurns(self.valEnergy / qm.valEnergyPerTurn)


def ParseMoveName(csv, iMove):
    # unlike in VBA version, this is zero based.  Use 0 for the Quick Move
    
    strMoveName = ParseSubstring(csv, iMove + 1, ",").title()
    
    if strMoveName.find("_") >= 0:
        strMoveName = strMoveName.replace("_", " ")

    assert isinstance(strMoveName, str)
    return strMoveName

def SymbolForType(strType):
    # return a single unicode character which represents the type.
    # Because of use in Excel Spreadsheet, return "" for ""
    # if the type does not exist in the table, return "?"

    if strType == "":
        return ""

    for rowData in xlf.tableTypeSymbols:
        if rowData[0] == strType : return rowData[1]

    return "?"

