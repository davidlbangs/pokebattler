import openpyxl
import typing
from pkb_util import *

class XLFile:
    def __init__(self, strXLFile):
        self.strXLFile = strXLFile

        wb = openpyxl.load_workbook(strXLFile, data_only=True, read_only=True)
        self.tablePokemon = ReadXLTable(wb, "_PokemonTable")
        self.tableQuickMoves = ReadXLTable(wb, "_QuickMoveTable")
        self.tableChargeMoves = ReadXLTable(wb, "_ChargeMoveTable")
        self.tableTypeSymbols = ReadXLTable(wb, "_TypeSymbolTable")
        self.tableStatMultiplier = ReadXLTable(wb, "_StatMultiplierTable")
        self.tableBattleLeagues = ReadXLTable(wb, "_BattleLeagueTable")
    pass

def ReadXLTable(wb, strTable):

    if '!' in strTable:
        # passed a worksheet!cell reference
        ws_name, reg = strTable.split('!')
        if ws_name.startswith("'") and ws_name.endswith("'"):
            # optionally strip single quotes around sheet name
            ws_name = ws_name[1:-1]
        region = wb[ws_name][reg]
    else:
        # passed a named range; find the cells in the workbook
        full_range = wb.defined_names[strTable]
        if full_range is None:
            raise ValueError(
                'Range "{}" not found in workbook "{}".'.format(strTable, xlsx_file)
            )
        # convert to list (openpyxl 2.3 returns a list but 2.4+ returns a generator)
        destinations = list(full_range.destinations)
        if len(destinations) > 1:
            raise ValueError(
                'Range "{}" in workbook "{}" contains more than one region.'
                .format(strTable, xlsx_file)
            )
        ws, reg = destinations[0]
        # convert to worksheet object (openpyxl 2.3 returns a worksheet object
        # but 2.4+ returns the name of a worksheet)
        if isinstance(ws, str):
            ws = wb[ws]
        region = ws[reg]

    table = []
    for row in region:
        rowData = []
        for cell in row:
            rowData.append(cell.value)
        table.append(rowData)

    return table

def TupleFromTable(table, strKey):
    for row in table:
        if row[0] == strKey:
            return tuple(row)

    return ()  # not found'

def ValFromTuple(tuple: tuple, iProp, valDefault=0):

    if len(tuple) <= iProp: return valDefault
    if tuple[iProp] is None: return valDefault
    return tuple[iProp]

def ValFromCsv(csv: str, iProp, valDefault=0):

    str = ParseSubstring(csv, iProp)
    if str != "" and str.isnumeric():
        return float(str)

    return valDefault

def StrFromTuple(tuple: tuple, iProp):
    if len(tuple) <= iProp: return ""
    if tuple[iProp] is None: return ""
    return tuple[iProp]

xlf = XLFile("Smart Battle Table.xlsm")