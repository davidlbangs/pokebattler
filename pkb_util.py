import math


def RoundUp(val):
    return math.ceil(val)


def RoundDown(val):
    return math.floor(val)


def Round(val, cDigits=0):
    return round(val, cDigits)


def MinMax(val1, valMin, valMax):
    if val1 > valMax:
        return valMax
    elif val1 < valMin:
        return valMin
    return val1


def WeightedAverage(stat1, stat2, weightStat1):
    # weight should be a value between zero and one, inclusive.
    # 0.5 makes it an average, 1 makes it all stat1, 0 makes it all stat2
    return stat1 * weightStat1 + stat2 * (1 - weightStat1)


def ParseSubstring(strInput, iSubstring, strSep=","):
    # Unlike the VBA version, this is 0 based.  iSubstring = 0 to get the first substring.

    if strInput == "":
        return ""

    strSubstring = ""
    ichSubstring = 0
    ichSeparator = strInput.find(strSep)

    while iSubstring > 0:
        if ichSeparator < 0:
            return ""  # not found

        ichSubstring = ichSubstring + ichSeparator + len(strSep)
        ichSeparator = strInput[ichSubstring:].find(strSep)
        iSubstring = iSubstring - 1

    if ichSeparator >= 0:
        strSubstring = strInput[ichSubstring: ichSubstring + ichSeparator]
    elif ichSubstring > 0:
        strSubstring = strInput[ichSubstring:]

    return strSubstring.strip()

