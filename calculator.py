import json
import pyexcel as pe
from pyexcel_ods3 import get_data
# from pyexcel.ext import ods3
from math import floor


def main():
    trackerFName = "pkgo_IV_tracker.ods"
    baseStatsFName = "pkgo_base_stats_table.ods"
    statsDict = {}
    lvlByDust = {}
    CPMByLvl = {}
    '''
    data = get_data(odsFName)
    print(json.dumps(data))
    for s in data:
        for line in s:
            print(line)
    '''
    statsBook = pe.get_book(file_name=baseStatsFName)
    bstatsSheet = statsBook[0]
    cpmSheet = statsBook[1]
    dustSheet = statsBook[2]

    for pkmnLine in bstatsSheet:
        if pkmnLine[0] != "PkMn num":
            statsDict[pkmnLine[1].lower()] = pkmnLine

    for pkmnLine in cpmSheet:
        CPMByLvl[pkmnLine[0]] = pkmnLine[1]

    for pkmnLine in dustSheet:
        dust = pkmnLine[0]
        lv = pkmnLine[1]
        if dust in lvlByDust:
            lvlByDust[dust].append(lv)
        else:
            lvlByDust[dust] = [lv]

    bestStatPosStrings = ['hp', 'stam', 'atk', 'def']
    tierList = [range(0, 22+1), range(23, 29+1), range(30, 36+1), range(37, 45+1), range(46)]
    bestIVList = [range(0, 7+1), range(8, 12+1), range(13, 14+1), [15], range(16)]
    tierTranslationBuilder = {0: ['not', 'headway', 'low'], 1: ['above', 'avg', 'average', 'mid'],
                              2: ['caught', 'attention', 'attn', 'high'], 3: ['wonder', 'top'], 4: ['test']}
    ivTranslationBuilder = {0: ['norm', 'not', 'low'], 1: ['pos', 'trending', 'mid'],
                            2: ['impressed', 'imp', 'must', 'high'], 3: ['exceed', 'exc', 'incredible', 'top'], 4: ['test']}

    tierTranslation = {}
    for t in tierTranslationBuilder:
        for k in tierTranslationBuilder[t]:
            tierTranslation[k] = t

    ivTranslation = {}
    for t in ivTranslationBuilder:
        for k in ivTranslationBuilder[t]:
            ivTranslation[k] = t
    qcon = {'hp': 'stam', 'stam': 'stam', 'atk': 'atk', 'def': 'def'}

    b = pe.get_book(file_name=trackerFName)
    # print(b)
    # for s in b:
    headerKeys = {}
    for pkmnLine in b[0]:
        if pkmnLine[0] != 'pkmn':
            pkmn = pkmnLine[headerKeys['pkmn']]
            print(pkmnLine)
            baseStam, baseAtk, baseDef = statsDict[pkmn.lower()][2:5]
            maxLvl = 35
            lvlList = []
            comboVals = []
            sumLine = []
            if pkmnLine[headerKeys['hatched']].lower() == 'x':
                maxLvl = 20
            for lv in lvlByDust[pkmnLine[headerKeys['dust']]]:
                if lv <= maxLvl:
                    if pkmnLine[headerKeys['powered']].lower() == 'x':
                        lvlList.append(lv)
                    elif lv%1 == 0:
                        lvlList.append(lv)
            for level in lvlList:
                cpm = CPMByLvl[level]
                for stamIV in range(16):
                    calcdHP = (baseStam + stamIV) * cpm
                    if floor(calcdHP) == pkmnLine[headerKeys['hp']]:
                        comboVals.append([level, stamIV])

            for level, stamIV in comboVals:
                cpm = CPMByLvl[level]
                for atkIV in range(16):
                    for defIV in range(16):
                        calcdCP = ((baseAtk + atkIV) * (baseDef + defIV)**0.5 * (baseStam + stamIV)**0.5 * (cpm)**2)/10
                        if floor(calcdCP) == pkmnLine[headerKeys['cp']]:
                            sumtier = tierTranslation[pkmnLine[headerKeys['tier']]]
                            bestStatListTmp = pkmnLine[headerKeys['bestStat']].lower().split(' ')
                            bestStatList = []
                            for bs in bestStatListTmp:
                                bestStatList.append(qcon[bs])
                            bestStatTier = ivTranslation[pkmnLine[headerKeys['bestVal']]]
                            ivsum = stamIV + atkIV + defIV
                            if ivsum in tierList[sumtier]:
                                inrange = True
                                for bs in bestStatList:
                                    if bs == "stam" or bs == "hp":
                                        if stamIV not in bestIVList[bestStatTier]:
                                            inrange = False
                                    elif bs == "atk":
                                        if atkIV not in bestIVList[bestStatTier]:
                                            inrange = False
                                    elif bs == "def":
                                        if defIV not in bestIVList[bestStatTier]:
                                            inrange = False
                                    else:
                                        print('oops, bad stat string entered')
                                        exit()
                                if inrange:
                                    ivDict = {'stam': stamIV, 'atk': atkIV, 'def': defIV}
                                    bsVal = ivDict[qcon[bestStatList[0]]]
                                    incl = True
                                    for k in ivDict:
                                        k = qcon[k]
                                        if k not in bestStatList:
                                            if ivDict[k] >= bsVal:
                                                incl = False
                                    if incl:
                                        print(pkmn, level, stamIV, atkIV, defIV)
                                        sumLine.append(ivsum)
            sumLine = sorted(sumLine)
            print(round((sumLine[0]/45.)*100., 2), sumLine, round((sumLine[-1]/45.)*100., 2))
            print()
        else:
            print(pkmnLine)
            i = 0
            for k in pkmnLine:
                headerKeys[k] = i
                i += 1


if __name__ == '__main__':
    main()
