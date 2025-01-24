#-*- coding: utf-8 -*-
import sys
import collections
import os
import csv
import json
import math
import xlsxwriter
import importlib
import re

from xlsxwriter.utility import xl_rowcol_to_cell
from xlsxwriter.utility import xl_range

#from PyPDF2 import PdfFileWriter, PdfFileReader, PdfFileMerger
from PyPDF2 import PdfWriter, PdfReader, PdfMerger

import io
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter

import shutil
import matplotlib.pyplot as plt
from matplotlib.ticker import MultipleLocator, FormatStrFormatter

from uuid import getnode as get_mac

g_cellFreqListStart = None
g_cellFreqListEnd = None

g_configDict = {'deviceName':'Not assigned',
                'aittFold':'C:\\Program Files\\Advanced Interconnect Test Tool (64-Bit)',
                'fittedReverse': 'false'}

g_dut_list = []
g_layer_list = []
g_length_list = []
g_sampleFreqList = []
#g_sampleFreqList = [1, 3, 4, 6, 8, 10 ,12.89, 16, 20, 25, 40]
#g_sampleFreqList = [4, 8, 12.89, 16, 28]

###################
# utility
def getDefaultFormat(workbook,colorHex):
    format_default = workbook.add_format({
        'bold': False,
        'border': 1,
        'align': 'center',
        'valign': 'vcenter',
        'fg_color': colorHex,
        'font':'calibri'
    })
    return format_default

def readCsvToDict(filePath):
    with open(filePath) as csvFile:
        csv_reader = csv.DictReader(csvFile, delimiter=',')
        for row in csv_reader:
            print(row)

def read_freq_report_table(filePath, db_l1l2):
    with open(filePath, newline='') as csvFile:
        csv_reader = csv.reader(csvFile,delimiter=',')
        rows = list(csv_reader)
        rowIdx = -1
        freq = 0
        for row in rows[1:]:
            rowIdx = rowIdx + 1
            if (rowIdx % 2) == 0:
                #freq = float(row[1].split(' ')[0])
                freq = row[1]
                db_l1l2.setdefault(freq, {})
            else:
                uncertainty = float(row[1])
                dB_in       = float(row[2])
                db_l1l2[freq]['Uncertainty'] = uncertainty
                db_l1l2[freq]['dB_in'] = dB_in
    return

def read_common(filePath, db_common):
    with open(filePath, newline='') as csvFile:
        csv_reader = csv.DictReader(csvFile, delimiter=',')
        for row in csv_reader:
            db_common.update(row)
            db_common.pop('', None)

def read_trace_report_table(filePath, db_common):
    read_common(filePath, db_common)
    return

def read_impedance_report_table(filePath, db_common):
    tmpDuct = {}
    read_common(filePath, tmpDuct)

    #"Z of Trace 1"
    pattern = r"Z of (Trace .+) "
    impedance_duct = {}
    for key, val in tmpDuct.items():
        match = re.match(pattern, key) 
        if match:
            trace = match.group(1)
            #print(f'trace:{trace}, val:{val}')
            impedance_duct[trace] = float(val)    
        pass
    return impedance_duct

def get_uncertainty_duct(csv_reader):
    #"Uncertainty @12.89GHz 3.083% - Frequency (GHz)"
    pattern = r"Uncertainty @(.+)GHz (.+)% - Frequency \(GHz\)" 

    uncertainty_duct = {}
    tmpRow = next(iter(csv_reader))
    for key in tmpRow.keys():
        #print(key)
        match = re.match(pattern, key) 
        if match: 
            frequency = match.group(1)
            uncertainty = match.group(2) 
            print(f"Frequency: {frequency:10}GHZ, Uncertainty: {uncertainty:10}%")
            uncertainty_duct[float(frequency)] = float(uncertainty)
        # else: 
        #     print("No match found.")
    
    for configSamplingFreq in g_sampleFreqList:
        if configSamplingFreq not in uncertainty_duct:
            print(f'[ERROR] uncertainty sanity check error!!, confg freq not found: {configSamplingFreq}', )

            len_duct = len(uncertainty_duct)
            print(f'===== DBGINFO =====, uncertainty_duct.size:{len_duct}')
            for freq, parcent in uncertainty_duct.items():
                print(f'[DBGINFO] uncertainty_duct[{freq}]: {parcent}')
            sys.exit(-1)
    
    return uncertainty_duct

def read_uncertainty_plotl1l2(filePath, db):
    key_list = ['Frequency (GHz)',
                'Fitted',
                'Insertion Loss']
    freq_list = []
    fitted_list = []
    iLoss_list = []
    uncertainty_duct = {}

    with open(filePath, newline='') as csvFile:
        csv_reader = csv.DictReader(csvFile, delimiter=',')
        uncertainty_duct = get_uncertainty_duct(csv_reader)

        for row in csv_reader:
            freq_list.append(float(row['Frequency (GHz)']))
            fitted_list.append(float(row['Fitted']))
            iLoss_list.append(float(row['Insertion Loss']))
            
        
    return freq_list, fitted_list, iLoss_list, uncertainty_duct

def genWorkBook(collectResFold):
    print("[INFO][SummaryReport] start {0}> ".format('='*30))
    collectResPath = os.path.join(collectResFold, 'SummaryReport.xlsx')
    workbook = xlsxwriter.Workbook(collectResPath)
    return workbook

def finishWorkBook(workbook):
    workbook.close()
    print("[INFO][SummaryReport] done <{0} ".format('='*30))

def readConfig(dataFold, collectResFold):
    def csv_read(file):
        with open(file, newline='') as f:
            row_list = f.read().splitlines()
            return row_list

    global g_configDict, g_dut_list, g_layer_list, g_length_list, g_sampleFreqList
    # g_dut_list = csv_read('dut.csv')
    # g_layer_list = csv_read('layer.csv')
    # g_length_list = csv_read('length.csv')
    # config_list = csv_read('config.csv')

    with open('dut_layer_length.json', 'r') as file:
        jdb = json.load(file)

    print('#'*30)
    print('# dump config list')
    for item in jdb.items():
        print("[INFO][config] {0} ".format(item))
    print('#'*30)
    print()

    # global g_configDict
    # for parameter in config_list:
    #     pair = parameter.split(',')
    #     attr, configedVal = pair[0], pair[1]
    #     if attr not in g_configDict.keys():
    #         print("[ERROR] config fail, attr:{0}".format(attr))
    #         print("[ERROR] config fail, keys:{0}".format(g_configDict.keys()))
    #         sys.exit(-2)
    #     g_configDict[attr] = configedVal
    #     pass

    # for item in g_configDict.items():
    #     print("[INFO][config] {0} ".format(item))

    g_configDict  = jdb['config']
    g_dut_list    = jdb['dut']
    g_layer_list  = jdb['layer']
    g_length_list = jdb['length']
    g_sampleFreqList = jdb['sampleFreq']

    format = g_configDict['format'] # format:{s2p, s4p}
    s4pDict = {}
    for dut in g_dut_list:
        s4pDict[dut] = {}
        #s4pDict[dut]['name'] = dut
        for layer in g_layer_list:
            for length in g_length_list:
                fileName = f'{dut}-{layer}-{length}.{format}'
                filePath = os.path.join(dataFold, f'input_{format}', fileName)
                if not os.path.isfile(filePath):
                    print("[ERROR] {0} doest not exist!!".format(filePath))
                    sys.exit(-1)
                layerDict = s4pDict[dut].setdefault(layer,collections.OrderedDict())
                layerDict.setdefault(length, filePath)
                layerDict['outFold'] = os.path.join(collectResFold, "output_{0}_{1}".format(dut,layer))
        pass

    ##############################
    ##### dump input file list####
    print('#'*30)
    print('# dump input file list')
    for dut, layerDict in s4pDict.items():
        for layer, lengthDict in layerDict.items():
            for length, inputFileName in lengthDict.items():
                print(f'[INFO][input] dut:{dut}, layer:{layer}, leng:{length}, inputFile:{inputFileName}')
    print('#'*30)
    print()
    #sys.exit(0)

    #########################
    # for aitt.exe parameter
    numOfLength = len(g_length_list)
    if(numOfLength != 2) and (numOfLength != 3):
        print("[ERROR] num of length is wrong: {0}".format(numOfLength))
        print("[ERROR] please check length.csv")
        sys.exit(-1)

    #aittFold = os.path.join("C:\Program Files\Advanced Interconnect Test Tool (64-Bit)")
    aittFold = g_configDict['aittFold']
    filePath_script = \
        os.path.join(aittFold, 'script_examples','deltal_{0}l_report.js'.format(numOfLength))
    filePath_script = "\"{0}\"".format(filePath_script)

    ############################
    ##### dump aitt command ####
    isDump = False
    if isDump:
        print('#'*30)
        print('# dump aitt command')    
        pass
    for dut in g_dut_list:
        for layer in g_layer_list:
            aittCmd = 'aitt.exe -s {script}'.format(script=filePath_script)
            layerDict = s4pDict[dut][layer]
            for length in g_length_list:
                fileName = layerDict[length]
                filePath = os.path.join(dataFold,fileName)
                opt = 'true'
                length = int(length.split('IN')[0])
                aittCmd += " {0} {1} {2}".format(filePath, opt, length)
            aittCmd = "{0} {1}".format(aittCmd, os.path.join(dataFold, layerDict['outFold']))
            if isDump:
                print (aittCmd)
                print()
                pass
            layerDict['aittCmd'] = aittCmd
    if isDump:
        print('#'*30)
        print()

    return s4pDict

def binarySearchPrevNext(target, inList):
    maxIdx = len(inList) - 1
    minIdx = 0
    if target > inList[maxIdx]:
        print("[ERROR] out of range, target:{0}, freq range({1}, {2})".format(
            target, inList[minIdx], inList[maxIdx]))
    #    return maxIdx, maxIdx
    elif target < inList[minIdx]:
        print("[ERROR] out of range, target:{0}, freq range({1}, {2})".format(
            target,inList[minIdx], inList[maxIdx]))
    #    return minIdx, minIdx
    else:
        # inList[0] < value < inList[-1]
        pass

    start = 0
    end = maxIdx
    while start<=end:
        mid = start + (end-start)//2
        midVal = inList[mid]
        if target == midVal:
            break
        elif target > midVal:
            start = mid + 1
        else:
            end = mid - 1
    return (start - 1), start


####################
# [pre process] mag png
def runFreqMagEach(filePath, length):
    def dumpFreqMagInfo_s4p(attrDict):
    # ['!', 'FREQ.GHZ', 'S11RE', 'S11IM', 'S12RE', 'S12IM', 'S13RE', 'S13IM', 'S14RE', 'S14IM']
    # ['!', 'S21RE', 'S21IM', 'S22RE', 'S22IM', 'S23RE', 'S23IM', 'S24RE', 'S24IM']
    # ['!', 'S31RE', 'S31IM', 'S32RE', 'S32IM', 'S33RE', 'S33IM', 'S34RE', 'S34IM']
    # ['!', 'S41RE', 'S41IM', 'S42RE', 'S42IM', 'S43RE', 'S43IM', 'S44RE', 'S44IM']
        print("<===== Freq Ghz[{0}] ========>".format(attrDict['FREQ.GHZ']))
        print("S31RE:{0}".format(attrDict['S31RE']))
        print("S31IM:{0}".format(attrDict['S31IM']))

        print("S32RE:{0}".format(attrDict['S32RE']))
        print("S32IM:{0}".format(attrDict['S32IM']))

        print("S41RE:{0}".format(attrDict['S41RE']))
        print("S41IM:{0}".format(attrDict['S41IM']))

        print("S42RE:{0}".format(attrDict['S42RE']))
        print("S42IM:{0}".format(attrDict['S42IM']))
        return

    def dumpFreqMagInfo_s2p(attrDict):
        print("<===== Freq Ghz[{0}] ========>".format(attrDict['FREQ.GHZ']))
        print("S11RE:{0}".format(attrDict['S11RE']))
        print("S11IM:{0}".format(attrDict['S11IM']))

        print("S21RE:{0}".format(attrDict['S21RE']))
        print("S21IM:{0}".format(attrDict['S21IM']))

        print("S12RE:{0}".format(attrDict['S12RE']))
        print("S12IM:{0}".format(attrDict['S12IM']))

        print("S22RE:{0}".format(attrDict['S22RE']))
        print("S22IM:{0}".format(attrDict['S22IM']))
        return

    def dumpFreqMagInfo(format, attrDict):
        if format =='s2p':
            dumpFreqMagInfo_s2p(attrDict)
        elif format =='s4p':
            dumpFreqMagInfo_s4p(attrDict)
        return

    def calculateMag(format, attrDict):
        if format == 's2p':
            reVal = (attrDict['S21RE'] )
            imVal = (attrDict['S21IM'] )
            tmpVal = math.sqrt(pow(reVal, 2) + pow(imVal, 2))
            val = 20 * math.log(tmpVal, 10)
            freq = attrDict['FREQ.GHZ']
            return (val, freq)
        elif format == 's4p':
            reVal = (attrDict['S31RE'] - attrDict['S41RE'] - attrDict['S32RE'] + attrDict['S42RE'])/2
            imVal = (attrDict['S31IM'] - attrDict['S41IM'] - attrDict['S32IM'] + attrDict['S42IM'])/2
            tmpVal = math.sqrt(pow(reVal, 2) + pow(imVal, 2))
            val = 20 * math.log(tmpVal, 10)
            freq = attrDict['FREQ.GHZ']
            return (val, freq)
        else:
            print(f'[ERROR] input wrong 2, format={format}')
            sys.exit(-1)


    ############################
    # config offset, fields
    def getOffsets(format):
        s2p_offset = 10
        s4p_offset = 13
        s2p_rowUnit = 1
        s4p_rowUnit = 4
        rowIdxStart = 8
        if format == 's2p':
            return (s2p_offset, s2p_rowUnit, rowIdxStart, rowIdxStart + s2p_rowUnit)
        elif format == 's4p':
            return (s4p_offset, s4p_rowUnit, rowIdxStart, rowIdxStart + s4p_rowUnit)
        else:
            print(f'[ERROR] input wrong, format={format}')
            sys.exit(-1)
        return(-1, -1)

    #filePath = 'AD001_L01_05IN.s4p'
    #print(f'runFreqMagEach(), filePath:{filePath}')
    with open(filePath, newline='', encoding="utf-8") as f:
        row_list = f.read().splitlines()

        global g_configDict
        format = g_configDict['format'] # format:{s2p, s4p}
        fielidOffset, rowUnit, rowIdxStart, rowIdxEnd = getOffsets(format)
        #print(f'format:{format}, fielidOffset:{fielidOffset}, rowUnit:{rowUnit}')
        #sys.exit(0)

        attrIdx = 0
        attrIdxMap = {}
        for rowIdx in range(rowIdxStart, rowIdxEnd):
            for attr in row_list[rowIdx].split()[1:]:
                attrIdxMap[attrIdx] = attr
                attrIdx += 1

        #print(attrIdxMap)
        #sys.exit(0)

        row_list = row_list[fielidOffset:]
        if (len(row_list)%rowUnit)!=0:
            print("[ERROR] input wrong, len(row_list)={0}".format(len(row_list)))
            sys.exit(-1)

        numOfDataGroup = len(row_list)//rowUnit
        freqInfo = [[],[]]
        for groupIdx in range(numOfDataGroup):
            rowIdxStart = groupIdx * rowUnit
            rowIdxEnd = rowIdxStart + rowUnit
            attrIdx = 0
            attrDict = collections.OrderedDict() 
            #print("rowStart:{0}, rowEnd:{1}".format(rowIdxStart, rowIdxEnd))
            for rowIdx in range(rowIdxStart, rowIdxEnd):
                for attr in row_list[rowIdx].split():
                    attrName = attrIdxMap[attrIdx]
                    attrDict[attrName] = float(attr)
                    attrIdx += 1
                #print(row_list[rowIdx])
                pass
            #dumpFreqMagInfo(format, attrDict)
            #pass
            val, freq = calculateMag(format, attrDict)
            freqInfo[0].append(freq)
            freqInfo[1].append(val)
            pass

        x = freqInfo[0]
        y = freqInfo[1]
        #plt.cla()
        #plt.grid(color='r', linestyle='--', linewidth=1, alpha=0.3)
        #plt.grid(b=True)
        plt.plot(x, y, label=length)
        plt.legend()

        ax = plt.gca()
        ax.yaxis.set_major_locator(MultipleLocator(2))
        ax.yaxis.set_minor_locator(MultipleLocator(1))

        plt.xlabel("Frequency (GHz)")
        plt.ylabel("Magnitude (dB)")
        plt.title("Input")
        #plt.show()
        #plt.savefig(outputFilePath, dpi=300, format="png")
    return

def runFreqMag(s4pDict,dataFold):
    #print(f'runFreqMag(), dataFold:{dataFold}')
    for dut in g_dut_list:
        for layer in g_layer_list:
            print("[INFO] parsing {0}_{1}".format(dut, layer))
            outFold = s4pDict[dut][layer]['outFold']
            outFilePath = os.path.join(outFold, 'magnitude.png')
            #print(f'outFilePath:{outFilePath}')

            plt.cla()
            for leng in g_length_list:
                fileName = s4pDict[dut][layer][leng]
                filePath = os.path.join(dataFold, fileName)
                #print(f'leng:{leng}, fileName:{fileName}, filePath:{filePath}')
                runFreqMagEach(filePath,leng)

            plt.savefig(outFilePath, dpi=300, format="png")

    print("[INFO][Magnitude] successfully generate\n")
    return
def make_database(workbook,dataSheet,dut,dutDict):
    for layer, layerDict in dutDict.items():
        outFold = layerDict['outFold']
        #print(outFold)

        csv1 = '{0}/freq_report_table.csv'.format(outFold)
        csv2 = '{0}/impedance_report_table.csv'.format(outFold)
        csv3 = '{0}/trace_report_table.csv'.format(outFold)
        #csv4 = '{0}/uncertainty_plot_l1l2.csv'.format(outFold)
        csv4 = os.path.join(outFold,'uncertainty_plot_l1l2.csv')


        db_l1l2 = {}
        db_common = {}
        db_uPlot = {}

        #read_freq_report_table(csv1, db_l1l2)
        impedance_duct = read_impedance_report_table(csv2, db_common)
        #read_trace_report_table(csv3, db_common)
        freq_list, fitted_list, insertLoss_list, uncertainty_duct = read_uncertainty_plotl1l2(csv4, db_uPlot)

        print(impedance_duct)
        #print(freq_list)
        #dutDict['name'] = dut
        global g_configDict
        reverseFlag = -1 if g_configDict['fittedReverse'] else 1
        layerDict['freqList'] = freq_list
        layerDict['fittedList'] = [float(i)*reverseFlag for i in fitted_list]
        layerDict['iLossList'] = [float(i) for i in insertLoss_list]
        layerDict['uncertaintyDuct'] = uncertainty_duct
        layerDict['impedance'] = impedance_duct
    print("[INFO][make_database] done!!, {dut}".format(dut=dut))
    return


####################
# [SummaryReport]
def run_summary_sheet(workbook, summarySheet, s4pDict):
    print("[INFO][SummaryReport] running SummarySheet")

    global g_sampleFreqList
    global g_dut_list, g_layer_list, g_length_list
    tmpDut = next(iter(s4pDict.values()))
    tmpLayer = next(iter(tmpDut.values()))

    format_bold = workbook.add_format({'bold': True})
    format_default = workbook.add_format({
        'bold': False,
        'border': 1,
        'align': 'center',
        'valign': 'vcenter',
        #'fg_color': '#bedcf2',
        'font':'calibri'
    })

    ########################
    # Summary - header
    summarySheet.set_column(0, 0, 1)
    summarySheet.merge_range('B1:F1', 'Impedance design:  Line width/ Line space')
    summarySheet.merge_range('B2:C2', 'Ohm', getDefaultFormat(workbook,'#bedcf2'))
    summarySheet.merge_range('B3:C3', '',format_default)

    offset_row = 1
    offset_col = 3
    idx = 0
    for layer in g_layer_list:
        summarySheet.write(offset_row, offset_col+idx, layer,getDefaultFormat(workbook,'#bedcf2'))
        summarySheet.write(offset_row+1, offset_col + idx, '',format_default)
        idx += 1
    
    ########################
    # Summary - Summary 
    offset_row_summary = 4
    offset_col_summary = 1

    summarySheet.write(offset_row_summary, offset_col_summary, 'Summary', format_bold)
    summarySheet.merge_range(offset_row_summary + 1, offset_col_summary,
                       offset_row_summary + 3, offset_col_summary+1,
                       'Panel', getDefaultFormat(workbook,'#bedcf2'))


    layerSize = len(g_layer_list)
    lengthSize = len(g_length_list)
    numOfDataPerLayer = 3 #(measured, Fitted, Uncertainty)/Per layer
    table0Size = layerSize * lengthSize

    ###########################
    # Summary - Table0 - field
    print(f'===== Generate Table 0 start =====>>>')
    offset_row_table0 = offset_row_summary +1
    offset_col_table0 = offset_col_summary +2

    summarySheet.merge_range(offset_row_table0, offset_col_table0, 
                             offset_row_table0, offset_col_table0 + table0Size -1,
                             "impedance", getDefaultFormat(workbook, '#FF69B4'))
    for layerIdx, layerStr in enumerate(g_layer_list):
        offset_row = offset_row_table0
        offset_col = offset_col_table0 + layerIdx * lengthSize

        summarySheet.merge_range(offset_row+2, offset_col, 
                                 offset_row+2, offset_col + 1,
                                 layerStr, getDefaultFormat(workbook, '#c3e4bc'))

        # for lengIdx, lengthStr in enumerate(g_length_list):
        #     summarySheet.write(offset_row+1, offset_col + lengIdx,
        #                        f'{lengthStr} inch', getDefaultFormat(workbook, '#c3e4bc'))
        #     pass
        pass

    ###########################
    # Summary - Table0 - body
    dutIdx = 0
    for dut, ductDict in s4pDict.items():
        for layerIdx, layerStr in enumerate(g_layer_list):
            offset_row = offset_row_table0 + 3 + dutIdx
            offset_col = offset_col_table0 + layerIdx * lengthSize

            layerDict = ductDict[layerStr]
            impedance_duct = layerDict['impedance']
            #print(impedance_duct)
            traceIdx = 0
            for trace, val in impedance_duct.items():
                summarySheet.write(offset_row, offset_col + traceIdx,
                               val, getDefaultFormat(workbook, '#FFB6C1'))

                summarySheet.write(offset_row_table0 + 1, offset_col + traceIdx,
                               trace, getDefaultFormat(workbook, '#c3e4bc'))

                traceIdx += 1
                pass
                    
            pass
        
        dutIdx += 1
        pass

    def run_statistic_table0(offset_row_input, offset_col_input, attr, offset_row_summary, numOfDut):
        summarySheet.write(offset_row_input, offset_col_summary+1, attr, getDefaultFormat(workbook,'#bedcf2'))
        summarySheet.write(offset_row_input, offset_col_summary, '', getDefaultFormat(workbook,'#E77D40'))

        for idx in range(numOfDataPerLayer):    
            for layerIdx, layerStr in enumerate(g_layer_list):
                for traceIdx in range(lengthSize):
                    offset_row = offset_row_input
                    offset_col = offset_col_input + layerIdx * lengthSize + traceIdx

                    if attr == 'Range':
                        max_cell = xl_rowcol_to_cell(offset_row-2, offset_col)
                        min_cell = xl_rowcol_to_cell(offset_row-1, offset_col)

                        val = "=({max_cell} - {min_cell})".format(
                            max_cell=max_cell, min_cell=min_cell)
                        summarySheet.write(offset_row, offset_col,
                                        val, getDefaultFormat(workbook, '#FFB6C1'))

                    else:
                        offset_dut_start = offset_row_summary + 4
                        offset_dut_end = offset_dut_start + (numOfDut - 1)
                        dutRange = xl_range(offset_dut_start, offset_col,
                                            offset_dut_end, offset_col)
                        val = "={func}({range})".format(func=attr, range=dutRange)

                        colorHex = '#e8f426' if attr =='Average' else '#FFB6C1'
                        summarySheet.write(offset_row, offset_col,
                                        val, getDefaultFormat(workbook, colorHex))
                    pass
                pass
            pass

    numOfDut = len(s4pDict)
    offset_row_statistic = offset_row_summary + 4 + numOfDut
    offset_col_statistic = offset_col_summary + 2
    run_statistic_table0(offset_row_statistic+0, offset_col_statistic, 'Average', offset_row_summary, numOfDut)
    run_statistic_table0(offset_row_statistic+1, offset_col_statistic, 'Max'    , offset_row_summary, numOfDut)
    run_statistic_table0(offset_row_statistic+2, offset_col_statistic, 'Min'    , offset_row_summary, numOfDut)
    run_statistic_table0(offset_row_statistic+3, offset_col_statistic, 'Range'  , offset_row_summary, numOfDut)

    print(f'===== Generate Table 0 done <<<=====')

    ########################
    # Summary - Table1
    print('\n')
    print(f'===== Generate Table 1 start =====>>>')
    offset_row_summary_sample = offset_row_summary + 1
    offset_col_summary_sample = offset_col_summary + 2
    offset_col_summary_sample += table0Size
    sampleIdx = 0
    for sampleFreq in g_sampleFreqList:
        sampleStr = "(dB/in,@{0}GHz)".format(sampleFreq)
        lengOfSampleStr = len(sampleStr)
        offset_col = offset_col_summary_sample + (sampleIdx * (layerSize * numOfDataPerLayer))
        
        if layerSize >= 1:
            summarySheet.merge_range(offset_row_summary_sample, offset_col,
                                     offset_row_summary_sample, offset_col + (layerSize * numOfDataPerLayer -1),
                                     sampleStr, getDefaultFormat(workbook, '#c3e4bc'))
            summarySheet.set_column(offset_col,offset_col + (layerSize *numOfDataPerLayer-1), lengOfSampleStr)
        else:
            summarySheet.write(offset_row_summary_sample, offset_col,
                               sampleStr, getDefaultFormat(workbook, '#c3e4bc'))
            summarySheet.set_column(offset_col,offset_col,lengOfSampleStr)


        for layerIdx in range(layerSize):
            col_field = offset_col+ layerIdx*numOfDataPerLayer
            summarySheet.write(offset_row_summary_sample +1, col_field +0, 'measured'      , getDefaultFormat(workbook, '#c3e4bc'))
            summarySheet.write(offset_row_summary_sample +1, col_field +1, 'Fitted'        , getDefaultFormat(workbook, '#c3e4bc'))
            summarySheet.write(offset_row_summary_sample +1, col_field +2, 'Uncertainty(%)', getDefaultFormat(workbook, '#c3e4bc'))

            offset_row = offset_row_summary_sample + 2
            offset_col = offset_col_summary_sample + (sampleIdx * layerSize * numOfDataPerLayer)
            summarySheet.merge_range(offset_row, offset_col + layerIdx * numOfDataPerLayer,
                                     offset_row, offset_col + layerIdx * numOfDataPerLayer + (numOfDataPerLayer - 1),
                               g_layer_list[layerIdx], getDefaultFormat(workbook, '#c3e4bc'))

        sampleIdx += 1

    offset_row_data = offset_row_summary + 4
    offset_col_data = offset_col_summary_sample
    dutIdx = 0

    for dut, ductDict in s4pDict.items():
        offset_row = offset_row_data + dutIdx
        offset_col = offset_col_data
        
        summarySheet.write(offset_row, offset_col_summary, '', getDefaultFormat(workbook,'#E77D40'))
        summarySheet.write(offset_row, offset_col_summary+1, dut, getDefaultFormat(workbook,'#bedcf2'))
        for sampleIdx in range(len(g_sampleFreqList)):
            sampleFreq = g_sampleFreqList[sampleIdx]
            for layerIdx in range(layerSize):
                layerDict = ductDict[g_layer_list[layerIdx]]
                freq_list = layerDict['freqList']
                iLoss_list = layerDict['iLossList']
                fitted_list = layerDict['fittedList']
                uncertainty_duct = layerDict['uncertaintyDuct']

                # Interpolate
                prevIdx, nextIdx = binarySearchPrevNext(sampleFreq, freq_list)
                deltaFreq = (freq_list[nextIdx] - freq_list[prevIdx])
                deltaSampleFreq = (sampleFreq - freq_list[prevIdx])
                ratio = deltaSampleFreq / deltaFreq
                valAfterInterpoplate_iLoss  = iLoss_list[prevIdx]  + (ratio * (iLoss_list[nextIdx]  - iLoss_list[prevIdx]))
                valAfterInterpoplate_fitted = fitted_list[prevIdx] + (ratio * (fitted_list[nextIdx] - fitted_list[prevIdx]))
                '''
                print("freq:{0},{1}, iLoss:{2},{3}, sample:{4}, ratio:{5}, deltaFreq:{6}".format(
                    freq_list[prevIdx], freq_list[nextIdx],
                    iLoss_list[prevIdx],iLoss_list[nextIdx],
                    sampleFreq, ratio, deltaFreq))
                '''

                offset_col = offset_col_data + (sampleIdx * layerSize * numOfDataPerLayer)
                offset_col_data_measured    = offset_col + layerIdx * numOfDataPerLayer + 0
                offset_col_data_fitted      = offset_col + layerIdx * numOfDataPerLayer + 1
                offset_col_data_uncertainty = offset_col + layerIdx * numOfDataPerLayer + 2

                summarySheet.write(offset_row, offset_col_data_measured,
                                   valAfterInterpoplate_iLoss , getDefaultFormat(workbook, '#fbf5dc'))
                
                summarySheet.write(offset_row, offset_col_data_fitted,
                                   valAfterInterpoplate_fitted , getDefaultFormat(workbook, '#fbf5dc'))
                
                summarySheet.write(offset_row, offset_col_data_uncertainty,
                                   uncertainty_duct[sampleFreq] , getDefaultFormat(workbook, '#fbf5dc'))

        dutIdx += 1

    #statistic
    #numOfDataPerLayer = 3 #(measured, Fitted, Uncertainty)/Per layer
    def run_statistic(offset_row_input, offset_col_input, attr, offset_row_summary, numOfDut):
        #summarySheet.write(offset_row_input, offset_col_summary+1, attr, getDefaultFormat(workbook,'#bedcf2'))

        for idx in range(numOfDataPerLayer):
            #offset_col_input = offset_col_input + idx
                
            for sampleIdx in range(len(g_sampleFreqList)):
                for layerIdx in range(layerSize):
                    offset_row = offset_row_input
                    offset_col = offset_col_input + (sampleIdx * layerSize*numOfDataPerLayer) + layerIdx*numOfDataPerLayer +idx

                    if attr == 'Range':
                        max_cell = xl_rowcol_to_cell(offset_row-2, offset_col)
                        min_cell = xl_rowcol_to_cell(offset_row-1, offset_col)

                        val = "=({max_cell} - {min_cell})".format(
                            max_cell=max_cell, min_cell=min_cell)
                        summarySheet.write(offset_row, offset_col,
                                        val, getDefaultFormat(workbook, '#fbf5dc'))

                    else:
                        offset_dut_start = offset_row_summary + 4
                        offset_dut_end = offset_dut_start + (numOfDut - 1)
                        dutRange = xl_range(offset_dut_start, offset_col,
                                            offset_dut_end, offset_col)
                        val = "={func}({range})".format(func=attr, range=dutRange)

                        colorHex = '#e8f426' if attr =='Average' else '#fbf5dc'
                        summarySheet.write(offset_row, offset_col,
                                        val, getDefaultFormat(workbook, colorHex))
                    pass
                pass
            pass

    numOfDut = len(s4pDict)
    offset_row_statistic = offset_row_summary + 4 + numOfDut
    offset_col_statistic = offset_col_summary_sample
    run_statistic(offset_row_statistic+0, offset_col_statistic, 'Average', offset_row_summary, numOfDut)
    run_statistic(offset_row_statistic+1, offset_col_statistic, 'Max'    , offset_row_summary, numOfDut)
    run_statistic(offset_row_statistic+2, offset_col_statistic, 'Min'    , offset_row_summary, numOfDut)
    run_statistic(offset_row_statistic+3, offset_col_statistic, 'Range'  , offset_row_summary, numOfDut)

    print(f'===== Generate Table 1 done <<<=====')
    return

def run_data_sheet(workbook, dataSheet, s4pDict):
    print("[INFO][SummaryReport] running DataSheet")

    global g_dut_list, g_layer_list, g_length_list, g_sampleFreqList
    tmpDut = next(iter(s4pDict.values()))
    tmpLayer = next(iter(tmpDut.values()))

    lengthStr = "-".join(g_length_list)

    # common part
    font_calibri = workbook.add_format({'font':'calibri',
                                        'border': 1})
    bold = workbook.add_format({'bold': True,
                                'font':'calibri'})
    format_freq = workbook.add_format({
        'bold': False,
        'border': 1,
        'align': 'center',
        'valign': 'vcenter',
        'fg_color': '#bcd7e4',
        'font':'calibri'
    })

    format_freqAttr = workbook.add_format({
        'bold': False,
        'border': 1,
        'align': 'center',
        'valign': 'vcenter',
        'fg_color': '#c3e4bc',
        'font':'calibri'
    })

    format_DutLayerAttr = workbook.add_format({
        'bold': False,
        'border': 1,
        'align': 'center',
        'valign': 'vcenter',
        'fg_color': '#bedcf2',
        'font':'calibri'
    })

    dataSheet.merge_range('A1:B1', 'Raw data', bold)
    dataSheet.write('C1', lengthStr, bold)

    dataSheet.write('A2', g_configDict['deviceName'],format_freqAttr)
    dataSheet.merge_range('B2:B3', 'Freq.(GHz)',format_freqAttr)
    dataSheet.write('A3', 'Freq.(MHz)',format_freqAttr)
    dataSheet.set_column('A:B',len(g_configDict['deviceName']))

    # frequence
    offset_row_freqMhz = 3
    offset_col_freqMhz = 0
    offset_row_freqGhz = offset_row_freqMhz
    offset_col_freqGhz = offset_col_freqMhz + 1

    freq_list = tmpLayer['freqList']
    numOfFreq = len(freq_list)
    for freqIdx in range(numOfFreq):
        freq = freq_list[freqIdx]
        dataSheet.write(offset_row_freqMhz + freqIdx, offset_col_freqMhz, freq*1000, format_freq)
        dataSheet.write(offset_row_freqGhz + freqIdx, offset_col_freqGhz, freq, format_freq)

    global g_cellFreqListStart, g_cellFreqListEnd
    g_cellFreqListStart = xl_rowcol_to_cell(offset_row_freqGhz, offset_col_freqGhz)
    g_cellFreqListEnd   = xl_rowcol_to_cell(offset_row_freqGhz + (numOfFreq-1), offset_col_freqGhz)

    #data
    offset_row_attr_dutName = 1
    offset_col_attr_dutName = 2
    dutIdx = 0
    #layerNameList = ['L3','L5','L12','L14']
    layerNameList = g_layer_list
    layerSize = len(g_layer_list)

    for dut, dutDict in s4pDict.items():
        offset_dut = layerSize * dutIdx
        if layerSize > 1:
            dataSheet.merge_range(offset_row_attr_dutName, offset_col_attr_dutName + offset_dut,
                                  offset_row_attr_dutName, offset_col_attr_dutName + offset_dut + layerSize-1,
                                  dut, format_DutLayerAttr)
        else:
            dataSheet.write(offset_row_attr_dutName, offset_col_attr_dutName + offset_dut,
                              dut, format_DutLayerAttr)

        offset_row_freq = offset_row_attr_dutName + 2
        offset_col_freq = 2
        for layerIdx in range(layerSize):
            layer = g_layer_list[layerIdx]
            layerDict = dutDict[layer]
            # write layer name
            dataSheet.write(offset_row_attr_dutName + 1,
                            offset_col_attr_dutName + offset_dut + layerIdx,
                            g_layer_list[layerIdx], format_DutLayerAttr)

            # write frequence for layer
            for freqIdx in range(len(freq_list)):
                offset_row = offset_row_freq + freqIdx
                offset_col = offset_col_freq + offset_dut + layerIdx
                cell = xl_rowcol_to_cell(offset_row, offset_col)
                layerDict.setdefault('cellList', []).append(cell)
                dataSheet.write(offset_row,
                                offset_col,
                                layerDict['fittedList'][freqIdx],font_calibri)

        dutIdx = dutIdx + 1

    return

def runPictureSheet(workbook, picSheet, s4pDict):
    row_jump = 18
    col_jump = 7
    idx = 0
    for dut in g_dut_list:
        for layer in g_layer_list:
            outFold = s4pDict[dut][layer]['outFold']
            filePath_magnitude = os.path.join(outFold, 'magnitude.png')
            offset_row = 1 + (row_jump * idx)
            offset_col = 1
            picSheet.insert_image(offset_row, offset_col, filePath_magnitude, {'x_scale': 0.65, 'y_scale': 0.65})

            offset_row = offset_row
            offset_col = offset_col + col_jump
            filePath_uncertainty = os.path.join(outFold,'uncertainty_plot_l1l2.png')
            picSheet.insert_image(offset_row, offset_col, filePath_uncertainty, {'x_scale': 0.5, 'y_scale': 0.5})

            idx += 1
            #print("filePath_magnitude:{0}".format(filePath_magnitude))
            #print("filePath_uncertainty:{0}".format(filePath_uncertainty))
    print("[INFO][SummaryReport] running PictureSheet")
    return

def runLayerSheet(workbook,s4pDict):
    layer_list =['L01']
    for layer in g_layer_list:
        layerSheet = workbook.add_worksheet(layer)

        # Create a new chart object. In this case an embedded chart.
        chart1 = workbook.add_chart({'type': 'line'})
        chart1.set_size({'width':750,
                         'height':400})

        # Add the worksheet data that the charts will refer to.
        for dut, dutDict in s4pDict.items():
            layerDict = dutDict[layer]
            cellStart = layerDict['cellList'][0]
            cellEnd = layerDict['cellList'][-1]
            #print("{0}:{1},{2}".format(layer, cellStart, cellEnd))

            # Configure the first series.
            #global g_cellFreqListStart, g_cellFreqListEnd
            #print("g_cellFreqListStart:{0}".format(g_cellFreqListStart))
            #print("g_cellFreqListEnd:{0}".format(g_cellFreqListEnd))

            chart1.add_series({
                'name': "{0}_{1}".format(g_configDict['deviceName'],dut),
                'categories': '=Data!{cellStart}:{cellEnd}'.format(cellStart = g_cellFreqListStart,
                                                                   cellEnd = g_cellFreqListEnd),
                'values': '=Data!{cellStart}:{cellEnd}'.format(cellStart = cellStart,
                                                               cellEnd = cellEnd),
                'line': {'width': 2}
            })

        # Add a chart title and some axis labels.
        chart1.set_title({'name': g_configDict['deviceName']})
        chart1.set_x_axis({'name': 'Frequency (GHz)',
                           'label_position': 'low'})
        chart1.set_y_axis({'name': 'Insertion Loss (dB/in)'})

        # Set an Excel chart style. Colors with white outline and shadow.
        chart1.set_style(10)

        # Insert the chart into the worksheet (with an offset).
        layerSheet.insert_chart('C2', chart1, {'x_offset': 25, 'y_offset': 10})
        print("[INFO][SummaryReport] running {0}_Sheet".format(layer))

    return


####################
# [PDF] report
def mergePdf(inputFold, pdfName):
    packet = io.BytesIO()
    # create a new PDF with Reportlab
    can = canvas.Canvas(packet, pagesize=letter)
    can.drawString(50, 750, inputFold)
    can.save()

    # move to the beginning of the StringIO buffer
    packet.seek(0)
    existing_pdf = PdfReader(packet)
    # read your existing PDF
    filePath = os.path.join(os.getcwd(), inputFold, pdfName)
    new_pdf = PdfReader(open(filePath, "rb"))

    output = PdfWriter()
    # add the "watermark" (which is the new pdf) on the existing page
    page = existing_pdf.pages[0]
    page.merge_page(new_pdf.pages[0])
    output.add_page(page)
    # finally, write "output" to a real file
    newPdfName = "tmp_deltal_2l_report.pdf"
    filePath = os.path.join(os.getcwd(), inputFold, newPdfName)
    outputStream = open(filePath, "wb")
    output.write(outputStream)
    outputStream.close()
    return filePath

def generatePdfReport(s4pDict, collectResFold):
    print("[INFO][SummaryReportPdf] running...")
    mergedObject = PdfMerger()

    for dut, dutDict in s4pDict.items():
        for layer, layerDict in dutDict.items():
            inputFold = layerDict['outFold']
            pdfName = 'deltal_2l_report.pdf'
            #newPdfPath = os.path.join(os.getcwd(), inputFold,pdfName)
            newPdfPath = mergePdf(inputFold, pdfName)
            #,layerDict['outFold'], os.sep, 'deltal_2l_report.pdf'
            #print(newPdfPath)
            mergedObject.append(PdfReader(newPdfPath, 'rb'))

    outputFilePath = os.path.join(collectResFold, "SummaryReportPdf.pdf")
    mergedObject.write(outputFilePath)

    return

####################
# [Main]
def myGetMac():
    mac = get_mac()
    mac = ':'.join(("%012X" % mac)[i:i + 2] for i in range(0, 12, 2))
    return mac

if __name__ == '__main__':
    #importlib.reload(sys)
    #sys.setdefaultencoding('utf-8')
    #result = os.system('aitt.exe -h')
    #print(f'result:{result}')
    #sys.exit(0)

    #print(myGetMac())
    #[Step0] preprocess
    print("Current Working Directory:{0} ", os.getcwd())
    dataFold = os.getcwd()
    
    collectResFold = os.path.join(dataFold, 'output')
    if os.path.exists(collectResFold):
        shutil.rmtree(collectResFold)

    # os.system(f'mkdir {collectResFold}')
    s4pDict = readConfig(dataFold, collectResFold)
    # print(s4pDict['AD001']['L01'])

    for dev, layerDict in s4pDict.items():
        for layer, db in layerDict.items():
            output_fold_per_layer = db['outFold']
            os.system(f'mkdir {output_fold_per_layer}')

    #os.system("C:\Program Files\Advanced Interconnect Test Tool (64-Bit)\aitt.exe")
    #path = os.path.join('C:', os.sep, 'meshes', 'as')


    os.chdir(g_configDict['aittFold'])
    for dut in g_dut_list:
        for layer in g_layer_list:
            #print(s4pDict[dut][layer]['aittCmd'])
            os.system(s4pDict[dut][layer]['aittCmd'])
            print("[INFO][aitt.exe]process done {dut}_{layer}".format(dut=dut, layer=layer))
    os.chdir(dataFold)
    print()

    #sys.exit(0)

    runFreqMag(s4pDict, dataFold)

    #aittCmd = s4pDict['AD001']['L01']['aittCmd']
    #print(aittCmd)
    #os.system('aitt.exe -h')
    #os.system(aittCmd)


    #[Step1] Summary Report
    workbook = genWorkBook(collectResFold)
    summarySheet = workbook.add_worksheet('Summary')
    dataSheet = workbook.add_worksheet('Data')
    picSheet = workbook.add_worksheet('Picture')

    for dut, dutDict in s4pDict.items():
        make_database(workbook, dataSheet, dut, dutDict)

    run_data_sheet(workbook, dataSheet, s4pDict)
    run_summary_sheet(workbook, summarySheet, s4pDict)
    runLayerSheet(workbook,s4pDict)
    runPictureSheet(workbook, picSheet, s4pDict)

    finishWorkBook(workbook)

    #########################
    #[Step2] PDF report
    generatePdfReport(s4pDict, collectResFold)

    print("")
    print("[INFO] All process done !!!")
    print("[INFO] All process done !!!")
    print("[INFO] All process done !!!")
    os.system("PAUSE")

    sys.exit(0)
