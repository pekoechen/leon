# -*- coding:utf-8 -*-
import sys
import collections
import os
import csv
import math
import xlsxwriter
from xlsxwriter.utility import xl_rowcol_to_cell
from xlsxwriter.utility import xl_range

from PyPDF2 import PdfFileWriter, PdfFileReader, PdfFileMerger
import io
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter

import matplotlib.pyplot as plt
from matplotlib.ticker import MultipleLocator, FormatStrFormatter
import numpy as np

#reload(sys)
#sys.setdefaultencoding('utf-8')

g_cellFreqListStart = None
g_cellFreqListEnd = None

g_devName = 'Not assigned'

g_dut_list = []
g_layer_list = []
g_length_list = []
#g_sampleFreqList = [1, 3, 4, 6, 8, 10 ,12.89, 16, 20, 25, 40]
g_sampleFreqList = [4, 8, 12.89, 16, 28]

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
    read_common(filePath, db_common)
    return

def read_uncertainty_plotl1l2(filePath, db):
    key_list = ['Frequency (GHz)',
                'Fitted',
                'Insertion Loss']
    freq_list = []
    fitted_list = []
    iLoss_list = []

    with open(filePath, newline='') as csvFile:
        csv_reader = csv.DictReader(csvFile, delimiter=',')
        for row in csv_reader:
            freq_list.append(float(row['Frequency (GHz)']))
            fitted_list.append(float(row['Fitted']))
            iLoss_list.append(float(row['Insertion Loss']))
    return freq_list, fitted_list, iLoss_list

def genWorkBook():
    print("[INFO][SummaryReport] start ===> ")
    workbook = xlsxwriter.Workbook('SummaryReport.xlsx')
    return workbook

def finishWorkBook(workbook):
    workbook.close()
    print("[INFO][SummaryReport] done <=== ")

def readConfig():
    def csv_read(file):
        with open(file, newline='') as f:
            row_list = f.read().splitlines()
            return row_list

    global g_dut_list, g_layer_list, g_length_list
    g_dut_list = csv_read('dut.csv')
    g_layer_list = csv_read('layer.csv')
    g_length_list = csv_read('length.csv')
    config_list = csv_read('config.csv')

    global g_devName
    g_devName = config_list[0].split(',')[1]
    #print(g_devName)

    s4pDict = {}
    for dut in g_dut_list:
        s4pDict[dut] = {}
        #s4pDict[dut]['name'] = dut
        for layer in g_layer_list:
            for length in g_length_list:
                fileName = "{0}_{1}_{2}.s4p".format(dut,layer,length)
                if not os.path.isfile(fileName):
                    print("[ERROR] {0} doest not exist!!".format(fileName))
                    sys.exit(-1)
                layerDict = s4pDict[dut].setdefault(layer,collections.OrderedDict())
                layerDict.setdefault(length, fileName)
                layerDict['outFold'] = "output_{0}_{1}".format(dut,layer)
                #print(fileName)

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
    # ['!', 'FREQ.GHZ', 'S11RE', 'S11IM', 'S12RE', 'S12IM', 'S13RE', 'S13IM', 'S14RE', 'S14IM']
    # ['!', 'S21RE', 'S21IM', 'S22RE', 'S22IM', 'S23RE', 'S23IM', 'S24RE', 'S24IM']
    # ['!', 'S31RE', 'S31IM', 'S32RE', 'S32IM', 'S33RE', 'S33IM', 'S34RE', 'S34IM']
    # ['!', 'S41RE', 'S41IM', 'S42RE', 'S42IM', 'S43RE', 'S43IM', 'S44RE', 'S44IM']
    def dumpFreqMagInfo(attrDict):
        print("<===== Freq Ghz[{0}] ========>".format(attrDict['FREQ.GHZ']))
        print("S31RE:{0}".format(attrDict['S31RE']))
        print("S31IM:{0}".format(attrDict['S31IM']))

        print("S32RE:{0}".format(attrDict['S32RE']))
        print("S32IM:{0}".format(attrDict['S32IM']))

        print("S41RE:{0}".format(attrDict['S41RE']))
        print("S41IM:{0}".format(attrDict['S41IM']))

        print("S42RE:{0}".format(attrDict['S42RE']))
        print("S42IM:{0}".format(attrDict['S42IM']))

    #filePath = 'AD001_L01_05IN.s4p'
    with open(filePath, newline='') as f:
        row_list = f.read().splitlines()

        rowUnit = 4
        rowIdxStart = 8
        rowIdxEnd = rowIdxStart + rowUnit
        attrIdx = 0
        attrIdxMap = {}
        for rowIdx in range(rowIdxStart, rowIdxEnd):
            for attr in row_list[rowIdx].split()[1:]:
                attrIdxMap[attrIdx] = attr
                attrIdx += 1

        row_list = row_list[13:]
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
            #dumpFreqMagInfo(attrDict)

            reVal = (attrDict['S31RE'] - attrDict['S41RE'] - attrDict['S32RE'] + attrDict['S42RE'])/2
            imVal = (attrDict['S31IM'] - attrDict['S41IM'] - attrDict['S32IM'] + attrDict['S42IM'])/2
            tmpVal = math.sqrt(pow(reVal, 2) + pow(imVal, 2))
            val = 20 * math.log(tmpVal, 10)
            freq = attrDict['FREQ.GHZ']

            freqInfo[0].append(freq)
            freqInfo[1].append(val)
            pass

        x = freqInfo[0]
        y = freqInfo[1]
        #plt.cla()
        #plt.grid(color='r', linestyle='--', linewidth=1, alpha=0.3)
        plt.grid(b=True)
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

def runFreqMag():
    for dut in g_dut_list:
        for layer in g_layer_list:
            outFilePath = "output_{dut}_{layer}/magnitude.png".format(dut=dut, layer=layer)
            #print(outFilePath)

            plt.cla()
            for leng in g_length_list:
                filePath = \
                    "{dut}_{layer}_{leng}.s4p".format(dut=dut, layer=layer, leng=leng)
            #    print(filePath)
                runFreqMagEach(filePath,leng)

            plt.savefig(outFilePath, dpi=300, format="png")

    print("[INFO][Magnitude] successfully generate")
    return
def make_database(workbook,dataSheet,dut,dutDict):
    for layer, layerDict in dutDict.items():
        outFold = layerDict['outFold']
        #print(outFold)

        csv1 = '{0}/freq_report_table.csv'.format(outFold)
        csv2 = '{0}/impedance_report_table.csv'.format(outFold)
        csv3 = '{0}/trace_report_table.csv'.format(outFold)
        csv4 = '{0}/uncertainty_plot_l1l2.csv'.format(outFold)

        db_l1l2 = {}
        db_common = {}
        db_uPlot = {}

        #read_freq_report_table(csv1, db_l1l2)
        #read_impedance_report_table(csv2, db_common)
        #read_trace_report_table(csv3, db_common)
        freq_list, fitted_list, insertLoss_list = read_uncertainty_plotl1l2(csv4, db_uPlot)

        #print(freq_list)
        #dutDict['name'] = dut
        layerDict['freqList'] = freq_list
        layerDict['fittedList'] = [float(i) for i in fitted_list]
        layerDict['iLossList'] = [float(i) for i in insertLoss_list]
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

    #Summary
    offset_row_summary = 4
    offset_col_summary = 1

    summarySheet.write(offset_row_summary, offset_col_summary, 'Summary', format_bold)
    summarySheet.merge_range(offset_row_summary + 1, offset_col_summary,
                       offset_row_summary + 2, offset_col_summary+1,
                       'Panel', getDefaultFormat(workbook,'#bedcf2'))

    offset_row_summary_sample = offset_row_summary + 1
    offset_col_summary_sample = offset_col_summary + 2
    layerSize = len(g_layer_list)
    sampleIdx = 0
    for sampleFreq in g_sampleFreqList:
        sampleStr = "(dB/in,@{0}GHz)".format(sampleFreq)
        offset_col = offset_col_summary_sample + (sampleIdx * layerSize)
        summarySheet.merge_range(offset_row_summary_sample, offset_col,
                                 offset_row_summary_sample, offset_col + (layerSize-1),
                                 sampleStr, getDefaultFormat(workbook, '#c3e4bc'))
        for layerIdx in range(layerSize):
            offset_row = offset_row_summary_sample + 1
            offset_col = offset_col_summary_sample + (sampleIdx * layerSize)
            summarySheet.write(offset_row, offset_col + layerIdx,
                               g_layer_list[layerIdx], getDefaultFormat(workbook, '#c3e4bc'))

        sampleIdx += 1

    offset_row_data = offset_row_summary + 3
    offset_col_data = offset_col_summary + 1
    dutIdx = 0

    for dut, ductDict in s4pDict.items():
        offset_row = offset_row_data + dutIdx
        offset_col = offset_col_data
        summarySheet.write(offset_row, offset_col,
                           dut, getDefaultFormat(workbook,'#bedcf2'))
        for sampleIdx in range(len(g_sampleFreqList)):
            sampleFreq = g_sampleFreqList[sampleIdx]
            for layerIdx in range(layerSize):
                layerDict = ductDict[g_layer_list[layerIdx]]
                freq_list = layerDict['freqList']
                iLoss_list = layerDict['iLossList']

                # Interpolate
                prevIdx, nextIdx = binarySearchPrevNext(sampleFreq, freq_list)
                deltaFreq = (freq_list[nextIdx] - freq_list[prevIdx])
                deltaSampleFreq = (sampleFreq - freq_list[prevIdx])
                ratio = deltaSampleFreq / deltaFreq
                valAfterInterpoplate = iLoss_list[prevIdx] + (ratio * (iLoss_list[nextIdx] - iLoss_list[prevIdx]))
                '''
                print("freq:{0},{1}, iLoss:{2},{3}, sample:{4}, ratio:{5}, deltaFreq:{6}".format(
                    freq_list[prevIdx], freq_list[nextIdx],
                    iLoss_list[prevIdx],iLoss_list[nextIdx],
                    sampleFreq, ratio, deltaFreq))
                '''

                offset_col = offset_col_data + 1 + (layerSize * sampleIdx)
                summarySheet.write(offset_row, offset_col + layerIdx,
                                   valAfterInterpoplate , getDefaultFormat(workbook, '#fbf5dc'))

        dutIdx += 1

    #statistic
    def run_statistic(offset_row_input, offset_col_input, attr, offset_row_summary, numOfDut):
        summarySheet.write(offset_row_input, offset_col_input,attr, getDefaultFormat(workbook,'#bedcf2'))
        offset_col_input += 1

        for sampleIdx in range(len(g_sampleFreqList)):
            for layerIdx in range(layerSize):
                offset_row = offset_row_input
                offset_col = offset_col_input + (sampleIdx * layerSize) + layerIdx

                if attr == 'Range':
                    max_cell = xl_rowcol_to_cell(offset_row-2, offset_col)
                    min_cell = xl_rowcol_to_cell(offset_row-1, offset_col)

                    val = "=({max_cell} - {min_cell})".format(
                        max_cell=max_cell, min_cell=min_cell)
                    summarySheet.write(offset_row, offset_col,
                                    val, getDefaultFormat(workbook, '#fbf5dc'))

                else:
                    offset_dut_start = offset_row_summary + 3
                    offset_dut_end = offset_dut_start + (numOfDut - 1)
                    dutRange = xl_range(offset_dut_start, offset_col,
                                        offset_dut_end, offset_col)
                    val = "={func}({range})".format(func=attr, range=dutRange)

                    colorHex = '#e8f426' if attr =='Average' else '#fbf5dc'
                    summarySheet.write(offset_row, offset_col,
                                    val, getDefaultFormat(workbook, colorHex))

    numOfDut = len(s4pDict)
    offset_row_statistic = offset_row_summary + 3 + numOfDut
    offset_col_statistic = offset_col_summary + 1
    run_statistic(offset_row_statistic+0, offset_col_statistic, 'Average', offset_row_summary, numOfDut)
    run_statistic(offset_row_statistic+1, offset_col_statistic, 'Max', offset_row_summary, numOfDut)
    run_statistic(offset_row_statistic+2, offset_col_statistic, 'Min', offset_row_summary, numOfDut)
    run_statistic(offset_row_statistic+3, offset_col_statistic, 'Range', offset_row_summary, numOfDut)

    return

def run_data_sheet(workbook, dataSheet, s4pDict):
    print("[INFO][SummaryReport] running DataSheet")

    global g_dut_list, g_layer_list, g_length_list
    tmpDut = next(iter(s4pDict.values()))
    tmpLayer = next(iter(tmpDut.values()))

    lengthStr = "-".join(g_length_list)

    # common part
    font_calibri = workbook.add_format({'font':'calibri'})
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

    dataSheet.write('A2', g_devName,format_freqAttr)
    dataSheet.merge_range('B2:B3', 'Freq.(GHz)',format_freqAttr)
    dataSheet.write('A3', 'Freq.(MHz)',format_freqAttr)
    dataSheet.set_column('A:B',len(g_devName))

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
        dataSheet.write(offset_row_freqGhz + freqIdx, offset_col_freqGhz, freq,format_freq)

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
        dataSheet.merge_range(offset_row_attr_dutName, offset_col_attr_dutName + offset_dut,
                              offset_row_attr_dutName, offset_col_attr_dutName + offset_dut + layerSize-1,
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

def runPictureSheet():
    row_jump = 18
    col_jump = 7
    idx = 0
    for dut in g_dut_list:
        for layer in g_layer_list:
            outFold = 'output_{dut}_{layer}'.format(dut=dut, layer=layer)
            filePath_magnitude = '{outFold}/magnitude.png'.format(outFold=outFold)
            offset_row = 1 + (row_jump * idx)
            offset_col = 1
            picSheet.insert_image(offset_row, offset_col, filePath_magnitude, {'x_scale': 0.65, 'y_scale': 0.65})

            offset_row = offset_row
            offset_col = offset_col + col_jump
            filePath_uncertainty = '{outFold}/uncertainty_plot_l1l2.png'.format(outFold=outFold)
            picSheet.insert_image(offset_row, offset_col, filePath_uncertainty, {'x_scale': 0.5, 'y_scale': 0.5})

            idx += 1
            #print("filePath_magnitude:{0}".format(filePath_magnitude))
            #print("filePath_uncertainty:{0}".format(filePath_uncertainty))
    print("[INFO][SummaryReport] running PictureSheet")
    return

def runLayerSheet():
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
                'name': "{0}_{1}".format(g_devName,dut),
                'categories': '=Data!{cellStart}:{cellEnd}'.format(cellStart = g_cellFreqListStart,
                                                                   cellEnd = g_cellFreqListEnd),
                'values': '=Data!{cellStart}:{cellEnd}'.format(cellStart = cellStart,
                                                               cellEnd = cellEnd),
                'line': {'width': 2}
            })

        # Add a chart title and some axis labels.
        chart1.set_title({'name': g_devName})
        chart1.set_x_axis({'name': 'Frequency (GHz)'})
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
    existing_pdf = PdfFileReader(packet)
    # read your existing PDF
    filePath = os.path.join(os.getcwd(), inputFold, pdfName)
    new_pdf = PdfFileReader(open(filePath, "rb"))

    output = PdfFileWriter()
    # add the "watermark" (which is the new pdf) on the existing page
    page = existing_pdf.getPage(0)
    page.mergePage(new_pdf.getPage(0))
    output.addPage(page)
    # finally, write "output" to a real file
    newPdfName = "tmp_deltal_2l_report.pdf"
    filePath = os.path.join(os.getcwd(), inputFold, newPdfName)
    outputStream = open(filePath, "wb")
    output.write(outputStream)
    outputStream.close()
    return filePath

def generatePdfReport(s4pDict):
    print("[INFO][SummaryReportPdf] running...")
    mergedObject = PdfFileMerger()

    for dut, dutDict in s4pDict.items():
        for layer, layerDict in dutDict.items():
            inputFold = layerDict['outFold']
            pdfName = 'deltal_2l_report.pdf'
            #newPdfPath = os.path.join(os.getcwd(), inputFold,pdfName)
            newPdfPath = mergePdf(inputFold, pdfName)
            #,layerDict['outFold'], os.sep, 'deltal_2l_report.pdf'
            #print(newPdfPath)
            mergedObject.append(PdfFileReader(newPdfPath, 'rb'))

    mergedObject.write("SummaryReportPdf.pd"
                       "f")

    return

####################
# [Main]
if __name__ == '__main__':
    #[Step0] preprocess
    s4pDict = readConfig()
    runFreqMag()
    '''
    #s4pDict = readConfig()
    print(s4pDict)
    #for dut, dutDict in s4pDict.items():
    #    run_dut_layer(dutDict)

    #os.system("C:\Program Files\Advanced Interconnect Test Tool (64-Bit)\aitt.exe")
    #path = os.path.join('C:', os.sep, 'meshes', 'as')
    print("Current Working Directory:{0} ", os.getcwd())
    os.chdir("C:\Program Files\Advanced Interconnect Test Tool (64-Bit)")
    print("Current Working Directory:{0} ", os.getcwd())
    #os.system('aitt.exe -h')
    sys.exit(0)
    '''

    #[Step1] Summary Report
    workbook = genWorkBook()
    summarySheet = workbook.add_worksheet('Summary')
    dataSheet = workbook.add_worksheet('Data')
    picSheet = workbook.add_worksheet('Picture')

    for dut, dutDict in s4pDict.items():
        #print(dutDict)
        make_database(workbook, dataSheet, dut, dutDict)

    run_data_sheet(workbook, dataSheet, s4pDict)
    run_summary_sheet(workbook, summarySheet, s4pDict)
    runLayerSheet()
    runPictureSheet()

    finishWorkBook(workbook)

    #########################
    #[Step2] PDF report
    generatePdfReport(s4pDict)

    print("")
    print("[INFO] All process done !!!")
    print("[INFO] All process done !!!")
    print("[INFO] All process done !!!")

    '''
    s4pDict = readConfig()
    print(s4pDict)
    for dut, dutDict in s4pDict.items():
        run_dut_layer(dutDict)
    # runTransfer()
    #print(fileList)
    #os.system("C:\Program Files\Advanced Interconnect Test Tool (64-Bit)\aitt.exe")
    #path = os.path.join('C:', os.sep, 'meshes', 'as')
    #print("Current Working Directory ", os.getcwd())
    #os.chdir("C:\Program Files\Advanced Interconnect Test Tool (64-Bit)")
    #print("Current Working Directory ", os.getcwd())
    #os.system('aitt.exe -h')
    '''

    sys.exit(0)