#!/usr/bin/python3

import requests
import time
from xlwt import Workbook

def getQuoteData(stock, payload):
    ep = "https://api.tdameritrade.com/v1/marketdata/{}/quotes".format(stock)
    data = requests.get(url = ep , params = payload).json()
    return data

def getFundamentalData(stock, key, payload):
    ep = "https://api.tdameritrade.com/v1/instruments?apikey={}&symbol={}&projection=fundamental".format(key, stock)
    data = requests.get(url = ep , params = payload).json()
    return data

def parseFundamentalData(fundamentalData, stock):
   return list(fundamentalData[stock]["fundamental"].values())

def parseQuoteData(quoteData, stock):
    return list(quoteData[stock].values())
    
def fill(sheet, values, row, start):
    for x, value in zip(range(start, len(values)+start), values):
        sheet.write(row, x, value)

def fillNoData(sheet, index, row):
    for i in range(1, index):
        sheet.write(row, i, "No Data")

def failedToGetData():
	fillNoData(sheet1, S1Headings, row)
	fillNoData(sheet2, S2Headings, row)
	sheet2.write(row, 0 , stock)


def Initialize():
	stock = "AAPL"
	Qdata = getQuoteData(stock, payload)
	Fdata = getFundamentalData(stock, key, payload)
	if Qdata and Fdata:
		print("Connection to API successful")
		sheet1.write(0, 0, "symbol")
		fill(sheet1, list(Qdata[stock].keys()), 0, 1) 
		fill(sheet2, list(Fdata[stock]["fundamental"].keys()), 0, 0)
	time.sleep(0.5)
	return len(list(Qdata[stock].keys()))+1, len(list(Fdata[stock]["fundamental"].keys()))

init = time.time()
stocks = ['AAB', 'AAV', 'ABX', 'AC', 'ACB', 'ACD','ACO-X', 'ACQ', 'ADN', 'ADW-A', 'AEM', 'AEZS', 'AFN', 'AGF-B', 'AGI', 'AI', 'AIF', 'AIM', 'AJX', 'AKT-A', 'AKU', 'ALA', 'ALC', 'ALS', 'ALYA', 'AMM', 'AND', 'ANX', 'AOI', 'AOT', 'AP-UN', 'AQA', 'AQN', 'AR', 'ARE', 'ARX', 'ASM', 'ASND', 'ASP', 'AT', 'ATA', 'ATD-A', 'ATZ', 'AUP', 'AVCN', 'AVL', 'AW-UN', 'AXU', 'AZZ', 'BAM-A', 'BB', 'BBD-A', 'BBL-A', 'BBU-UN', 'BCE', 'BDI', 'BDT', 'BEI-UN', 'BHC', 'BIR', 'BKI', 'BLDP', 'BLU', 'BLX', 'BMO', 'BNS', 'BOS', 'BPF-UN', 'BR', 'BRE', 'BSO-UN', 'BSX', 'BTE', 'BTO', 'BU', 'BUI', 'BYD', 'BYL', 'CAE', 'CAR-UN', 'CAS', 'CCA', 'CCL-A', 'CCM', 'CCO', 'CDAY', 'CEE', 'CERV', 'CET', 'CEU', 'CF', 'CFP', 'CFW', 'CFX', 'CG', 'CGG', 'CGO', 'CGX', 'CGY', 'CHP-UN', 'CHR', 'CIA', 'CIGI', 'CIQ-UN', 'CIX', 'CJ', 'CJR-B', 'CJT', 'CKI', 'CLIQ', 'CLS', 'CM', 'CMG', 'CMMC', 'CNE', 'CNQ', 'CNR', 'CNT', 'CNU', 'COG', 'CP', 'CPG', 'CPX', 'CR', 'CRDL', 'CRON', 'CRP', 'CRR-UN', 'CRT-UN', 'CRWN', 'CS', 'CSH-UN', 'CSM', 'CSU', 'CSW-A', 'CTC', 'CU', 'CUF-UN', 'CUP-U', 'CVE', 'CVG', 'CWB', 'CWEB', 'CWL', 'CXB', 'CXI', 'DBO', 'DC-A', 'DCBO', 'DCM', 'DII-A', 'DIR-UN', 'DML', 'DN', 'DNT', 'DOL', 'DOO', 'DPM', 'DR', 'DRM', 'DRT', 'DRX', 'DSG', 'E', 'ECN', 'EDR', 'EDV', 'EFL', 'EFN', 'EFR', 'EFX', 'EGLX', 'EIF', 'ELD', 'ELEF', 'ELF', 'EMA', 'EMP-A', 'ENB', 'ENGH', 'EOX', 'EQB', 'EQX', 'ERD', 'ERF', 'ERO', 'ESI', 'ESM', 'ESN', 'ET', 'EXE', 'EXF', 'EXN', 'FAF', 'FAR', 'FC', 'FCU', 'FEC', 'FF', 'FFH', 'FFI-UN', 'FIH-U', 'FM', 'FN', 'FNV', 'FOOD', 'FR', 'FRII', 'FRU', 'FRX', 'FSV', 'FSZ', 'FT', 'FTG', 'FTS', 'FTT', 'FVI', 'FXC', 'GBT', 'GC', 'GCG', 'GCL', 'GCM', 'GDC', 'GDI', 'GDL', 'GEI', 'GFL', 'GIB-A', 'GIL', 'GLG', 'GLO', 'GOLD', 'GOOS', 'GPR', 'GRT-UN', 'GSC', 'GSV', 'GSY', 'GTE', 'GTMS', 'GUD', 'GUY', 'GVC', 'GWO', 'GWR', 'GXE', 'H', 'HBM', 'HCG', 'HDI', 'HEXO', 'HLF', 'HLS', 'HOM-U', 'HPS-A', 'HR-UN', 'HRT', 'HRX', 'HSM', 'HUT', 'HWO', 'HWX', 'HZM', 'IAG', 'IBG', 'ICE', 'IDG', 'IFC', 'IFP', 'IGM', 'III', 'IIP-UN', 'IMG', 'IMO', 'IMP', 'IMV', 'IN', 'INE', 'INQ', 'IPCO', 'IPL', 'IPO', 'ISV', 'ITP', 'IVN', 'IVQ', 'JAG', 'JOSE', 'JWEL', 'K', 'KBL', 'KEL', 'KEY', 'KL', 'KLS', 'KMP-UN', 'KOR', 'KPT', 'KXS', 'L', 'LABS', 'LAC', 'LAS-A', 'LB', 'LCS', 'LGD', 'LGO', 'LGT-A', 'LN', 'LNF', 'LNR', 'LSPD', 'LUC', 'LUG', 'LUN', 'LXR', '', 'MAL', 'MAV', 'MAW', 'MAXR', 'MBA', 'MBX', 'MDF', 'MDI', 'MDNA', 'MEG', 'MFC', 'MFI', 'MG', 'MI-UN', 'MIN', 'MKP', 'MMP-UN', 'MMX', 'MND', 'MOGO', 'MOZ', 'MPVD', 'MRC', 'MRD', 'MRE', 'MRU', 'MSV','MTL', 'MTY', 'MUX', 'MX', 'MXG', 'NA', 'NB', 'NCP', 'NCU', 'NDM', 'NEO', 'NEPT', 'NEXA', 'NEXT', 'NFI', 'NG', 'NGD', 'NGT', 'NOA', 'NPI', 'NPK', 'NTR', 'NVA', 'NWC', 'NXE', 'NXJ', 'OBE', 'OGC', 'OGD', 'OGI', 'OLA', 'OLY', 'OMI', 'ONC', 'ONEX', 'OPT', 'OR', 'ORA', 'ORL', 'OSK', 'OSP', 'OTEX', 'OVV', 'PAAS', 'PBH', 'PBL', 'PD', 'PEY', 'PFB', 'PHO', 'PHX', 'PKI', 'PLC', 'PMN', 'PMT', 'PNC-A', 'POM', 'POU', 'POW', 'PPL', 'PRMW', 'PRN', 'PRU', 'PSD', 'PSI', 'PSK', 'PTM', 'PTS', 'PVG', 'PXT', 'QBR-A', 'QSP-UN', 'QSR', 'QTRH', 'RAY-A', 'RBA', 'RBN-UN', 'RCH', 'RCI-A', 'REAL', 'RECP', 'REI-UN', 'RFP', 'RNW', 'ROOT', 'RSI', 'RUS', 'RVX', 'RY', 'S', 'SAP', 'SAU', 'SBB', 'SCL', 'SCY', 'SEA', 'SES', 'SFC', 'SFD', 'SGQ', 'SGY', 'SHLE', 'SHOP', 'SIA', 'SII', 'SIL', 'SIS', 'SJ', 'SJR-B', 'SLF', 'SMC', 'SMT', 'SMU-UN', 'SNC', 'SOLG', 'SOY', 'SPB', 'SPG', 'SRU-UN', 'SRX', 'SSF-UN', 'SSL', 'SSRM', 'STEP', 'STGO', 'STLC', 'STN', 'SU', 'SVM', 'SW', 'SWP', 'SZLS', 'T', 'TA', 'TC', 'TCL-A', 'TCN', 'TCS', 'TCW', 'TD', 'TECK-A', 'TF', 'TFII', 'TGO', 'TGOD', 'TH', 'TI', 'TIH', 'TKO', 'TLG', 'TLO', 'TMD', 'TML', 'TMQ', 'TNX', 'TOT', 'TOU', 'TOY', 'TRI', 'TRIL', 'TRL', 'TRP', 'TRQ', 'TRZ', 'TSL', 'TSU', 'TV', 'TVA-B', 'TVE', 'TVK', 'TWC', 'TWM', 'TXG', 'TXP', 'UFS', 'UNI', 'UNS', 'UR', 'URE', 'USA', 'VB', 'VCM', 'VET', 'VFF', 'VGCX', 'VIVO', 'VLE', 'VLN', 'VMD', 'WCM-A', 'WCN', 'WCP', 'WDO', 'WEED', 'WEF', 'WELL', 'WFC', 'WILD', 'WIR-U', 'WJX', 'WLLW', 'WM', 'WN', 'WPK', 'WPM', 'WPRT', 'WRN', 'WRX', 'WSP', 'WTE', 'X', 'XAM', 'XAU', 'XCT', 'XTC', 'Y', 'YRI']
key = ""
payload = {'apikey': "{}".format(key)}
wb = Workbook()
sheet1 = wb.add_sheet('Quote Data')
sheet2  = wb.add_sheet('Fundamental Data')
row = 1

S1Headings, S2Headings = Initialize()

for stock in stocks:
    start = time.time()
    try:
        quoteData = getQuoteData(stock, payload)
        fundamentalData = getFundamentalData(stock, key, payload)
        sheet1.write(row, 0 , stock)
        if quoteData and fundamentalData: 	# data is available for stock
            values = parseQuoteData(quoteData, stock)
            fill(sheet1, values, row, 1)
            values = parseFundamentalData(fundamentalData, stock)
            fill(sheet2, values, row, 0)
        else:    				# data is not available for stock
            failedToGetData()
            print("No Data Available for {}".format(stock))
    except:
        print("Unknown Data Exception")

    row+=1
    difference = time.time() - start
    if not difference > 1:
        time.sleep(1-difference)

wb.save('data.xls')
print("Data Saved")
print("Ran in {:.2f} Seconds".format(time.time()-init))
