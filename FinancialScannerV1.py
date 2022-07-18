#!/usr/bin/python3

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
import time
from xlwt import Workbook
stocknames = ['AAB', 'AAV', 'ABX', 'AC', 'ACB', 'ACD', 'ACO-X', 'ACQ', 'ADN', 'ADW-A', 'AEM', 'AEZS', 'AFN', 'AGF-B', 'AGI', 'AI', 'AIF', 'AIM', 'AJX', 'AKT-A', 'AKU', 'ALA', 'ALC', 'ALS', 'ALYA', 'AMM', 'AND', 'ANX', 'AOI', 'AOT', 'AP-UN', 'AQA', 'AQN', 'AR', 'ARE', 'ARX', 'ASM', 'ASND', 'ASP', 'AT', 'ATA', 'ATD-A', 'ATZ', 'AUP', 'AVCN', 'AVL', 'AW-UN', 'AXU', 'AZZ', 'BAM-A', 'BB', 'BBD-A', 'BBL-A', 'BBU-UN', 'BCE', 'BDI', 'BDT', 'BEI-UN', 'BHC', 'BIR', 'BKI', 'BLDP', 'BLU', 'BLX', 'BMO', 'BNS', 'BOS', 'BPF-UN', 'BR', 'BRE', 'BSO-UN', 'BSX', 'BTE', 'BTO', 'BU', 'BUI', 'BYD', 'BYL', 'CAE', 'CAR-UN', 'CAS', 'CCA', 'CCL-A', 'CCM', 'CCO', 'CDAY', 'CEE', 'CERV', 'CET', 'CEU', 'CF', 'CFP', 'CFW', 'CFX', 'CG', 'CGG', 'CGO', 'CGX', 'CGY', 'CHP-UN', 'CHR', 'CIA', 'CIGI', 'CIQ-UN', 'CIX', 'CJ', 'CJR-B', 'CJT', 'CKI', 'CLIQ', 'CLS', 'CM', 'CMG', 'CMMC', 'CNE', 'CNQ', 'CNR', 'CNT', 'CNU', 'COG', 'CP', 'CPG', 'CPX', 'CR', 'CRDL', 'CRON', 'CRP', 'CRR-UN', 'CRT-UN', 'CRWN', 'CS', 'CSH-UN', 'CSM', 'CSU', 'CSW-A', 'CTC', 'CU', 'CUF-UN', 'CUP-U', 'CVE', 'CVG', 'CWB', 'CWEB', 'CWL', 'CXB', 'CXI', 'DBO', 'DC-A', 'DCBO', 'DCM', 'DII-A', 'DIR-UN', 'DML', 'DN', 'DNT', 'DOL', 'DOO', 'DPM', 'DR', 'DRM', 'DRT', 'DRX', 'DSG', 'E', 'ECN', 'EDR', 'EDV', 'EFL', 'EFN', 'EFR', 'EFX', 'EGLX', 'EIF', 'ELD', 'ELEF', 'ELF', 'EMA', 'EMP-A', 'ENB', 'ENGH', 'EOX', 'EQB', 'EQX', 'ERD', 'ERF', 'ERO', 'ESI', 'ESM', 'ESN', 'ET', 'EXE', 'EXF', 'EXN', 'FAF', 'FAR', 'FC', 'FCU', 'FEC', 'FF', 'FFH', 'FFI-UN', 'FIH-U', 'FM', 'FN', 'FNV', 'FOOD', 'FR', 'FRII', 'FRU', 'FRX', 'FSV', 'FSZ', 'FT', 'FTG', 'FTS', 'FTT', 'FVI', 'FXC', 'GBT', 'GC', 'GCG', 'GCL', 'GCM', 'GDC', 'GDI', 'GDL', 'GEI', 'GFL', 'GIB-A', 'GIL', 'GLG', 'GLO', 'GOLD', 'GOOS', 'GPR', 'GRT-UN', 'GSC', 'GSV', 'GSY', 'GTE', 'GTMS', 'GUD', 'GUY', 'GVC', 'GWO', 'GWR', 'GXE', 'H', 'HBM', 'HCG', 'HDI', 'HEXO', 'HLF', 'HLS', 'HOM-U', 'HPS-A', 'HR-UN', 'HRT', 'HRX', 'HSM', 'HUT', 'HWO', 'HWX', 'HZM', 'IAG', 'IBG', 'ICE', 'IDG', 'IFC', 'IFP', 'IGM', 'III', 'IIP-UN', 'IMG', 'IMO', 'IMP', 'IMV', 'IN', 'INE', 'INQ', 'IPCO', 'IPL', 'IPO', 'ISV', 'ITP', 'IVN', 'IVQ', 'JAG', 'JOSE', 'JWEL', 'K', 'KBL', 'KEL', 'KEY', 'KL', 'KLS', 'KMP-UN', 'KOR', 'KPT', 'KXS', 'L', 'LABS', 'LAC', 'LAS-A', 'LB', 'LCS', 'LGD', 'LGO', 'LGT-A', 'LN', 'LNF', 'LNR', 'LSPD', 'LUC', 'LUG', 'LUN', 'LXR', '', 'MAL', 'MAV', 'MAW', 'MAXR', 'MBA', 'MBX', 'MDF', 'MDI', 'MDNA', 'MEG', 'MFC', 'MFI', 'MG', 'MI-UN', 'MIN', 'MKP', 'MMP-UN', 'MMX', 'MND', 'MOGO', 'MOZ', 'MPVD', 'MRC', 'MRD', 'MRE', 'MRU', 'MSV','MTL', 'MTY', 'MUX', 'MX', 'MXG', 'NA', 'NB', 'NCP', 'NCU', 'NDM', 'NEO', 'NEPT', 'NEXA', 'NEXT', 'NFI', 'NG', 'NGD', 'NGT', 'NOA', 'NPI', 'NPK', 'NTR', 'NVA', 'NWC', 'NXE', 'NXJ', 'OBE', 'OGC', 'OGD', 'OGI', 'OLA', 'OLY', 'OMI', 'ONC', 'ONEX', 'OPT', 'OR', 'ORA', 'ORL', 'OSK', 'OSP', 'OTEX', 'OVV', 'PAAS', 'PBH', 'PBL', 'PD', 'PEY', 'PFB', 'PHO', 'PHX', 'PKI', 'PLC', 'PMN', 'PMT', 'PNC-A', 'POM', 'POU', 'POW', 'PPL', 'PRMW', 'PRN', 'PRU', 'PSD', 'PSI', 'PSK', 'PTM', 'PTS', 'PVG', 'PXT', 'QBR-A', 'QSP-UN', 'QSR', 'QTRH', 'RAY-A', 'RBA', 'RBN-UN', 'RCH', 'RCI-A', 'REAL', 'RECP', 'REI-UN', 'RFP', 'RNW', 'ROOT', 'RSI', 'RUS', 'RVX', 'RY', 'S', 'SAP', 'SAU', 'SBB', 'SCL', 'SCY', 'SEA', 'SES', 'SFC', 'SFD', 'SGQ', 'SGY', 'SHLE', 'SHOP', 'SIA', 'SII', 'SIL', 'SIS', 'SJ', 'SJR-B', 'SLF', 'SMC', 'SMT', 'SMU-UN', 'SNC', 'SOLG', 'SOY', 'SPB', 'SPG', 'SRU-UN', 'SRX', 'SSF-UN', 'SSL', 'SSRM', 'STEP', 'STGO', 'STLC', 'STN', 'SU', 'SVM', 'SW', 'SWP', 'SZLS', 'T', 'TA', 'TC', 'TCL-A', 'TCN', 'TCS', 'TCW', 'TD', 'TECK-A', 'TF', 'TFII', 'TGO', 'TGOD', 'TH', 'TI', 'TIH', 'TKO', 'TLG', 'TLO', 'TMD', 'TML', 'TMQ', 'TNX', 'TOT', 'TOU', 'TOY', 'TRI', 'TRIL', 'TRL', 'TRP', 'TRQ', 'TRZ', 'TSL', 'TSU', 'TV', 'TVA-B', 'TVE', 'TVK', 'TWC', 'TWM', 'TXG', 'TXP', 'UFS', 'UNI', 'UNS', 'UR', 'URE', 'USA', 'VB', 'VCM', 'VET', 'VFF', 'VGCX', 'VIVO', 'VLE', 'VLN', 'VMD', 'WCM-A', 'WCN', 'WCP', 'WDO', 'WEED', 'WEF', 'WELL', 'WFC', 'WILD', 'WIR-U', 'WJX', 'WLLW', 'WM', 'WN', 'WPK', 'WPM', 'WPRT', 'WRN', 'WRX', 'WSP', 'WTE', 'X', 'XAM', 'XAU', 'XCT', 'XTC', 'Y', 'YRI']
urltemplate =[] # an array of the pages for each stock
infoheadings = ["Name",'Market Cap (intraday)', 'Enterprise Value', 'Trailing P/E', 'Forward P/E', 'PEG Ratio (5 yr expected)', 'Price/Sales', 'Price/Book', 'Enterprise Value/Revenue', 'Enterprise Value/EBITDA', 'Beta (5Y Monthly)', '52-Week Change', 'S&P500 52-Week Change', '52 Week High', '52 Week Low', '50-Day Moving Average', '200-Day Moving Average', 'Avg Vol (3 month)', 'Avg Vol (10 day)', 'Shares Outstanding', 'Implied Shares Outstanding', 'Float', '% Held by Insiders', '% Held by Institutions', 'Shares Short (Aug 13, 2021)', 'Short Ratio (Aug 13, 2021)', 'Short % of Float (Aug 13, 2021)', 'Short % of Shares Outstanding (Aug 13, 2021)', 'Shares Short (prior month Jul 15, 2021)', 'Forward Annual Dividend Rate', 'Forward Annual Dividend Yield', 'Trailing Annual Dividend Rate', 'Trailing Annual Dividend Yield', '5 Year Average Dividend Yield', 'Payout Ratio', 'Dividend Date', 'Ex-Dividend Date', 'Last Split Factor', 'Last Split Date', 'Fiscal Year Ends', 'Most Recent Quarter', 'Profit Margin', 'Operating Margin', 'Return on Assets', 'Return on Equity', 'Revenue', 'Revenue Per Share', 'Quarterly Revenue Growth', 'Gross Profit', 'EBITDA', 'Net Income Avi to Common', 'Diluted EPS', 'Quarterly Earnings Growth', 'Total Cash', 'Total Cash Per Share', 'Total Debt', 'Total Debt/Equity', 'Current Ratio', 'Book Value Per Share', 'Operating Cash Flow', 'Levered Free Cash Flow',"Sector","Industry","FT Employees","1st executive","2nd executive","3rd executive","4th executive","5th executive","Description","General URL","Statistics URL", "Financials URL", "Holders URL"]
row = 1

start = time.time()

def Refiner(word):
    try:
        factor=1
        ending = word[-1]
        if (ending != "A"):
            if(ending == "M"):
                factor = 1000000
            elif(ending == "B"):
                factor = 1000000000
            elif(ending == "T"):
                factor = 1000000000000 
            elif(ending == "k"):
                factor = 1000
            return float(word[:-1])*factor
        else:    
            return word
    except:
        return 1


serv = Service(r"path")
driver = webdriver.Firefox(service=serv)
wb = Workbook()
sheet = wb.add_sheet('Data')

for x,y in zip(infoheadings,range(len(infoheadings))):
    sheet.write(0,y,x)

for name in stocknames:
    try:
        driver.get(urltemplate[1].format(name))
        time.sleep(1)
        info = []
        info.append(driver.find_element(by=By.XPATH, value="/html/body/div[1]/div/div/div[1]/div/div[2]/div/div/div[5]/div/div/div/div[2]/div[1]/div[1]/h1").text)
        for x in range(1,4):
            for y in range(1,7):
                for c in range(1,14):
                    try:
                        temp = Refiner(driver.find_element(by=By.XPATH, value="/html/body/div[1]/div/div/div[1]/div/div[3]/div[1]/div/div[1]/div/div/section/div[2]/div[{0}]/div/div[{1}]/div/div/table/tbody/tr[{2}]/td[2]".format(x,y,c)).text)
                        info.append(temp)
                    except:
                        pass  
        driver.get(urltemplate[3].format(name))
        time.sleep(1)
        for x in range(2,7,2):
            info.append(driver.find_element(by=By.XPATH, value="/html/body/div[1]/div/div/div[1]/div/div[3]/div[1]/div/div[1]/div/div/section/div[1]/div/div/p[2]/span[{}]".format(x)).text)
        for x in range(1,6):
            info.append(driver.find_element(by=By.XPATH, value="/html/body/div[1]/div/div/div[1]/div/div[3]/div[1]/div/div[1]/div/div/section/section[1]/table/tbody/tr[{}]/td[1]/span".format(x)).text)
        info.append(driver.find_element(by=By.XPATH, value="/html/body/div[1]/div/div/div[1]/div/div[3]/div[1]/div/div[1]/div/div/section/section[2]/p").text)
        for template in urltemplate:
            info.append(template.format(name))
        for x,y in zip(info,range(len(info))):
            sheet.write(row,y,x)
        row+=1
    except:
        print("Failed to retrieve data for: {}".format(name))
wb.save('data.xls')
driver.quit()
print("Data Saved")
print("Finished execution in {:.2f} seconds".format(time.time() - start))
