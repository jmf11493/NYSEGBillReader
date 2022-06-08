import PyPDF2
import re
import os
import json
import xlsxwriter
from datetime import date
from datetime import datetime

# Load sensitive data from config file
test_file_directory = ''
pod_id = ''
meter_number = ''
nyseg_csv_directory = ''
config_file_path = os.getcwd() + '\config.csv'

with open(config_file_path,'r') as config_file:
    contents = config_file.read()
    data = contents.split(',')
    test_file_directory = data[0]
    pod_id = data[1]
    meter_number = data[2]
    nyseg_csv_directory = data[3]

# Fix the pdf 
EOF_MARKER = b'%%EOF'
directory = test_file_directory

month                 = []
billing_cycle         = []
excess_generation     = []
current_generation    = []
current_useage        = []
net_usage_billed      = []
remaining_excess_gen  = []
average_daily_use     = []
average_daily_temp    = []
total_energy_charges  = []
theoretical_charges   = []
theoretical_savings   = []
basic_service_charges = []
price_per_kwh_est     = []
file_names            = []
actual_pkwh           = []
# Not every bill has an associated charge for electric (free months) set it to 0.1 to start
kwh_charge = 0.1
for filename in os.listdir(directory):
    if os.path.splitext(filename)[1] != '.pdf':
        continue
    file_names.append(filename)
    path = os.path.join(directory, filename)
    with open(path,'rb') as file:
        contents = file.read()
    # Just remove everything after the EOF, remove this condition
    if b'<!DOCTYPE html>' in contents:
        
        remove2 = contents.split(EOF_MARKER)[1]
        contents = contents.replace(remove2, b'')
        with open(path, 'wb') as file:
            file.write(contents)
            file.close()
    
    with open(path, 'rb') as f:
        contents = f.read()
    
    
    pdfFileOjb = open(path, 'rb')
    
    doContinue = False
    try:
        pdfReader = PyPDF2.PdfFileReader(pdfFileOjb)
    except Exception as exc:
        doContinue = True
        print('Skipping file: ' + path)
        print(exc)
    if doContinue:
        continue
    ## Attempt to get the kwh charge
    page2 = pdfReader.getPage(1)
    page2Contents = page2.extractText()
    supplyStr = "providing electricity supply during this billing period was"
    actualPpkwh = "n/a"
    if supplyStr in page2Contents:
        ppkwhStr = page2Contents.split(supplyStr)[1]
        actualPpkwh = ppkwhStr.split('/kwh')[0].strip()
    actual_pkwh.append(actualPpkwh)
    print('Actual ppkwh: ' + actualPpkwh)
    ## Page 3 has the important information
    page3 = pdfReader.getPage(2)
    contents = page3.extractText()
    basic_service_charge = ''
    if 'Basic service charge -' in contents:
        basic_service_charge = contents.split('Basic service charge -')
        basic_service_charge_1 = basic_service_charge[1][4:7]
        basic_service_charge_2 = basic_service_charge[2][4:7]
        basic_service_charge = str(float(basic_service_charge_1) + float(basic_service_charge_2))
    else:
        basic_service_charge = contents.split('Basic service charge')[1][:5]
    print(basic_service_charge)
    serviceFrom   = contents.split('Service from:PoD ID:')[1].split(pod_id)[0]
    print(serviceFrom)
    importantInfo = contents.split('ExcessGenerationCurrent UsageMeter Number')[1].split(meter_number)[0].split(' kwh')
    excessGeneration      = importantInfo[0]
    currentGeneration     = importantInfo[1]
    priorExcessGeneration = importantInfo[2]
    currentUsage          = importantInfo[3]
    netUsageBilled        = importantInfo[4]
    
    totalEnergyCharges = contents.split('Total Energy Charges')[1].split('BillingPeriodAverageDaily')[0].split('$')[1]
    
    # Average Info
    averageInfo = contents.split('Usage Chart Information')[1].split('Miscellaneous Charges-')[0].split('kwh')
    # 14 ,59  F26 ,62  FOct-20Oct-19
    averageDailyUseCurrentYear  = averageInfo[0].strip()
    averageDailyTempCurrentYear = averageInfo[1].split(' ')[0]
    averageDailyUsePastYear     = averageInfo[1].split('F')[1].strip()
    averageDailyTempPastYear    = averageInfo[2].split(' F')[0]
    currentYear                 = averageInfo[2].split(' F')[1][:6]
    pastYear                    = averageInfo[2].split(' F')[1][6:12]
    # Total Energy Charges
    charge = contents.split('Total Energy Charges')[1].split('BillingPeriodAverageDaily')[0].split('$')[1];
    
    # If billed for electric, update the price per kwh charge
    if int(netUsageBilled) > 0:
        kwh_charge = (float(charge) - float(basic_service_charge)) / float(netUsageBilled)
    print('Price per kwh: ' + str(kwh_charge))
    
    peg = 'Prior Excess Generation: '        + priorExcessGeneration       + ' kwh\n'
    cg  = 'Current Generation: '             + currentGeneration           + ' kwh\n'
    cu  = 'Current Usage: '                  + currentUsage                + ' kwh\n'
    nub = 'Net Usage Billed: '               + netUsageBilled              + ' kwh\n'
    reg = 'Remaining Excess Generation: '    + excessGeneration            + ' kwh\n'
    avgd   = 'Average Daily Use ' + currentYear  + ': ' + averageDailyUseCurrentYear  + ' kwh\n'
    avgdly = 'Average Daily Use ' + pastYear     + ': ' + averageDailyUsePastYear     + ' kwh\n'
    avgt   = 'Average Daily Temp ' + currentYear + ': ' + averageDailyTempCurrentYear + ' F\n'
    avgtly = 'Average Daily Temp ' + pastYear    + ': ' + averageDailyTempPastYear    + ' F\n'
    charges = 'Total Energy Charges: $' + charge
    
    # Billing Cycle
    print(peg + cg + cu + nub + reg + avgd + avgdly + avgt + avgtly + charges)
    print('-------------------------------------')
    month.append(currentYear)
    billing_cycle.append(serviceFrom)
    excess_generation.append(priorExcessGeneration)
    current_generation.append(currentGeneration)
    current_useage.append(currentUsage)
    net_usage_billed.append(netUsageBilled)
    remaining_excess_gen.append(excessGeneration)
    average_daily_use.append(averageDailyUseCurrentYear)
    average_daily_temp.append(averageDailyTempCurrentYear)
    total_energy_charges.append(charge)
    basic_service_charges.append(float(basic_service_charge))
    theoretical_charges.append((float(currentUsage) * kwh_charge) + float(basic_service_charge))
    theoretical_savings.append(((float(currentUsage) * kwh_charge) + float(basic_service_charge)) - float(charge))
    price_per_kwh_est.append(kwh_charge)
    # print(pdfReader.numPages)
row = 1
column = 0
workbook = xlsxwriter.Workbook(directory + '\solar panels.xlsx')
worksheet = workbook.add_worksheet()

red_format = workbook.add_format({'bg_color': 'red'})
green_format = workbook.add_format({'bg_color': 'green'})
yellow_format = workbook.add_format({'bg_color': 'yellow'})
blue_format = workbook.add_format({'bg_color': 'blue', 'font_color': 'white'})

month_text                     = 'Month'
excess_gen_text                = 'Prior Excess Generation (kwh)'
contribution_text              = 'NYSEG Contribution to Grid (kwh)'
useage_text                    = 'NYSEG Useage (kwh)'
useage_billed_text             = 'NYSEG Useage Billed'
billed_amount_text             = 'NYSEG Billed Amount'
billing_cycle_text             = 'NYSEG Billing Cycle'
days_text                      = 'Billing Days'
avg_daily_tmp_text             = 'Avg Daily Tmp'
avg_daily_useage_text          = 'Avg Daily Useage'
basic_service_text             = 'Basic Service Charge'
theoretical_charge_text        = 'Theoretical Charge'
theoretical_savings_text       = 'Theoretical Savings'
price_per_kwh                  = 'Calculated Price Per KWH'
actual_price_per_kwh           = 'Actual Price Per KWH'
# theoretical_savings_total_text = 'Theoretical Savings Total'
worksheet.write(0, 0, month_text)
worksheet.write(0, 1, excess_gen_text, blue_format)
worksheet.write(0, 2, contribution_text, green_format)
worksheet.write(0, 3, useage_text, red_format)
worksheet.write(0, 4, useage_billed_text, yellow_format)
worksheet.write(0, 5, billed_amount_text)
worksheet.write(0, 6, billing_cycle_text)
worksheet.write(0, 7, days_text)
worksheet.write(0, 8, avg_daily_tmp_text)
worksheet.write(0, 9, avg_daily_useage_text)
worksheet.write(0, 10, basic_service_text)
worksheet.write(0, 11, theoretical_charge_text)
worksheet.write(0, 12, theoretical_savings_text)
worksheet.write(0, 13, price_per_kwh)
worksheet.write(0, 14, actual_price_per_kwh)
# worksheet.write(0, 12, theoretical_savings_total_text)
worksheet.set_column(0, 0, len(month_text))
worksheet.set_column(0, 1, len(excess_gen_text)+3)
worksheet.set_column(0, 2, len(contribution_text)+5)
worksheet.set_column(0, 3, len(useage_text))
worksheet.set_column(0, 4, len(useage_billed_text))
worksheet.set_column(0, 5, len(billed_amount_text))
worksheet.set_column(0, 6, len(billing_cycle_text))
worksheet.set_column(0, 7, len(days_text))
worksheet.set_column(0, 8, len(avg_daily_tmp_text))
worksheet.set_column(0, 9, len(avg_daily_useage_text))
worksheet.set_column(0, 10, len(basic_service_text))
worksheet.set_column(0, 11, len(theoretical_charge_text))
worksheet.set_column(0, 12, len(theoretical_savings_text))
worksheet.set_column(0, 13, len(price_per_kwh))
worksheet.set_column(0, 14, len(actual_price_per_kwh))
# worksheet.set_column(0, 12, len(theoretical_savings_total_text))
currency_format = workbook.add_format({'num_format': '[$$-409]#,##0.00'})
NYSEGBillFileDataPath = nyseg_csv_directory
if os.path.exists(NYSEGBillFileDataPath):
    os.remove(NYSEGBillFileDataPath)
nysegFile = open(NYSEGBillFileDataPath, "a")
file_index = 0
# Text file output
# File name, Billing Cycle, Billing Days, Month, Prior Excess Generation, Contribution To Grid, Useage, Billed Useage, Billed Amount, Avg Daily Tmp, Avg Daily Useage, Basic Service Charge, Theoretical Charge, Theoretical Savings, Price Per KWH
for m in (month):
    excess_gen = int(excess_generation[row-1])
    current_gen = int(current_generation[row-1])
    current_use = int(current_useage[row-1])
    net_billed = int(net_usage_billed[row-1])
    energy_charge = float(total_energy_charges[row-1])
    cycle = billing_cycle[row-1]
    dates = cycle.split(' - ')
    d1 = datetime.strptime(dates[0], "%m/%d/%y")
    d2 = datetime.strptime(dates[1], "%m/%d/%y")
    days = abs((d2 - d1).days) + 1
    avg_daily_tmp = average_daily_temp[row-1]
    avg_daily_use = int(average_daily_use[row-1])
    base_service_charge = basic_service_charges[row-1]
    theory_charges = theoretical_charges[row-1]
    theory_savings = theoretical_savings[row-1]
    ppkwh = price_per_kwh_est[row-1]
    actual_ppkwh = actual_pkwh[row-1]
    worksheet.write(row, 0, m)
    worksheet.write(row, 1, excess_gen, blue_format)
    worksheet.write(row, 2, current_gen, green_format)
    worksheet.write(row, 3, current_use, red_format)
    worksheet.write(row, 4, net_billed, yellow_format)
    worksheet.write(row, 5, energy_charge, currency_format)
    worksheet.write(row, 6, cycle)
    worksheet.write(row, 7, days)
    worksheet.write(row, 8, avg_daily_tmp)
    worksheet.write(row, 9, avg_daily_use)
    worksheet.write(row, 10, base_service_charge, currency_format)
    worksheet.write(row, 11, theory_charges, currency_format)
    worksheet.write(row, 12, theory_savings, currency_format)
    worksheet.write(row, 13, ppkwh, currency_format)
    worksheet.write(row, 14, actual_ppkwh, currency_format)
    
    row += 1
    if "n/a" != actual_ppkwh:
        actual_ppkwh = actual_ppkwh[1:]
    nysegData = [
        str(file_names[file_index]),
        str(cycle),
        str(days),
        str(m),
        str(excess_gen),
        str(current_gen),
        str(current_use),
        str(net_billed),
        str(energy_charge),
        str(avg_daily_tmp),
        str(avg_daily_use),
        str(base_service_charge),
        str(theory_charges),
        str(theory_savings),
        str(ppkwh),
        str(actual_ppkwh)
    ]
    nysegFile.write(','.join(nysegData) + '\n')
    file_index += 1
nysegFile.close()
worksheet.write(row+1, 11, sum(theoretical_charges), currency_format)
worksheet.write(row+1, 12, sum(theoretical_savings), currency_format)
chart = workbook.add_chart({'type': 'line'})
chart.set_style(48)
min_date = billing_cycle[0].split('-')[1].split('/')
max_date = billing_cycle[row-2].split('-')[1].split('/')
chart.set_x_axis({
    'name': 'KWH per Month',
    })
chart.add_series({'categories': '=Sheet1!$A2:$A'+ str(row), 'values': '=Sheet1!$B$2:$B$'+ str(row), 'line': {'color': 'blue'}, 'name': 'Prior Excess Generation' })
chart.add_series({'categories': '=Sheet1!$A2:$A'+ str(row), 'values': '=Sheet1!$C$2:$C$'+ str(row), 'line': {'color': 'green'}, 'name': 'Contribution' })
chart.add_series({'categories': '=Sheet1!$A2:$A'+ str(row), 'values': '=Sheet1!$D$2:$D$'+ str(row), 'line': {'color': 'red'}, 'name': 'Useage' })
chart.add_series({'categories': '=Sheet1!$A2:$A'+ str(row), 'values': '=Sheet1!$E$2:$E$'+ str(row), 'line': {'color': 'yellow'}, 'name': 'Useage Billed' })
chart_row = row + 4
worksheet.insert_chart('Q1', chart)
workbook.close()



