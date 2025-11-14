#! python3
"""autoReportGenerator.py - A program that reads a .csv file with data to be analysed, a
.json file with configuration information and a .xml file with settings and generates a
.docx report based on the information from those files. It also schedules the periodic
creation of the report and logs an entry in a log .csv file for every report generated."""

import os, csv, json, xmltodict, docx, datetime, time
from pathlib import Path

totalRows = 0

# Parsing 'sales.csv' and finding total processed lines
with open(Path('data_sources', 'sales.csv')) as csvFileObj:
    csvReader = csv.DictReader(csvFileObj)
    csvData = list(csvReader)
    totalRows += len(csvData) + 1

# Parsing 'reportConfig.json'
with open(Path('data_sources', 'reportConfig.json')) as jsonFileObj:
    jsonString = jsonFileObj.read()
    jsonData = json.loads(jsonString)

# Finding the number of processed lines in the json file    
with open(Path('data_sources', 'reportConfig.json')) as jsonFileObj:
    jsonLineList = jsonFileObj.readlines()    
    totalRows += len(jsonLineList)

# Parsing 'settings.xml'
with open(Path('data_sources', 'settings.xml')) as xmlFileObj:
    xmlString = xmlFileObj.read()
    xmlData = xmltodict.parse(xmlString)

# Finding the number of processed lines in the xml file
with open(Path('data_sources', 'settings.xml')) as xmlFileObj:
    xmlLineList = xmlFileObj.readlines()
    totalRows += len(xmlLineList)

# Creating the folder structure
reportFolder = xmlData['settings']['paths']['report_directory']
logFolder = xmlData['settings']['paths']['archive_directory']
os.makedirs(Path('output', reportFolder), exist_ok=True)
os.makedirs(Path('output', logFolder), exist_ok=True)

# Calculating summary statistics
unitsSold = 0
revenue = 0
summaryDict = {}
decimals = jsonData['decimal_places']
for row in range(len(csvData)):
    summaryDict.setdefault(csvData[row]['product'], {'units_sold': 0, 'revenue': 0})
    summaryDict[csvData[row]['product']]['units_sold'] += int(csvData[row]['units_sold'])
    summaryDict[csvData[row]['product']]['revenue'] += round(int(csvData[row]['units_sold']
                                                           )*float(csvData[row][
                                                               'unit_price']), decimals)
    unitsSold += int(csvData[row]['units_sold'])
    revenue += int(csvData[row]['units_sold'])*float(csvData[row]['unit_price'])
revenue = round(revenue, decimals)
avRevenuePerUnit = round(revenue/unitsSold, decimals)

# Generating the word file
while True:
    try:
        timestamp = time.time()
        folderName = datetime.datetime.now().strftime('%Y-%m-%d')
        reportName = 'Daily_Report_' + folderName + '.docx'
        os.makedirs(Path('output', reportFolder, folderName))
        doc = docx.Document()
        doc.add_heading(jsonData['report_title'], 0)
        doc.add_heading('Company Name: ' + jsonData['company_name'], 1)
        doc.add_heading('Author: ' + jsonData['author'], 1)
        doc.sections[0].footer.paragraphs[0].text = (jsonData['footer_text'])
        doc.sections[0].footer.add_paragraph(time.ctime(timestamp))
        doc.paragraphs[-1].runs[0].add_break(docx.enum.text.WD_BREAK.PAGE)
        table = doc.add_table(rows=len(summaryDict)+1, cols=3)
        table.rows[0].cells[0].text = 'Product'
        table.rows[0].cells[0].paragraphs[0].runs[0].bold = True
        table.rows[0].cells[1].text = 'Units Sold'
        table.rows[0].cells[1].paragraphs[0].runs[0].bold = True
        table.rows[0].cells[2].text = 'Revenue'
        table.rows[0].cells[2].paragraphs[0].runs[0].bold = True
        for i, (product, stats) in enumerate(summaryDict.items()):
            currency = jsonData['currency']
            table.rows[i+1].cells[0].text = product
            table.rows[i+1].cells[1].text = str(stats['units_sold'])
            if stats['units_sold'] >= jsonData['highlight_threshold']['high_sales']:
                table.rows[i+1].cells[1].paragraphs[0].runs[0].italic = True
            elif stats['units_sold'] < jsonData['highlight_threshold']['low_sales']:
                table.rows[i+1].cells[1].paragraphs[0].runs[0].underline = True
            table.rows[i+1].cells[2].text = str(stats['revenue']) + ' ' + currency
        doc.add_paragraph('\n\n')
        doc.add_paragraph('The total number of units sold was ' + str(unitsSold) +
                          ', the total revenue was ' + str(revenue) + ' ' + str(currency) +
                          ' and the average revenue per unit was ' + str(avRevenuePerUnit) +
                          ' ' + str(currency) + '.')
        doc.save(Path('output', reportFolder, folderName, reportName))
        duration = round(time.time() - timestamp, 3)

        # Generating/appending to the log csv file
        csvExists = 0
        if Path('output', logFolder, 'report_log.csv').exists():
            csvExists = 1     
        with open(Path('output', logFolder, 'report_log.csv'), 'a', newline='') as LogFile:
            csvWriter = csv.DictWriter(LogFile,['Timestamp', 'Filename', 'Rows Processed',
                                                   'Duration of Generation'])
            if csvExists == 0:
                csvWriter.writeheader()
            csvWriter.writerow({'Timestamp': time.ctime(timestamp), 'Filename': reportName,
                                'Rows Processed': str(totalRows), 'Duration of Generation':
                                str(duration)})
            
    except FileExistsError:
        pass
 
    # Scheduling runs
    if xmlData['settings']['schedule']['enabled'] == 'true':
        if xmlData['settings']['schedule']['run_every'] == '24h':
            time.sleep(86400)
        elif xmlData['settings']['schedule']['run_every'] == 'weekly':
            time.sleep(604800)
        else:
            break
    else:
        break  
