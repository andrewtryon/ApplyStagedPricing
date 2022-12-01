from dotenv import load_dotenv
load_dotenv()
import os
import json
import pyodbc
import pandas as pd
import subprocess
import requests
import time
import datetime
from datetime import date

def pullAccessTable(conn_str, table):
    cnxn = pyodbc.connect(conn_str)
    sql = """
    SELECT * FROM <TableName>
    """.replace('<TableName>',table)
    df = pd.read_sql(sql,cnxn)
    return df

def grabSagePricing(row):
    if row['PromoMAP'] == '':
        row['PromoMAP'] = 0
    if row['NewMAP'] == '':
        row['NewMAP'] = 0
    if row['RevertMAP'] == '':
        row['RevertMAP'] = 0
    if row['Status'] == '2-Regular Pricing':
        row['MSRP'] = row['NewMSRP']
        row['Sale'] = row['NewSale']
        row['Cost'] = row['NewCost']
        row['MAP'] = row['NewMAP']
        row['VendorEffectiveDate'] = pd.to_datetime(row['EffectiveDate'].strftime('%m/%d/%Y'))
    elif row['Status'] == '1-Expiring Promo':
        row['MSRP'] = row['RevertMSRP']
        row['Sale'] = row['RevertSale']
        row['Cost'] = row['RevertCost']
        row['MAP'] = row['RevertMAP']
        row['VendorEffectiveDate'] = pd.to_datetime(datetime.date.today().strftime('%m/%d/%Y'))
    elif row['Status'] == '3-On Promo':
        row['MSRP'] = row['PromoMSRP']
        row['Sale'] = row['PromoSale']
        row['Cost'] = row['PromoCost']
        row['MAP'] = row['PromoMAP']
        row['VendorEffectiveDate'] = pd.to_datetime(row['PromoStartDate'].strftime('%m/%d/%Y'))
    return row   

def calcMargins(row):
    if row['Cost'] == 0:
        row['VendorDiscount'] = 0
        row['SaleMargin'] = 0
    else:
        if row['MSRP'] == 0:
            row['VendorDiscount'] = 0
        else:
            row['VendorDiscount'] = (row['MSRP'] - row['Cost']) / row['MSRP']
            row['VendorDiscount'] = round(row['VendorDiscount'], 2)  * 100
        if row['Sale'] == 0:
            row['SaleMargin'] = 0
        else:            
            row['SaleMargin'] = (row['Sale'] - row['Cost']) / row['Sale']
            row['SaleMargin'] = round(row['SaleMargin'], 2) * 100
    return row         

def runAccessQuery(filepath, queryName):
    #Run Access Query Update 
    cmmd = r'"C:\Program Files (x86)\Microsoft Office\root\Office16\MSACCESS.EXE" "'
    cmmd = cmmd + filepath + r'" /x '
    cmmd = cmmd + queryName
    print(cmmd)
    process = subprocess.Popen(cmmd, shell=True, stdout=subprocess.PIPE)
    process.wait()    

def makeWrikeTask (title = "New Pricing Task", description = "No Description Provided", status = "Active", assignees = "KUAAY4PZ", folderid = "IEAAJKV3I4JBAOZD"):
    url = "https://www.wrike.com/api/v4/folders/" + folderid + "/tasks"
    querystring = {
        'title':title,
        'description':description,
        'status':status,
        'responsibles':assignees
        } 
    headers = {
        'Authorization': 'bearer TOKEN'.replace('TOKEN',os.environ.get(r"WRIKE_TOKEN"))
        }        
    response = requests.request("POST", url, headers=headers, params=querystring)
    return response

def attachWrikeTask (attachmentpath, taskid):
    url = "https://www.wrike.com/api/v4/tasks/" + taskid + "/attachments"
    headers = {
        'Authorization': 'bearer TOKEN'.replace('TOKEN',os.environ.get(r"WRIKE_TOKEN"))
    }

    files = {
        'X-File-Name': (attachmentpath, open(attachmentpath, 'rb')),
    }

    response = requests.post(url, headers=headers, files=files)
    return response        

def markWrikeTaskComplete (taskid):
    url = "https://www.wrike.com/api/v4/tasks/" + taskid + "/"
    querystring = {
        'status':'Completed'
        }     
    headers = {
        'Authorization': 'bearer TOKEN'.replace('TOKEN',os.environ.get(r"WRIKE_TOKEN"))
    }

    response = requests.request("PUT", url, headers=headers, params=querystring)
    return response    

if __name__ == '__main__':
    #Making Files and Tasks
    #KUACOUUA - Andrew
    #KUAEL7RV - Anthony
    #KUAAY4PZ - Kris
    #KUAAZIJV - Travis
    #KUALCDZR - Josiah
    assignees = '[KUACOUUA,KUAEL7RV,KUAAY4PZ,KUAAZIJV,KUALCDZR]'
    folderid = 'IEAAJKV3I4KM3YOP'    #this is the pricing folder ... had to use postman to figure these out ;)    
    try:
        conn_str = (
            r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
            r'DBQ=\\FOT00WEB\Alt Team\Qarl\Maintenance Tasks\LocalTables\MaintenanceLocalTables.accdb;'
            )
        #Pull and Merge Promo Tables
        promoLineItems = pullAccessTable(conn_str,'Pricing_Promo_LineItems')
        promoHeaders = pullAccessTable(conn_str,'Pricing_Promo_Headers')
        promoHeaders = promoHeaders.drop(['Status'], axis=1)
        promoDF = promoHeaders.merge(promoLineItems, how='outer', left_on='Promo', right_on='PromoName')
        logf2 = open(r"\\FOT00WEB\Alt Team\Qarl\Maintenance Tasks\StagedPricingLog\StagedPricingError2.log", "w")
        logf2.write("logf2")
        #Pull and Merge Regular Changes to Promo
        regularLineItems = pullAccessTable(conn_str,'Pricing_Regular_LineItems')
        pricingDF = promoDF.merge(regularLineItems, how='outer', left_on=['ItemCode','VendorNo'], right_on=['ItemCode','VendorNo'])
        pricingDF= pricingDF.drop(['PromoName'], axis=1)

        #Reintialize Status
        pricingDF['Status'] = ''

        #Adding Price Code
        pricingDF['PriceCode'] = 'REGULAR'        
        
        #Replace accidental line feeds in memo description
        pricingDF['MemoDescription'] = pricingDF['MemoDescription'].replace('\n','|', regex=True)

        #Today's date
        today = pd.to_datetime(datetime.date.today().strftime('%m/%d/%Y'))
        pricingDF['today'] = date.today().strftime('%m/%d/%Y')
        pricingDF['today'] = pd.to_datetime(pricingDF.today)

        #Expiring Promo
        mask = (pricingDF['PromoEndDate'] < pricingDF['today'])
        pricingDF.loc[mask,'Status'] = '1-Expiring Promo'
        pricingDF.loc[mask,'PriceCode'] = ''

        #Off Promo Regular Pricing
        mask = (pricingDF['EffectiveDate'] <= pricingDF['today'])
        pricingDF.loc[mask,'Status'] = '2-Regular Pricing'    

        #on Promo
        mask = (pricingDF['PromoStartDate'] <= pricingDF['today']) & (pricingDF['PromoEndDate'] >= pricingDF['today'])
        pricingDF.loc[mask,'Status'] = '3-On Promo'
        pricingDF.loc[mask,'PriceCode'] = 'PROMO'

        #Removing those needing no action
        pricingDF = pricingDF.loc[pricingDF['Status'] !='']

        #adding the sage specific columns
        sagecols = list(pricingDF.columns) + ['MSRP','Sale','Cost','MAP','VendorDiscount','SaleMargin','VendorEffectiveDate']
        pricingDF = pricingDF.reindex(columns=sagecols)

        #Add date col (will be overwritten if another date should be there)
        pricingDF['VendorEffectiveDate'] = today
        pricingDF['VendorEffectiveDate'] = pricingDF['VendorEffectiveDate'].dt.strftime('%m/%d/%Y')
        pricingDF = pricingDF.sort_values(by=['Status','VendorEffectiveDate']).reset_index(drop=True)

        #Move pricing and dates based upon status
        pricingDF = pricingDF.apply(grabSagePricing, axis=1)

        #Intialize Sage Margin Cols and calculate
        pricingDF['VendorDiscount'] = 0
        pricingDF['SaleMargin'] = 0
        pricingDF = pricingDF.apply(calcMargins, axis=1)

        #Round Sage Cols for sage
        pricingDF = pricingDF.round({'MSRP': 2, 'Sale': 2, 'Cost': 2, 'MAP': 2})

        #Make Auto Pricing VI csv
        csvfilepath = r'\\FOT00WEB\Alt Team\Qarl\Automatic VI Jobs\Maintenance\CSVs\AA_STAGEPRICES_VIWI5T_VIWI5U.csv'
        pricingDF.to_csv(csvfilepath, index=False, header=False, date_format='%m/%d/%Y', columns = ['ItemCode','VendorNo','MSRP','Sale','Cost','MAP','VendorDiscount','SaleMargin','VendorEffectiveDate','PriceCode'])
        csvfilepath = r'\\FOT00WEB\Alt Team\Qarl\Automatic VI Jobs\Maintenance\CSVs\StagedPricingBackupUpload.csv'
        pricingDF.to_csv(csvfilepath, index=False, header=True, date_format='%m/%d/%Y', columns = ['ItemCode','VendorNo','MSRP','Sale','Cost','MAP','VendorDiscount','SaleMargin','VendorEffectiveDate','PriceCode'])    

        CPU_cnxn = pyodbc.connect(conn_str)
        cur = CPU_cnxn.cursor()
        csvfilepath = r'\\FOT00WEB\Alt Team\Qarl\Automatic VI Jobs\Maintenance\CSVs'
        #insert sql
        print('Saving Pricing Changes')
        spc_sql = r"""
        INSERT INTO Pricing_MasterTable_RollingBackup (ItemCode,VendorNo,MSRP,Sale,Cost,MAP,VendorDiscount,SaleMargin,VendorEffectiveDate)
        SELECT ItemCode,VendorNo,MSRP,Sale,Cost,MAP,VendorDiscount,SaleMargin,VendorEffectiveDate
        FROM [text;HDR=Yes;FMT=Delimited(,);Database=path].StagedPricingBackupUpload.csv
        """.replace('path',csvfilepath)
        cur.execute(spc_sql)
        CPU_cnxn.commit()

        #Making Memo CSV
        pricingDF['AutoDisplay'] = 'Y'
        csvfilepath = r'\\FOT00WEB\Alt Team\Qarl\Automatic VI Jobs\Maintenance\CSVs\MasterMemoFile.csv'
        pricingDF.loc[pricingDF['Promo'] != ''].to_csv(csvfilepath, sep = '^', index=False, header=False, date_format='%m/%d/%Y', columns = ['ItemCode','Promo','MemoTitle','StartDate','EndDate','AutoDisplay','MemoDescription','File'])

        #Making Promo on and off CSV 
        csvfilepath = r'\\FOT00WEB\Alt Team\Qarl\Automatic VI Jobs\Maintenance\CSVs\PromoPriceCode.csv'
        pricingDF.loc[(pricingDF['Status'] == '1-Expiring Promo') | (pricingDF['Status'] == '3-On Promo')].to_csv(csvfilepath, date_format='%m/%d/%Y', index=False, header=False, columns = ['ItemCode','PriceCode','VendorEffectiveDate'])

        #Auto VI Price Changes
        print('VIing Stage Price Changes to CI_Item')
        p = subprocess.Popen('Auto_StagedPricing1_VIWI5T.bat', cwd=r"Y:\Qarl\Automatic VI Jobs\Maintenance", shell=True)
        stdout, stderr = p.communicate()
        p.wait()
        print('Sage VI Complete!')

        time.sleep(9)

        #Auto VI .... 
        print('VIing Stage Cost Changes to Vendor Cost levels')
        p = subprocess.Popen('Auto_StagedPricing2_VIWI5U.bat', cwd=r"Y:\Qarl\Automatic VI Jobs\Maintenance", shell=True)
        stdout, stderr = p.communicate()
        p.wait()
        print('Sage VI Complete!')  

        time.sleep(9)        

        #Auto VI .... 
        print('VIing Memo for Promos')
        p = subprocess.Popen('Auto_Memo1_VIWI4B.bat', cwd=r"Y:\Qarl\Automatic VI Jobs\Maintenance", shell=True)
        stdout, stderr = p.communicate()
        p.wait()
        print('Sage VI Complete!')

        time.sleep(9)   

        #Auto VI .... 
        print('VIing Promo On and Promo Off Price Codes for Promos')
        p = subprocess.Popen('Auto_PROMOCODES_VIWI7F.bat', cwd=r"Y:\Qarl\Automatic VI Jobs\Maintenance", shell=True)
        stdout, stderr = p.communicate()
        p.wait()
        print('Sage VI Complete!')            

        #Updating Access Tables
        maintFilepath = r'\\FOT00WEB\Alt Team\Qarl\Maintenance Tasks\LocalTables\StagedPricingJiggler.accdb'
        runAccessQuery(maintFilepath,"PricingStatusJiggler")
          
        if pricingDF.shape[0] != 0:
            description = date.today().strftime('%m/%d/%y')  + ' Successfully Pushed Prices' + '\n' + r'If you cannot find the pushed file in the task (too big too attach) look here: \\FOT00WEB\Alt Team\Qarl\Maintenance Tasks\StagedPricingLog\StagedPricingFile.xlsx'
            wriketitle = date.today().strftime('%m/%d/%y')  + ' Successfully Pushed Prices' 
        else:
            description = date.today().strftime('%m/%d/%y')  + ' No Price Changes Today!'
            wriketitle = date.today().strftime('%m/%d/%y')  + ' No Price Changes Today!'
    except Exception as e:     # most generic exception you can catch
        print(e)
        wriketitle = date.today().strftime('%m/%d/%y')  + ' CRAP Stage Pricing Failure!'
        logf = open(r"\\FOT00WEB\Alt Team\Qarl\Maintenance Tasks\StagedPricingLog\StagedPricingError.log", "w")
        logf.write("Error :( {0}\n".format(str(e)))
        description = r"""
        Pricing File Failed During execution- Attempted to attach the panda dataframe as it was when it errored"
        """
        description = description + '\n' + r'Error should be below or here "\\FOT00WEB\Alt Team\Qarl\Maintenance Tasks\StagedPricingLog\StagedPricingError.log"' + '\n' + "Error :( {0}\n".format(str(e))
    finally:
        response = makeWrikeTask(title = wriketitle, description = description, assignees = assignees, folderid = folderid)
        response_dict = json.loads(response.text)
        print('wrike task made!')
        taskid = response_dict['data'][0]['id']
        print('file attached!')    
        if pricingDF.shape[0] != 0:
            filetoattachpath = r"\\FOT00WEB\Alt Team\Qarl\Maintenance Tasks\StagedPricingLog\StagedPricingFile.xlsx"
            pricingDF.to_excel(filetoattachpath)     
            print('Attaching file')
            attachWrikeTask(attachmentpath = filetoattachpath, taskid = taskid)               
        markWrikeTaskComplete (taskid)