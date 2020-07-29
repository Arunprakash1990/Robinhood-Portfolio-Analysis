import getpass
import requests
import numpy as np
import pandas as pd
from pyrh import Robinhood
import xlsxwriter

def login():
    username = input('Username :: ')
    password = getpass.getpass('Password :: ')
    print("Check your email for 2 factor Authentication Code:")
    rh.login(username, password)    

def get_order_History():
    orderHistory = rh.order_history()
    orderHistoryList = orderHistory['results']
    while True:
        if(orderHistory['next'] is None):
            break
        orderHistory = rh.get_url(orderHistory['next'])
        orderHistoryList.extend(orderHistory['results'])
    return orderHistoryList

#print('Total Orders in your RobinHood Account:',len(get_order_History()))

def get_all_orders_sortedByDate(orderHistoryList):
    stockDict = {}
    allStockList = []
    for order in orderHistoryList:
        if(order['state'] != "cancelled" and order['state'] != "confirmed" and order['state'] != "rejected"):
            instrumentResponse = requests.get(order["instrument"])
            instrumentJSON = instrumentResponse.json()
            # print(instrumentJSON)

            stockDict['name'] = instrumentJSON["simple_name"]
            stockDict['ticker'] = instrumentJSON["symbol"]
            stockDict['BuyingPricePerShare'] = order["average_price"]
            stockDict['totalBuyingPrice'] = order["executed_notional"]["amount"]
            stockDict['quantity'] = order["cumulative_quantity"]
            stockDict['date'] = order["executions"][0]["settlement_date"]
            stockDict['tranType'] = order["side"]
            allStockList.append(stockDict.copy())
    return sorted(allStockList,key=lambda k:k['date'])

def split_BuyAndSellList(sortedOrders):
    buyList = []
    sellList = []
    tickerList =[]
    stockMap ={}
    for row in sortedOrders:
        if row['tranType'] == 'buy':
            if row['ticker'] not in stockMap:
                stockMap[row['ticker']] = row
                tickerList.append(row['ticker'])
            else:
                existingRow = stockMap[row['ticker']]
                newRow ={}
                newRow['name'] = str(row['name'])
                newRow['ticker'] = str(row['ticker'])
                newRow['tranType'] = str(row['tranType'])
                newRow['date'] = str(row['date'])
                newRow['quantity'] = float(row['quantity']) + float(existingRow['quantity'])
                newRow['totalBuyingPrice'] = float(row['totalBuyingPrice']) + float(existingRow['totalBuyingPrice'])
                newRow['BuyingPricePerShare'] = float(newRow['totalBuyingPrice'])/float(newRow['quantity'])
                stockMap[row['ticker']] = newRow
        else:
            if(row['ticker'] in stockMap):
                alreadyBoughtRow = stockMap.pop(row['ticker'])
            if((float(alreadyBoughtRow['quantity']) - float(row['quantity'])) != 0.0):
                alreadyBoughtRow['quantity'] = float(alreadyBoughtRow['quantity']) - float(row['quantity'])
                alreadyBoughtRow['totalBuyingPrice'] = float(alreadyBoughtRow['BuyingPricePerShare'])*float(alreadyBoughtRow['quantity'])
                stockMap[row['ticker']] = alreadyBoughtRow
            soldRow = alreadyBoughtRow.copy()
            soldRow['SoldDate'] = str(row['date'])
            soldRow['tranType'] = 'Sold'
            soldRow['sellingQuantity'] = float(row['quantity'])
            soldRow['totalSellingPrice'] = float(row['totalBuyingPrice'])
            soldRow['SellingPricePerShare'] = float(row['BuyingPricePerShare'])
            sellList.append(soldRow)
    buyList = stockMap.values()
    return buyList,sellList,tickerList

def get_curr_marketPrice(tickerList):
    currPriceList = rh.quotes_data(set(tickerList))
    #currPriceList
    currDict = {}
    currStockList = []
    for curr in currPriceList:
        if curr is not None:
            currDict['ticker'] = curr['symbol']
            currDict['currPrice'] = curr['last_trade_price']
            currStockList.append(currDict.copy())
    return currStockList

def fillnaByTypes(df):
    for col in df:
        #get dtype for column
        dt = df[col].dtype 
        #check if it is a number
        if dt == int or dt == float:
            df[col].fillna(0)
        else:
            df[col].fillna("Unknown")
    return df;

def prepare_data_frames(buyList,currPriceList):
    df1 = pd.DataFrame(buyList)
    convert_type = {'name' : str,
                    'ticker' : str,
                    'BuyingPricePerShare' : float,
                    'totalBuyingPrice' : float,
                    'quantity' : float,
                    'date' : str,
                    'tranType' : str
                   }
    df1 = df1.astype(convert_type)
    df1 = df1.rename(columns={  'name' : 'Name',
                                'ticker' : 'Ticker',
                                'BuyingPricePerShare' : 'Purchase Price (Per Share)',
                                'totalBuyingPrice' : 'Total Purchase Price',
                                'quantity' : 'Purchase Quantity',
                                'date' : 'Purchase Date',
                                'tranType' : 'Transaction Type'
                   })
    #print(df1.dtypes)

    df2 = pd.DataFrame(currPriceList)
    convert_type = {'ticker' : str,
                    'currPrice' : float
                   }
    df2 = df2.astype(convert_type)
    df2 = df2.rename(columns={  'ticker' : 'Ticker',
                                'currPrice' : 'Current Market Price (Per Share)'
                   })
    #print(df2.dtypes)
    
    df3 = pd.read_csv('companylist.csv',usecols=['Symbol','Sector','industry'])
    df3.columns = ['Ticker','Sector','Industry']
    convert_type = {'Ticker' : str,
                    'Sector' : str,
                    'Industry' : str
                   }
    
    df3 = df3.astype(convert_type)
    df3 = df3.drop_duplicates()
    #print(df3.dtypes)
    
    portfolio = df1.merge(df2,on='Ticker',how='left').merge(df3,on='Ticker',how='left')
    portfolio = portfolio.sort_values('Purchase Date',ascending=False)
    portfolio.reset_index(inplace=True,drop=True)
    portfolio = fillnaByTypes(portfolio)
    return portfolio

def re_order_columns(portfolio):
    return portfolio[['Name',
                     'Ticker',
                     'Purchase Date',
                     'Transaction Type',
                     'Purchase Quantity',
                     'Purchase Price (Per Share)',
                     'Current Market Price (Per Share)',
                     'Total Purchase Price',
                     'Total Market Price',
                     'Sector',
                     'Industry',
                     'Profit/Loss',
                     'Profit/Loss Percentage',
                     'Portfolio Diversity'
                     ]].sort_values('Profit/Loss Percentage')

def create_calculated_fields(portfolio):
    portfolio['Total Market Price'] = round(portfolio['Current Market Price (Per Share)'] * portfolio['Purchase Quantity'],2)
    portfolio["Profit/Loss"] = round(portfolio["Total Market Price"] - portfolio["Total Purchase Price"],2)
    portfolio["Profit/Loss Percentage"] = round((portfolio["Profit/Loss"]/portfolio["Total Purchase Price"])*100,2)
    portfolio["Portfolio Diversity"] = round((portfolio['Total Market Price']/portfolio['Total Market Price'].sum())*100,2)

    #portfolio["Profit/Loss Percentage"] = portfolio["Profit/Loss Percentage"].astype(str)
    #portfolio["Profit/Loss Percentage"] = portfolio["Profit/Loss Percentage"]  + "%"
    #portfolio["Portfolio Diversity"] = portfolio["Portfolio Diversity"].astype(str)
    #portfolio["Portfolio Diversity"] = portfolio["Portfolio Diversity"]  + "%"

    portfolio["Purchase Price (Per Share)"] = portfolio["Purchase Price (Per Share)"].round(2)
    portfolio["Total Market Price"] = portfolio["Total Market Price"].round(2)
    return re_order_columns(portfolio)

def get_Sector_and_Industry_Analysis(finalSummary):
    withIndustryColumns = [
                     'Sector',
                     'Industry',
                     'Total Purchase Price',
                     'Total Market Price',
                     'Profit/Loss'
                     ]
    industryAnalysis = finalSummary[withIndustryColumns].groupby(['Industry'],as_index=False).agg({'Total Purchase Price': np.sum,'Total Market Price': np.sum,'Profit/Loss': np.sum})  
    sectorAnalysis = finalSummary[withIndustryColumns].groupby(['Sector'],as_index=False).agg({'Total Purchase Price': np.sum,'Total Market Price': np.sum,'Profit/Loss': np.sum})  
    sectorAnalysis["Profit/Loss Percentage"] = round((sectorAnalysis["Profit/Loss"]/sectorAnalysis["Total Purchase Price"])*100,2)
    sectorAnalysis["Portfolio Diversity"] = round((sectorAnalysis['Total Market Price']/sectorAnalysis['Total Market Price'].sum())*100,2)
    industryAnalysis["Profit/Loss Percentage"] = round((industryAnalysis["Profit/Loss"]/industryAnalysis["Total Purchase Price"])*100,2)
    industryAnalysis["Portfolio Diversity"] = round((industryAnalysis['Total Market Price']/industryAnalysis['Total Market Price'].sum())*100,2)
    dropColumns =['Total Purchase Price','Total Market Price','Profit/Loss']
    sectorSummary = sectorAnalysis.drop(dropColumns,axis=1)
    industrySummary = industryAnalysis.drop(dropColumns,axis=1)
    
    
    return sectorSummary.sort_values('Profit/Loss Percentage'),industrySummary.sort_values('Profit/Loss Percentage')

def process_data_from_robinHood():
    orderHistoryList              =   get_order_History()
    sortedOrders                  =   get_all_orders_sortedByDate(orderHistoryList)
    buyList,sellList,tickerList   =   split_BuyAndSellList(sortedOrders)
    currPriceList                 =   get_curr_marketPrice(tickerList)
    portfolio                     =   prepare_data_frames(buyList,currPriceList)
    finalSummary                  =   create_calculated_fields(portfolio)
    sectorSummary,industrySummary =   get_Sector_and_Industry_Analysis(finalSummary)
    return finalSummary,sectorSummary,industrySummary

def write_to_Excel(finalSummary,sectorSummary,industrySummary):
    writer = pd.ExcelWriter('stock_Analysis_RobinHood.xlsx', engine='xlsxwriter')
    finalSummary.to_excel(writer, sheet_name='Final Summary',index=0)
    sectorSummary
    sectorSummary.to_excel(writer, sheet_name='Sector Summary',index=0)
    industrySummary.to_excel(writer, sheet_name='Industry Summary',index=0)
    workbook = writer.book
    
    formatPercentage = workbook.add_format({'num_format': '0.00\%'})
    
    sectorXlsx = writer.sheets['Sector Summary']
    sectorXlsx.set_column('B:C', None, formatPercentage)
    column_chart1 = workbook.add_chart({'type': 'column'})
    column_chart1.add_series({
    'name':     ['Sector Summary', 0, 1, 0, 1],    
    'values':     ['Sector Summary', 1, 1, sectorSummary.shape[0]-1, 1],
    'categories':     ['Sector Summary', 1, 0, sectorSummary.shape[0]-1, 0],
    })
    line_chart2 = workbook.add_chart({'type': 'line'})
    line_chart2.add_series({
    'name':     ['Sector Summary', 0, 2, 0, 2],    
    'values':     ['Sector Summary', 1, 2, sectorSummary.shape[0]-1, 2],
    'categories':     ['Sector Summary', 1, 0, sectorSummary.shape[0]-1, 0],
    })
    column_chart1.combine(line_chart2)
    column_chart1.set_title({ 'name': 'Sector Analysis of Portfolio Diversity and P/L %'})
    column_chart1.set_x_axis({'name': 'Sector','label_position': 'low'})
    column_chart1.set_y_axis({'name': 'Percentage %'})
    column_chart1.set_size({'width': 720, 'height': 576})

    sectorXlsx.insert_chart('F2', column_chart1)
    
    
    industryXlsx = writer.sheets['Industry Summary']
    industryXlsx.set_column('B:C', None, formatPercentage)
    column_chart1 = workbook.add_chart({'type': 'bar'})
    column_chart1.add_series({
    'name':     ['Industry Summary', 0, 1, 0, 1],    
    'values':     ['Industry Summary', 1, 1, industrySummary.shape[0]-1, 1],
    'categories':     ['Industry Summary', 1, 0, industrySummary.shape[0]-1, 0],
    })
    #line_chart2 = workbook.add_chart({'type': 'line'})
    #line_chart2.add_series({
    #'name':     ['Industry Summary', 0, 2, 0, 2],    
    #'values':     ['Industry Summary', 1, 2, industrySummary.shape[0]-1, 2],
    #'categories':     ['Industry Summary', 1, 0, industrySummary.shape[0]-1, 0],
    #})
    #column_chart1.combine(line_chart2)
    column_chart1.set_title({ 'name': 'Industry Analysis of Portfolio Diversity and P/L %'})
    column_chart1.set_x_axis({'name': 'Percentage %','label_position': 'low'})
    column_chart1.set_y_axis({'name': 'Industry','label_position': 'low'})
    column_chart1.set_size({'width': 1024, 'height': 768})

    industryXlsx.insert_chart('F2', column_chart1)
    
    writer.save()

if __name__ == "__main__":
    rh = Robinhood()
    login()
    finalSummary,sectorSummary,industrySummary = process_data_from_robinHood()
    write_to_Excel(finalSummary,sectorSummary,industrySummary)
    print("Check for an Excel file in the Project Directory")
    

    
