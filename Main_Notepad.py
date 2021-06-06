from Scraper import Scrape_Data, Scrape_Yahoo_Finance
import openpyxl
from openpyxl import load_workbook
import configparser

#Open the TickerList textfile
TickerList = open(r"tickerlist.txt", "r").readlines()

#Create a list for the Scrape_Data instances
Comp_Data = []

#Create a new texfile document to write new data to
#CompaniesForAnalysis = open(r"CompaniesForAnalysis", "a")


def main():
    Counter = 0
    Index = 0
    parser = configparser.ConfigParser()
    parser.read("StockCriteria.txt")

    for x in TickerList:
        Comp_Data.append(Scrape_Data(TickerList[Counter][:-1]))
        Counter += 1

    for a in range(len(Comp_Data)):
        Dividend = Calc_Dividend(a)

        if Comp_Data[a].Current_Ratio != None and Comp_Data[a].TrailingYield != None and Comp_Data[a].CurrentYield != None and Comp_Data[a].Profit_Margin != None and Comp_Data[a].RoA != None and Comp_Data[a].RoE != None:
            if (Comp_Data[a].Current_Ratio >= float(parser.get("criteria", "Current_Ratio"))) and (Comp_Data[a].TrailingYield >= float(parser.get("criteria", "TrailingYield")) or Comp_Data[a].CurrentYield >= float(parser.get("criteria", "CurrentYield"))) and Comp_Data[a].Profit_Margin >= float(parser.get("criteria", "Profit_Margin")) and Comp_Data[a].RoA >= float(parser.get("criteria", "RoA")) and Comp_Data[a].RoE >= float(parser.get("criteria", "RoE")):
                try:
                    CompaniesForAnalysis = open(Comp_Data[a].Comp_Name[:-4] + ".txt", "a+")
                    CompaniesForAnalysis.write('Ticker: ' + Comp_Data[a].Comp_Name + "-------------------------\n")
                    CompaniesForAnalysis.write('Profit Margin: ' + str(Comp_Data[a].Profit_Margin) + '%\n')
                    CompaniesForAnalysis.write('Return on Assets: ' + str(Comp_Data[a].RoA) + '%\n')
                    CompaniesForAnalysis.write('Return on Equity: ' + str(Comp_Data[a].RoE) + '%\n')
                    CompaniesForAnalysis.write('Current Ratio: ' + str(Comp_Data[a].Current_Ratio) + '\n')
                    CompaniesForAnalysis.write('Debt-to-Equity Ratio: ' + str(round(float(Comp_Data[a].D_to_E_Ratio)/100, 2)) + '\n')
                    CompaniesForAnalysis.write('Shares Outstanding: ' + Comp_Data[a].Shares_Outstanding + '\n')
                    CompaniesForAnalysis.write('Payout Ratio: ' + Comp_Data[a].Payout_Ratio + '\n')
                    CompaniesForAnalysis.write('Forward Annual Dividend Yield: ' + str(Comp_Data[a].CurrentYield) + '%\n')
                    CompaniesForAnalysis.write('Trailing Annual Dividend Yield: ' + str(Comp_Data[a].TrailingYield) + '%\n')
                    CompaniesForAnalysis.write('Average dividend growth: ' + Dividend.DividendGrowth + '\n')
                except Exception as e:
                    print("Problems inserting values into text document")
                    print(e)
                
            Index += 1

class Calc_Dividend:

    def __init__(self, Instance):
        self.DividendGrowth = None
        self.Instance = Instance
        Calc_Dividend.Dividend_Growth(self)

    def Dividend_Growth(self):
    #try:
        if len(Comp_Data[self.Instance].List) >= 5:
            Sum = 0
            Annual_Growth_Rate = []
            for a in range(len(Comp_Data[self.Instance].List)):
                if a > 0:
                    Annual_Growth_Rate.append((float(Comp_Data[self.Instance].List[a-1][1])/float(Comp_Data[self.Instance].List[a][1]))-1)

            for a in range(len(Annual_Growth_Rate)):
                Sum += Annual_Growth_Rate[a]

            self.DividendGrowth = str(round((Sum/float(len(Annual_Growth_Rate))*100), 2)) + '%'
        else:
            self.DividendGrowth = "Too few"
    """except Exception as e:
        print("Problems calculating dividend growth.")
        print(e)"""

main()
