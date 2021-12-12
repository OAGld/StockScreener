from Scraper import Scrape_Data, Scrape_Yahoo_Finance
import configparser
from datetime import date

#Create a list for the Scrape_Data instances
Comp_Data = []

def main():
    Counter = 0
    Index = 0
    parser = configparser.ConfigParser()
    parser.read("StockCriteria.txt")
    TickerList = open(parser.get("Path", "Tickerlist"), "r").readlines()

    for x in TickerList:
        print(TickerList[Counter][:-1])
        Comp_Data.append(Scrape_Data(TickerList[Counter][:-1]))

        Counter += 1

    for a in range(len(Comp_Data)):
        Dividend = Calc_Dividend(a)

        if Comp_Data[a].Current_Ratio != None and Comp_Data[a].D_to_E_Ratio != None and Comp_Data[a].PE_Ratio != None and Comp_Data[a].PB_Ratio != None and Comp_Data[a].Shares_Outstanding != None and Comp_Data[a].CurrentYield != None:
            if (Comp_Data[a].Current_Ratio >= float(parser.get("criteria", "Current_Ratio"))) and (Comp_Data[a].D_to_E_Ratio/100 <= float(parser.get("criteria", "Debt_Equity_Ratio")) and Comp_Data[a].CurrentYield >= float(parser.get("criteria", "CurrentYield"))) and Comp_Data[a].Shares_Outstanding <= float(parser.get("criteria", "Shares_Outstanding")) and (Comp_Data[a].PE_Ratio*Comp_Data[a].PB_Ratio) <= float(parser.get("criteria", "PE-PB_ratio")):
                try:
                    CompaniesForAnalysis = open(parser.get("Path", "Save") + Comp_Data[a].Comp_Name[:-3] + ".txt", "w")
                    CompaniesForAnalysis.write("---------------  " + str(date.today()) + "  --------------\n")
                    CompaniesForAnalysis.write('Ticker: ' + Comp_Data[a].Comp_Name + "\n")
                    CompaniesForAnalysis.write('Profit Margin: ' + str(Comp_Data[a].Profit_Margin) + '%\n')
                    CompaniesForAnalysis.write('PE ratio: ' + str(Comp_Data[a].PE_Ratio) + '\n')
                    CompaniesForAnalysis.write('PB ratio: ' + str(Comp_Data[a].PB_Ratio) + '\n')
                    CompaniesForAnalysis.write('PB and PE product: ' + str((Comp_Data[a].PB_Ratio*Comp_Data[a].PE_Ratio)) + '\n')
                    CompaniesForAnalysis.write('Return on Assets: ' + str(Comp_Data[a].RoA) + '%\n')
                    CompaniesForAnalysis.write('Return on Equity: ' + str(Comp_Data[a].RoE) + '%\n')
                    CompaniesForAnalysis.write('Current Ratio: ' + str(Comp_Data[a].Current_Ratio) + '\n')
                    CompaniesForAnalysis.write('Debt-to-Equity Ratio: ' + str(round(float(Comp_Data[a].D_to_E_Ratio)/100, 2)) + '\n')
                    CompaniesForAnalysis.write('Shares Outstanding: ' + str(Comp_Data[a].Shares_Outstanding) + 'M' + '\n')
                    CompaniesForAnalysis.write('Payout Ratio: ' + str(Comp_Data[a].Payout_Ratio) + '\n')
                    CompaniesForAnalysis.write('Forward Annual Dividend Yield: ' + str(Comp_Data[a].CurrentYield) + '%\n')
                    CompaniesForAnalysis.write('Trailing Annual Dividend Yield: ' + str(Comp_Data[a].TrailingYield) + '%\n')
                    CompaniesForAnalysis.write('Years with dividends: ' + str(Comp_Data[a].YearsOfDividends) + '\n')
                    CompaniesForAnalysis.write('Average dividend growth: ' + Dividend.DividendGrowth + '\n\n')
                    CompaniesForAnalysis.write('Dividends\n')

                    #Insert all dividends
                    for b in Comp_Data[a].List:
                        CompaniesForAnalysis.write(str(b) + "\n")

                except Exception as e:
                    print("Problems inserting values into text document")
                    print(e)

        else:
            print(Comp_Data[a].Comp_Name + ": Not alle values are present")                 
            Index += 1

class Calc_Dividend:

    def __init__(self, Instance):
        self.DividendGrowth = None
        self.Instance = Instance
        Calc_Dividend.Dividend_Growth(self)

    def Dividend_Growth(self):
        try:
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
        except Exception as e:
            print("Problems calculating dividend growth.")
            print(e)

main()
