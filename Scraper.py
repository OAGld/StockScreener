import requests
from bs4 import BeautifulSoup

class Scrape_Data:

    def __init__(self, Comp_Name):
        self.Comp_Name = Comp_Name
        self.Profit_Margin = None
        self.RoA = None
        self.RoE = None
        self.Current_Ratio = None
        self.Shares_Outstanding = None
        self.List = []
        self.NumberOfDividends = None
        self.YearsOfDividends = None
        self.Payout_Ratio = None
        self.D_to_E_Ratio = None
        self.CurrentYield = None
        self.TrailingYield = None

        #Valuation Measures
        self.PE_Ratio = None
        self.MarketCap = None
        self.PB_Ratio = None
        Scrape_Yahoo_Finance.Scrape_Stats(self)
        Scrape_Yahoo_Finance.Scrape_Annual_Dividends(self)


class Scrape_Yahoo_Finance(Scrape_Data):
    def Scrape_Stats(self):
        try:
            URL = 'https://finance.yahoo.com/quote/' + self.Comp_Name + '/key-statistics?p=' + self.Comp_Name
            headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/96.0.4664.93 Safari/537.36'}

            #Parse website
            stats_page = requests.get(URL, headers=headers)
            stats = BeautifulSoup(stats_page.content, 'html.parser')

            #Find all tags containing the relevant info from the "Valuation Measures" table
            Valuation_Measures = stats.find('div', class_="Fl(start) smartphone_W(100%) W(50%)")
            Table_Divs = Valuation_Measures.find_all('div', class_="Pos(r) Mt(10px)")

            #Extract data from the "Valuation Measures" table 
            for a in Table_Divs:
                try:
                    table = a.find_all('table')
                    for b in table:
                        try:
                            row = b.find_all('tr')
                            for c in row:
                                try:
                                    Name = c.find('span')
                                    Value = c.find('td', class_="Fw(500) Ta(end) Pstart(10px) Miw(60px)")

                                    #Assign the values
                                    if ("Market Cap" in Name):
                                        self.MarketCap = float(Value.text.replace('%', ''))
                                    if ("Trailing P/E" in Name):
                                        self.PE_Ratio = float(Value.text.replace('%', ''))
                                    if ("Price/Book" in Name):
                                        self.PB_Ratio = float(Value.text.replace('%', ''))
                                except Exception as e:
                                    print(self.Comp_Name + ": ")
                                    print(e)
                        except Exception as e:
                            print(self.Comp_Name + ": ")
                            print(e)
                except Exception as e:
                    print(self.Comp_Name + ": ")
                    print(e)
            #Find all tags containing the relevant info from the "Financial Highlights" table
            Financial_Highlights = stats.find('div', class_="Fl(start) W(50%) smartphone_W(100%)")
            Table_Divs = Financial_Highlights.find_all('div', class_="Pos(r) Mt(10px)")

            #Extract data from the "Financial Highlights" table
            for a in Table_Divs:
                try:
                    table = a.find_all('table')
                    for b in table:
                        try:
                            row = b.find_all('tr')
                            for c in row:
                                try:
                                    Name = c.find('span')
                                    Value = c.find('td', class_="Fw(500) Ta(end) Pstart(10px) Miw(60px)")

                                    #Assign the values
                                    if ("Profit Margin" in Name):
                                        self.Profit_Margin = float(Value.text.replace('%', ''))
                                    if ("Return on Assets" in Name):
                                        self.RoA = float(Value.text.replace('%', ''))
                                    if ("Return on Equity" in Name):
                                        self.RoE = float(Value.text.replace('%', ''))
                                    if ("Current Ratio" in Name):
                                        self.Current_Ratio = float(Value.text.replace('%', ''))
                                    if ("Total Debt/Equity" in Name):
                                        self.D_to_E_Ratio = float(Value.text)
                                except Exception as e:
                                    print(self.Comp_Name + ": ")
                                    print(e)
                        except Exception as e:
                            print(self.Comp_Name + ": ")
                            print(e)
                except Exception as e:
                    print(self.Comp_Name + ": ")
                    print(e)


            #Find all tags containing the relevant info from the "Trading Information" table
            Trading_Information = stats.find('div', class_="Fl(end) W(50%) smartphone_W(100%)")
            Table_Divs = Trading_Information.find_all('div', class_="Pos(r) Mt(10px)")

            #Extract data from the "Trading Information" table
            for a in Table_Divs:
                try:
                    table = a.find_all('table')
                    for b in table:
                        try:
                            row = b.find_all('tr')
                            for c in row:
                                try:
                                    Name = c.find('span')
                                    Value = c.find('td', class_="Fw(500) Ta(end) Pstart(10px) Miw(60px)")

                                    #Assign the values
                                    if ("Shares Outstanding" in Name):
                                        self.Shares_Outstanding = float(Value.text.replace('M', ''))
                                    if ("Payout Ratio" in Name):
                                        self.Payout_Ratio = Value.text
                                    if ("Forward Annual Dividend Yield" in Name):
                                        self.CurrentYield = float(Value.text.replace('%', ''))
                                    if ("Trailing Annual Dividend Yield" in Name):
                                        self.TrailingYield = float(Value.text.replace('%', ''))

                                except Exception as e:
                                    print(self.Comp_Name + ": ")
                                    print(e)
                        except Exception as e:
                            print(self.Comp_Name + ": ")
                            print(e)
                except Exception as e:
                    print(self.Comp_Name + ": ")
                    print(e)


        except Exception as e:
            print("Unable to scrape data from: " + 'https://finance.yahoo.com/quote/' + self.Comp_Name + '/key-statistics?p=' + self.Comp_Name)
            print(e)


    def Scrape_Annual_Dividends(self):

        Prev_Year = None
        Counter = 0
        List = []

        try:
            URL = 'https://finance.yahoo.com/quote/' + self.Comp_Name + '/history?period1=1265328000&period2=1639180800&interval=capitalGain%7Cdiv%7Csplit&filter=div&frequency=1d&includeAdjustedClose=true'
            headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/96.0.4664.93 Safari/537.36'}

            #Parse website
            Historical_Data_page = requests.get(URL, headers=headers)

            #Find all tags containing the relevant info
            Historical_Data = BeautifulSoup(Historical_Data_page.content, 'html.parser')
            Historical_Data = Historical_Data.find('div', class_='Pb(10px) Ovx(a) W(100%)')
            Table = Historical_Data.find('tbody')
            row = Table.find_all('tr', class_='BdT Bdc($seperatorColor) Ta(end) Fz(s) Whs(nw)')

            #Collect the historical data and append to List
            for b in row:
                try:
                    Year = int(b.find('td', class_="Py(10px) Ta(start) Pend(5px)").text[8:]) #Strips the first 8 characters from the string, so only the year is left and converts it to an int
                    Value = float(b.find('strong').text)
                    List.append([])
                    List[Counter].append(Year)
                    List[Counter].append(Value)
                    Counter += 1
                except Exception as e:
                    print(self.Comp_Name + ": ")
                    print(e)


            Flag = 0
            Counter = 0
            self.NumberOfDividends = len(List)

            #Create an annualised list with all partial dividends combined for each year
            for x in range(len(List)-1):
                try:
                    Year = List[x][0]
                    Value += List[x][1]

                    if List[x+1][0] != Year:
                        self.List.append([])
                        self.List[Counter].append(Year)
                        self.List[Counter].append(Value)
                        Value = 0
                        Counter += 1
                except Exception as e:
                    print(self.Comp_Name + ": ")
                    print(e)
            try:
                self.YearsOfDividends = Counter
                #Remove data which may be incomplete or wrong from list
                self.List.pop(0)
                self.List.pop(len(self.List)-1)
            except Exception as e:
                print(self.Comp_Name + ": ")
                print(e)

        except Exception as e:
            print("Unable to scrape data from: " + 'https://finance.yahoo.com/quote/' + self.Comp_Name + '/history?period1=1097452800&period2=1596326400&interval=div%7Csplit&filter=div&frequency=1d')
            print(e)
