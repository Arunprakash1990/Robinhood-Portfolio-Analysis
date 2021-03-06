# Robinhood Trades to Excel with Additional Analysis

This project connects to your Robinhood Account using pyrh (the unofficial API for Robinhood) and segregate the data based on Industry and Sector listed in NASDAQ,NYSE stock Exchange to give additional insights on your portfolio. This Python application exports your Robinhood trades to a .xlsx file with additional Analysis:
- Portfolio Diversity Classification 
- Profit/Loss % for each transaction
- Sector Summary Classification
- Industry Summary Classification



pyrh: [Robinhood library](https://github.com/robinhood-unofficial/pyrh) created by Rohan Pai. Works on Python 3.5+. Install Python3 before running this application.

## Instructions
- Open Terminal
- Clone this Repo
- Install all the dependencies related to the Project using requirements.txt. pip3 install if you have both python 2 and 3 installed in your machine.
- Run the Application
- Look for an Excel file named "stock_Analysis_RobinHood.xlsx" on your project directory which will contain all the above analysis
- If you are comfortable in python code and want to see everything on the fly, fork the iPython notebook and customize the code to your need

## Code snippet to run on Terminal

```terminal
$ git clone https://github.com/Arunprakash1990/Robinhood-Portfolio-Analysis.git 

$ cd Robinhood-Portfolio-Analysis

$ pip install -r requirements.txt

$ python robinhood.py
```

*Fig. 1: Sector Analysis in Plotly - iPython Notebook*
![Alt text](https://github.com/Arunprakash1990/Robinhood-Portfolio-Analysis/blob/master/Images/SectorAnalysis-Plotly.png?raw=true)

*Fig. 2: Industry Analysis in Plotly - iPython Notebook*
![Alt text](https://github.com/Arunprakash1990/Robinhood-Portfolio-Analysis/blob/master/Images/IndustryAnalysis-Plotly.png?raw=true)

*Fig. 3: Industry Analysis in Excel*
![Alt text](https://github.com/Arunprakash1990/Robinhood-Portfolio-Analysis/blob/master/Images/IndustryAnalysis-Excel.png?raw=true)

*Fig. 4: Sector Analysis in Excel*
![Alt text](https://github.com/Arunprakash1990/Robinhood-Portfolio-Analysis/blob/master/Images/SectorAnalysis-Excel.png?raw=true)


## Contributing
Pull requests are welcome. For major changes, please open an issue first to discuss what you would like to change.

Please make sure to update tests as appropriate.

