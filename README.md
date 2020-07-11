# Robinhood Trades to Excel with Additional Analysis

This project connects to your Robinhood Account using pyrh (the unofficial API for Robinhood) and segregate the data based on Industry and Sector listed in NASDAQ,NYSE stock Exchange to give additional insights on your portfolio. This Python application exports your Robinhood trades to a .xlsx file with additional Analysis:
- Portfolio Diversity Classification 
- Profit/Loss % for each transaction
- Sector Summary Classification
- Industry Summary Classification



pyrh: [Robinhood library](https://github.com/robinhood-unofficial/pyrh) created by Rohan Pai. Works on Python 2.7+ and 3.5+

## Installation

Install all the dependencies related to the Project. Also, make sure you have python installed in your local machine.

```bash
pip install -r requirements.txt
```

## Usage
- Run the below command on your terminal
```bash
python robinhood.py
```
- Look for an Excel file on your project directory which will contains all the above analysis
- If you are comfortable in python code and want to see everything on the fly, fork the iPython notebook and customize the code to your need

## Contributing
Pull requests are welcome. For major changes, please open an issue first to discuss what you would like to change.

Please make sure to update tests as appropriate.

