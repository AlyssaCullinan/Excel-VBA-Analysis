# Excel-VBA-Analysis

## Overview
This repository contains a VBA script designed to analyze stock market data. The script is designed to loop through all the stocks for one year, extracting key information such as yearly change, percentage change, and total stock volume. The script identifies the stock with the greatest percentage increase, greatest percentage decrease, and greatest total volume across all years.


## Contents
- [Technologies](#Technologies)
- [Usage](#Usage)
- [Output](#Output)
- [Functionality](#Functionality)
- [Instructions](#Instructions)
- [Example](#Example)

## Technologies
* Microsoft Excel
    * VBA

## Output
The script outputs the following information for each stock:

* Ticker symbol
* Yearly change in stock price
* Percentage change from the opening price to the closing price
* Total stock volume
  
The script also identifies and returns the stocks with:
    * Greatest percentage increase
    * Greatest percentage decrease
    * Greatest total volume
    
## Functionality
The VBA script uses the following functionalities:
* Loops through all the worksheets (years) in the Excel file to analyze each year's stock data.
* Calculates the yearly change, percentage change, and total stock volume for each stock.
* Identifies the stocks with the greatest percentage increase, greatest percentage decrease, and greatest total volume across all years.
  
<img width="464" alt="image" src="https://github.com/AlyssaCullinan/Excel-VBA-Analysis/assets/141466633/a9abfd06-8196-49c3-a1f3-d589a2fc448c">

<br>
<br>


<img width="721" alt="Screenshot 2023-09-18 071147" src="https://github.com/AlyssaCullinan/Excel-VBA-Analysis/assets/141466633/912e15cd-dab6-4ac8-8759-11eafd27d1d3">


## Usage
1. Open the Excel file containing the stock market data.
2. Enable macros if prompted.
3. Navigate to the Developer tab and click on "Visual Basic" to open the VBA editor.
4. Insert a new module.
5. Copy and paste the provided VBA script into the module.
6. Save the file and close the VBA editor.
7. Run the macro named analyzeStocks.
