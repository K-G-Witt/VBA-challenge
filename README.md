# VBA-Challenge

## Project Description:
The overall purpose of this project is to create and use a script using Visual Basic for Application (VBA) programming language to analyse trends in stock market data over three years: 2018, 2019, and 2020.
   
## Installation and Run Instructions:
This repository contains the following files:
1. **alphabetical_testing_macro.xlsm:** a testing dataset saved as a macro-enabled Excel workbook.
2. **Multiple_year_stock_data_macro.xlsm:** a validation dataset saved as a macro-enabled Excel workbook.
3. **VBA_Challenge_Script.txt:** the VBA script used to generate all analyses saved as a text file.

## Usage Instructions:
Executing the VBA script provided in the **VBA-Challenge_Script.txt** file will loop through all the stocks for each year in turn and will outputs the following information:
1. The ticker symbol;
2. Yearly change from the opening price at the beginning of a given year to the closing price at the end of that year;
3. The percentage change from the opening price at the beginning of a given year to the closing price at the end of that year, and;
4. The total stock volume of the stock.

Added functionality is also included in this VBA script to return, for each year, the stock with:
1. The "Greatest % increase";
2. The "Greatest % decrease", and;
3. The "Greatest total volume". 


## Sample Output:
Sample output is provided in the following screenshorts uploaded to this repository:
1. **2018_Output_Screenshot.png**
2. **2019_Output_Screenshot.png**
3. **2020_Output_Screenshot.png**

## Credits:
This code was compiled and written by me for the VBA class homework in the 2024 Data Analytics Boot Camp hosted by Monash University. Additional credits are declared below:

### Defining and holding the Opening_Price var to enable calculation of the Yearly_Change var:
During the course of this assignment, I encourtered a challenge in ensuring the script correctly identified and held the value of the Opening_Price var for each Ticker var to, in turn, correctly calculate the Yearly_Change var. After experimenting with different methods myself, I found the following resource from Jeremy Tallant's GitHub. Source: https://github.com/JeremyTallant/VBA-challenge/blob/main/VBA-code.vba#L1 (accessed 2 March 2024).

### Saving VBA Script as .txt file:
Following recommendations, VBA script was saved as a .txt file to ensure compatability between different computer systems. Source: https://www.geeksforgeeks.org/how-to-make-save-and-run-a-simple-vbscript-program/ (accessed 4 March 2024).
