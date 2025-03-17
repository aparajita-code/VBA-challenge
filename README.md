# VBA-challenge

Quarterly Stock Data Analysis with VBA
Welcome to the Quarterly Stock Data Analysis repository! 🚀 This project demonstrates how to use VBA (Visual Basic for Applications) to process and analyze stock data from a single quarter efficiently. The project includes data manipulation, calculated values, and the application of conditional formatting for better insights.

***🎯 Project Objectives***
This project aims to:

Retrieve and Store Stock Data: Extract critical details like the ticker symbol, total volume, open price, and close price from each row.

Perform Calculations: Compute essential metrics such as total stock volume, quarterly change, and percent change.

Automate Workflow: Ensure seamless analysis across multiple worksheets using a dynamic VBA script.

Enhance Visualization: Apply conditional formatting to highlight key data trends effectively.

***🛠️ Features***
# 1. Data Retrieval and Storage
Extracts and stores key information from stock data:

Ticker Symbol

Volume of Stock

Open Price

Close Price

# 2. Column Creation
Creates the following additional columns for analysis:

Ticker Symbol

Total Stock Volume

Quarterly Change ($)

Percent Change

# 3. Conditional Formatting
Highlights:

Quarterly Change column based on positive or negative values.

Percent Change column to emphasize significant increases or decreases.

# 4. Calculations
Automatically calculates and displays:

Greatest Percent Increase

Greatest Percent Decrease

Greatest Total Volume

# 5. Sheet Automation
Loops through all worksheets to ensure script functionality across multiple sheets seamlessly.

***📋 VBA Script Overview***
The VBA script includes:

Dynamic looping across rows and columns.

Effective use of conditional formatting for data visualization.

Logical structure for calculating key metrics.

***📊 Outputs***
The outputs include:

Tabulated analysis with calculated values.

Highlighted cells for easy trend identification.

Summary of greatest changes and volumes for quick interpretation.

***📝 Instructions to Run***
Open the VBA Editor (Alt + F11) in your Excel file.

Copy and paste the provided VBA script into a new module.

Run the script (F5) and watch as the data is processed and formatted automatically.

Check the output in the same worksheet or newly created sheets.

***📂 Repository Contents***
This repository includes:

Screenshots: Visual results of the analysis.

VBA Script: Modular and well-documented script files.

README: Detailed project documentation (this file).

***📈 Example Visualization***
Conditional formatting highlights data trends, providing a clear overview of stock performance across the quarter.

***📜 License***
This project is open-source and available under the MIT License.


-------------------------------------------------------------------------------------------------------------------------------------------------------

***Additional Notes**


There are Multiple attempts to solve the problem.
1st RAW attempt is in Initial_Solution_stock_analysis.vba
This file gets the results but without formmatting like % or color formatting. This also doesnt have the Max % Increase or Decrease summary results. But this code can run at once for all worksheets

Solution_final_stockanalysis.vba
This file contains the primary solution including % formatting

summarystock_values.vba
This file contains the Max % increase / decrease changes for the Tickers and Stock volumnes. To get the results for this file Solution_final_stockanalysis.vba needs to be run first


conditional_color_formatting.vba 
This file contains the color formatting for Quarterly +/- change



NOTE : Any files / code for Every Single worksheet.

References : stackoverflow.com and https://learn.microsoft.com/en-us/ for Syntaxes / logics at times.




