# challenge-VBA-HS

# There is 1 Excel file with the assignment completed as Mulitple_Year_Stock_data_HS
# There is 1 excel file for alphabetical testing for the sample size 
# There is a VB text file containing the actual code for the module; The Excel files should also have the Module inside with the code
# There are 3 Screenshots of, 1 of each Worksheet, showing the solution
# There is 1 ReadMe File with references

# References:

# In order to format the cells to percentage, a code from stackoverflow was studied
# ws.Range("L" & summaryTableRow) = Format(percentChange, "0.00%") was used in my code. The Format function in excel takes a variable and the style of format you want to apply as arguments and returns
# the variable converted in the format requested. For example, Format(0.95,"0.00%") would return 95.00 %. Therefore, in order to get the percent change as a percentage, the year change was divided by the 
# opening balance to get a decimal value and then Format() was used to display it as a percentage with 2 decimal points.
# The code was retrieved from here: https://stackoverflow.com/questions/38830864/format-to-percent-with-10-or-a-lot-of-decimals-in-vba 

# The following site was used to determine the color coding format of the stock market. As per the site: Red means Decrease in change, Green means Increase in change, and blue means no change 
# therefore, for the year change and percent change values where there was no change, i.e.  0, the color code is blue.
# The information was retrieved from here: 
# Investopedia Team. (2022, February 6). What Does a Green Candlestick Mean? Investopedia. 
# https://www.investopedia.com/articles/01/070401.asp#:~:text=Green%20indicates%20the%20stock%20is,from%20the%20previous%20closing%20price.
