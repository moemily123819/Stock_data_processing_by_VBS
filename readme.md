# ***VB Script***

 

### **Objective :**

A VB script is needed to process several years' stock information to find out each stock's total volume, net change and percentage change for the year.

 

### **Author :**

Emily Mo

 

### **About the data :**

An excel workbook containing multiple years stock data is provided.  stock symbols from A to Z, its monthly opening price, closing price, high and low prices as well as the monthly volume are provided.  Every worksheet is a year's data. 



### **About the VB script :**

A macro is created to process every all the worksheets one by one (every year's data).  Using for loop to process from the second row (row 1 is the columns) to the last row, the VB script processes all the rows for the same stock symbol to get its yearly opening price, the year-end closing price, accumulate the total volume, and to calculate the yearly percentage change.  If the change is a gain, the cell will be green, if not, it will be red.  At the same time, it records the stock for the greatest percentage increase, the greatest percentage decrease as well the highest yearly volume. 

Technical features used :

- VB script, 
- excel worksheet button and its corresponding macro.

  



### Deployment :

The VB script can be evoked by a macro in the excel workbook.



