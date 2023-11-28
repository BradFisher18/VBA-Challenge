# VBA-Challenge
Repository for Module 2 VBA Challenge
This code is targeted for analysts wanting to evaluate the performance of a company's stock value. The code is written to be run using VBA, with the data inputs to be presented in a Microsoft Excel Spreadsheet. 
The inputs required for successful running of this program, is the daily information for a comnpany with each worksheet including all information for a specified year. This list may include multiple companies, and the following are to be inpputted into the worksheet broken with a new row for each day:
  - company ticker
  - date of stock information
  - daily opening stock price
  - daily highest value of stock price
  - daily lowest value of stock price
  - daily closing stock price
  - daily stock volume
By runnning this code, it will produce the following outputs:
  - yearly change of stock price for each company
  - yearly change as a percentage
  - each companies total stock volume for the year
  - In a seperate table, it will output the outputs described above but only for the companies that have the highest and lowest percentage changes in stock price, as well as the company with the greatest total stock volume.

HOW TO RUN THIS CODE
To commence running this code, in your Microsoft Excel workbook ensure that the 'Developer' tab on the ribbon is selected and Macros is enabled. If the developer tab is not visible, go into the File tab at the top left, select Options down the bottom, and then under the Customize Ribboin tab ensure Developer is selected. To check in Macros is enabled, in the same Options menu select the Trust Centre and in the Trust Centre Settings you will find a Macros tab. Here you will select Enable VBA Macros.
You should now be able to select Visual Basic in the Developer tab, which will open a window. Create a new Module in the workbook, and insert the code from the repository there. You should now be able to run this code and it will provide you with the results as outlined above.
