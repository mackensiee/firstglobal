# FIRST Global Rankings Scraper

This is a program to web scrape match data for FIRST Global from the results webpage.
https://results.first.global/

To set this up:
1. Install ChromeDriver (https://developer.chrome.com/docs/chromedriver).
2. In the python script, change the string `chrome_driver_path` to be the path to your ChromeDriver.
3. Set up an Excel workbook.
4. Change the string `excel_wb_path` to be the path to your Excel file.
5. Change the string `excel_wb_tab_name` to be the tab in your Excel file where you put the matches.

When you run the script, it should insert the match data into the tab of your excel file. This will need to be run manually every time you want new data.

