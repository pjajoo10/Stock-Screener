# Stock-Screener (v2 in progress)

The goal is to filter about 500 stocks based on metrics (such as divdend yield, debt/equity etc.) from the National Stock Exchange of India to find value investments.

The filtering process, the most tedious task of checking each and every stock's performance metrics, is done automatically using Selenium to gather the data of each and every stock. Manual work is required, to upload the daily report of the stock market and select the files in the code (changing in v2 to be autmatic).

### Instructions to run
1. Requires the webdriver of the browser being used. If not using Firefox download the webdriver for your web browser and upload it in the `webdrivers` folder and update the path in `main.py` (geckodriver for Firefox in the repo)
2. Install the requirements.
3. Run `main.py`
4. Output will be saved to `stock-data.xlsx`


###### Version 2 being worked upon - deploy the script on a web app, further automation to make it user-friendly, more filtering mechanics.
