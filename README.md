# Monte-Carlo-Stock-Simulation
This utilized Python to perform a Monte Carlo Simulation, and then populates a Power BI Report. This can be used to help assist in a risk assessment on a stock.

Monte Carlo Simulations perform scenario analysis over many iterations to predict an outcome, the benefit of this is it will help to determine more precise risk profile on a potential outcome.

This Python Program utilizes the Monte Carlo technique as well as a technique called GARCH (Generalized Autoregressive Conditional Heteroskedasticity). GARCH can be summarized as a technique which is effective for utilizing time-series data with volatility clustering, where periods of high volatility are followed by periods of low volatility, this can be particularly useful for analyzing stocks.

This model takes no responsibility for any stock losses, any risks on trading stocks or securities can only be taken by the person investing in the security.

Instructions:
First, you need to download the Data file for a stock from Yahoo. For this analysis, I am performing analysis on the NVDA stock. Please use this link and select Download: "https://finance.yahoo.com/quote/NVDA/history/?guccounter=1&guce_referrer=aHR0cHM6Ly93d3cuZ29vZ2xlLmNvbS8&guce_referrer_sig=AQAAANlRrYwftbR-wvAtVOJSTxfOh05SCcq9q1dIjwRRvL7rJHoZXds8D3VhVMsCMquDtaNImI8JkpcZtITwRXh8qEBFkvQgLD1d-Nkc07Bxvln_fxWJ6VfiwZcI5GTZ1hBlBvpH5fRGYBV0VRAxAnZB8tjXLMuqVBcagKPWhh7wcK-L"

Second, save this to a folder as an .xlsx file, you will need to have it in an .xlsx file for the excel python program to work. It automatically saves as a .csv file when downloaded from the website.

Third, update the Python script URL to capture where you saved the .xlsx file. Make sure to save the python file in the same location as the .xlsx file. Remember, Python uses forward slashes in URL "/", when you copy from File Explorer, the path may show backslashes "\". If that isn't fixed, the program will not work.

Fourth, once you have performed running the Python file, you will need to update the path in the Power Query within Power BI, go to the Queries in Power BI and select Edit Query, then go to the applied steps and select Source, while there, you will see a URL, make sure the path matches where you saved the .xlsx file and make sure the file is named consistently.
