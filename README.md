<img align = "right" width="75" alt="logo" src="https://user-images.githubusercontent.com/97678601/149636901-fb79e698-7c0e-47fb-bb88-033785485fc7.png"> 

# DV Options Income Analyzer

The DV Options Income Analyzer serves as an investment research tool for investors who are interested in deploying **low-risk options trading strategies for the core purposes of generating passive income and maintaining healthy portfolio management.**

Options trading is commonly perceived as highly risky and speculative style of investing. Unfortunately, like many stereotypes, this perception may not be completely off the mark. As retail interest in investing continues to surge, less-experienced traders can find themselves enamored by the mindblowing potential % returns that options can offer off a limited amount of capital. Ironically, the intent for options was actually to serve as a tool to hedge risk against large movements in share prices. But, there's no fun in solely following the intent of anything in this world, is there?! Especially if an alternative use-case that can make you money for limited risk exists. 

Well-tested options strategies such as: selling covered calls, selling cash-secured puts, vertical spreads, and the iron condor, if deployed smartly, can offer a risk-minimal way to generate additional yield/income for investor portfolios. **The DV Options Income Analyzer automates the tedious work of analyzing the plethora of option chain and greek data across tickers and expiry dates and recommends risk-optimized contracts for writing calls and puts.**

## Functionality:
The program prompts users to input a stock portfolio or watchlist and generates an excel-based report that recommends risk-minimized strike prices for Put and Call option contracts that can be shorted by the user on their desired brokerage account for the core purpose of generating income through premiums.

As summarized below, for a given stock, the program considers two primary factors of a contract prior to recommendation: delta and volume.
- Calls:
  - Delta less than 0.2, and
  - Highest volume (i.e. most liquid contract)
- Puts:
  - Delta greater than 0.1 and less than 0.35, and
  - Highest volume (i.e. most liquid contract)

The auto-generated report includes relevant analytical measures, including:
- Strike vs. Current Price
- Implied Volatility
- Volume
- Delta
- Raw Yield
- Annual Yield
- Number of Contracts available to be sold (for covered calls and cash-secured puts)
- Income

## Input Files and Libraries:
The only input file required for this program is "Portfolio.xlsx". The spreadsheet contains three columns - "Ticker", "Cost", "Shares". "Ticker" is a required column, while "Cost" and "Shares" are optional and can be filled out if the user has an existing position in the security.

The program uses the following libraries:
- quantmod
- data.table
- dplyr
- openxlsx
- lubridate
- readxl
- tidyquant
- optionstrat
- fredr

## Future Enhancements
As a future enhancement, I would like to replace the excel-based report with a shiny-based user interface that can refresh periodically to make the user experience more enjoyable and efficient. 
