# Clear
rm(list = ls())
current_working_dir <- dirname(rstudioapi::getActiveDocumentContext()$path)
setwd(current_working_dir)

# Load Libraries
libraries = c(
  "shiny","tidyverse","readr", "plyr", "data.table","BatchGetSymbols","dplyr",
  "quantmod","stringi", "openxlsx", "Rcpp","xts","stringr","lubridate",
  "readxl", "tidyquant", "optionstrat", "fredr")
suppressWarnings(lapply(libraries, require, character.only = TRUE))
fredr_set_key("abcdefghijklmnopqrstuvwxyz123456")

portfolio <- data.frame(read_xlsx("Portfolio.xlsx"))

Options_Income_Analyzer <- function(portfolio = portfolio){
  # Get Options Expiry Dates - Third Friday of every month
  get_exp_date <- function(start.year, end.year){
    d <- seq(ISOdate(start.year - 1, 12, 1), ISOdate(end.year, 12, 1), by = "1 month")[-1]
    d <- as.Date(d)
    res <- lapply(d, function(x){
      s <- seq(x, by = "day", length.out = 28)
      i <- format(s, "%u") == "5"
      s[i][3]
    })
    
    res <- Reduce(c, res)
    data.frame(Month = format(d, "%Y-%B"), Day = res)
  }
  exp.date.list <- get_exp_date(year(Sys.Date()),year(Sys.Date())+1)
  exp.date.list <- exp.date.list[exp.date.list$Day > Sys.Date(),]
  
  # Get Annual Dividend
  div <- tq_get(portfolio$Ticker, "dividends", 
                from = as.Date(as.yearqtr(Sys.Date())-2/4,frac=1),
                to = as.Date(as.yearqtr(Sys.Date())-1/4,frac=1))
  
  div <-div %>% group_by(symbol) %>%
    arrange(value) %>% filter(row_number()==n())
  
  names(div)[1] <- "Ticker"
  div$value <- (div$value*4)/(getQuote(div$Ticker)$Last)
  portfolio <- merge(portfolio,div,"Ticker", all = TRUE)
  portfolio$date <- NULL
  names(portfolio)[4] <- "Annual Dividend"
  portfolio <- portfolio[!is.na(portfolio$Ticker),]
  
  # Set Exp Date
  exp.date <- exp.date.list[2,2]
  
  # Current Price across all tickers
  portfolio$Price = getQuote(portfolio$Ticker)$Last
  
  #Risk-free rate (US 10Y Treasury Yield)
  getSymbols('DGS10',src='FRED')
  r = as.numeric(tail(DGS10,1))/100
  
  chain.list <- c()
  chain.list.p <- c()
  for (ticker in 1:length(portfolio$Ticker)){
    # Get Call and Put Options Chain for each ticker
    chain = getOptionChain(portfolio$Ticker[ticker], exp.date)$calls
    chain.list[[ticker]] = data.frame(Price = rep(getQuote(portfolio$Ticker[ticker])$Last, nrow(chain)),
                                      Strike = chain$Strike,
                                      Strike.vs.Price = (chain$Strike/rep(getQuote(portfolio$Ticker[ticker])$Last, nrow(chain)))-1,
                                      Premium = (chain$Bid+chain$Ask)/2,
                                      IV = chain$IV, 
                                      Vol = chain$Vol)
    
    chain.p = getOptionChain(portfolio$Ticker[ticker], exp.date)$puts
    chain.list.p[[ticker]] = data.frame(Price = rep(getQuote(portfolio$Ticker[ticker])$Last, nrow(chain.p)),
                                        Strike = chain.p$Strike,
                                        Strike.vs.Price = (chain.p$Strike/rep(getQuote(portfolio$Ticker[ticker])$Last, nrow(chain.p)))-1,
                                        Premium = (chain.p$Bid+chain.p$Ask)/2,
                                        IV = chain.p$IV, 
                                        Vol = chain.p$Vol)
    
    delta <- c()
    raw.yield <- c()
    annual.yield <- c()
    ifreturn.yield <- c()
    contracts <- c()
    income <- c()
    for (i in 1:length(chain.list[[ticker]]$Strike)){
      #Calculate Delta
      temp.delta <- calldelta(s = portfolio$Price[ticker],
                              x = as.numeric(chain.list[[ticker]]$Strike[i]),
                              sigma = as.numeric(chain.list[[ticker]]$IV[i]),
                              t = as.numeric((exp.date-Sys.Date())/365),
                              r = r,
                              d = ifelse(is.na(portfolio$`Annual Dividend`[ticker]),
                                         0,
                                         portfolio$`Annual Dividend`[ticker]))
      delta <- rbind(delta, temp.delta)
      
      #Calculate Raw Yield
      temp.raw.yield <- chain.list[[ticker]]$Premium[i]/portfolio$Cost[ticker]
      raw.yield <- rbind(raw.yield, temp.raw.yield)
      
      #Calculate Annualized Yield
      temp.annual.yield <- (chain.list[[ticker]]$Premium[i]/portfolio$Cost[ticker]) * (365/as.numeric(exp.date-Sys.Date()))
      annual.yield <- rbind(annual.yield, temp.annual.yield)
      
      #Calculate If-Return Yield
      temp.ifreturn.yield <- ((chain.list[[ticker]]$Premium[i] + (chain.list[[ticker]]$Strike[i] - portfolio$Cost[ticker])) / portfolio$Cost[ticker]) 
      ifreturn.yield <- rbind(ifreturn.yield, temp.ifreturn.yield)
      
      #Calculate Contracts
      temp.contracts <- floor(portfolio$Shares[ticker]/100) 
      contracts <- rbind(contracts,  temp.contracts)
      
      #Calculate Income
      temp.income <- 100*chain.list[[ticker]]$Premium[i]*temp.contracts
      income <- rbind(income,  temp.income)
      
      ticker.list <- rep(portfolio$Ticker[ticker],length(chain.list[[ticker]]$Strike))
    }
    chain.list[[ticker]]$Delta <- as.numeric(delta)
    chain.list[[ticker]]$`Raw Yield` <- as.numeric(raw.yield)
    chain.list[[ticker]]$`Annual Yield` <- as.numeric(annual.yield)
    chain.list[[ticker]]$`If-Return Yield` <- as.numeric(ifreturn.yield)
    chain.list[[ticker]]$Contracts <- as.numeric(contracts)
    chain.list[[ticker]]$Income <- as.numeric(income)
    chain.list[[ticker]]$Ticker <- ticker.list
    
    delta.p <- c()
    raw.yield.p <- c()
    annual.yield.p <- c()
    contracts.p <- c()
    income.p <- c()
    cash.p <- c()
    
    for (i in 1:length(chain.list.p[[ticker]]$Strike)){  
      temp.delta.p <- putdelta(s = portfolio$Price[ticker],
                               x = as.numeric(chain.list.p[[ticker]]$Strike[i]),
                               sigma = as.numeric(chain.list.p[[ticker]]$IV[i]),
                               t = as.numeric((exp.date-Sys.Date())/365),
                               r = r,
                               d = ifelse(is.na(portfolio$`Annual Dividend`[ticker]),
                                          0,
                                          portfolio$`Annual Dividend`[ticker]))
      delta.p <- rbind(delta.p, temp.delta.p)
      
      #Calculate Raw Yield
      temp.raw.yield.p <- chain.list.p[[ticker]]$Premium[i]/chain.list.p[[ticker]]$Strike[i]
      raw.yield.p <- rbind(raw.yield.p, temp.raw.yield.p)
      
      #Calculate Annualized Yield
      temp.annual.yield.p <- (chain.list.p[[ticker]]$Premium[i]/chain.list.p[[ticker]]$Strike[i]) * (365/as.numeric(exp.date-Sys.Date()))
      annual.yield.p <- rbind(annual.yield.p, temp.annual.yield.p)
      
      #Calculate Contracts
      temp.contracts.p <- ifelse((floor(20000/chain.list.p[[ticker]]$Strike[i]/100))<0,
                                 0,
                                 floor(20000/chain.list.p[[ticker]]$Strike[i]/100))
      contracts.p <- rbind(contracts.p,  temp.contracts.p)
      
      # Cash requred
      temp.cash.p <- temp.contracts.p*chain.list.p[[ticker]]$Strike[i]*100
      cash.p <- rbind(cash.p, temp.cash.p)
      
      #Calculate Income
      temp.income.p <- 100*chain.list.p[[ticker]]$Premium[i]*temp.contracts.p
      income.p <- rbind(income.p,  temp.income.p)
      
      ticker.list.p <- rep(portfolio$Ticker[ticker],length(chain.list.p[[ticker]]$Strike))
      
    }
    chain.list.p[[ticker]]$Delta <- as.numeric(delta.p)
    chain.list.p[[ticker]]$`Raw Yield` <- as.numeric(raw.yield.p)
    chain.list.p[[ticker]]$`Annual Yield` <- as.numeric(annual.yield.p)
    chain.list.p[[ticker]]$Contracts <- as.numeric(contracts.p)
    chain.list.p[[ticker]]$`Cash Required` <- as.numeric(cash.p)
    chain.list.p[[ticker]]$Income <- as.numeric(income.p)
    chain.list.p[[ticker]]$Ticker <- ticker.list.p
  }
  
  # COVERED CALLS
  
  # Filter Option Chains with Delta <= 0.12 
  chain <- c()
  for (i in 1:length(chain.list)){
    chain[[i]] <- chain.list[[i]][chain.list[[i]]$Delta<0.2 
                                  & chain.list[[i]]$Delta>0,]}
  
  # Find most liquid Option Chain (highest volume)
  res <- c()
  for (i in 1:length(chain)){
    temp.max <- chain[[i]][which.max(chain[[i]]$Vol),]
    res <- rbind(res, temp.max)
  }
  
  res<-merge(portfolio,res,"Ticker")
  names(res)[5] <- "Price"
  res$Price.y <- NULL
  
  # Assuming 50% of Shares on-hand
  res$`Contracts (50%)` <- floor((res$Shares/100)*0.5)
  res$`Income (50%)` <- 100*res$Premium*res$`Contracts (50%)`
  
  res <- res[res$Contracts != 0, ]
  res <- res[order(res$`Annual Yield`, decreasing = TRUE),]
  
  res$Ticker <- paste(res$Ticker,"-",rep("Call",nrow(res)))
  
  # CASH-COVERED SHORT PUTS
  
  # Filter Option Chains with Delta <-0.15 & >-0.35
  chain.p <- c()
  for (i in 1:length(chain.list.p)){
    chain.p[[i]] <- chain.list.p[[i]][chain.list.p[[i]]$Delta<(-0.10) 
                                      & chain.list.p[[i]]$Delta>(-0.35),]}
  
  # Find most liquid Option Chain
  res.p <- c()
  for (i in 1:length(chain.p)){
    temp.max.p <- chain.p[[i]][which.max(chain.p[[i]]$Vol),]
    res.p <- rbind(res.p, temp.max.p)
  }
  
  res.p<-merge(portfolio,res.p,"Ticker")
  names(res.p)[5] <- "Price"
  res.p$Price.y <- NULL
  
  # Assuming 10K Cash
  res.p$`Contracts (10K)` <- ifelse((floor(10000/res.p$Strike/100))<0,
                                    0,
                                    floor(10000/res.p$Strike/100))
  res.p$`Cash Required (10K)` <- res.p$`Contracts (10K)`*res.p$Strike*100
  
  res.p$`Income (10K)` <- 100*res.p$Premium*res.p$`Contracts (10K)`
  
  res.p <- res.p[res.p$Contracts != 0, ]
  res.p <- res.p[order(res.p$`Annual Yield`, decreasing = TRUE),]
  
  res.p$Ticker <- paste(res.p$Ticker,"-",rep("Put",nrow(res.p)))
  
  summary <- merge(res,res.p, all = TRUE)
  summary <- summary[order(summary$`Annual Yield`, decreasing = TRUE),]
  
  # Create workbook summarizing Options Income Analysis
  options_data <- function(){
    wb <- createWorkbook()
    addWorksheet(wb, "Summary")
    writeData(wb, "Summary", summary)
    
    addWorksheet(wb, "Call Details ->")
    for (i in 1:length(chain)){
      addWorksheet(wb, paste(portfolio$Ticker[i], " (Call)"))
      writeData(wb, paste(portfolio$Ticker[i], " (Call)"), chain[[i]])
    }
    addWorksheet(wb, "Put Details ->")
    for (j in 1:length(chain.p)){
      addWorksheet(wb, paste(portfolio$Ticker[j], " (Put)"))
      writeData(wb, paste(portfolio$Ticker[j], " (Put)"), chain.p[[j]])
    }
    saveWorkbook(wb, paste0("Option Income Analysis (",exp.date-Sys.Date()," days from exp of ",exp.date,").xlsx"), overwrite = TRUE)
  }
  options_data()
}

Options_Income_Analyzer(portfolio = portfolio)
