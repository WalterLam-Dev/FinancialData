library(tidyquant)
library(ggplot2)
library(dplyr)
library(tibble)
library(Rcpp)

### Path to save the files
setwd("C:/Users/cwlam/Documents/R Projects")

### Functions
lnRtn <- function(series, wdw) {
  # series: a vector, column of prices
  # wdw: return window, wdw=5 for 5D returns
  emptyArray = rep(NA, wdw)
  logDiff = diff(log(series), lag = wdw)
  
  return(c(emptyArray, logDiff))
}
rollingSD <- function(series, wdw) {
  ### Input variable
  ### series: the series of numbers from a column, in terms of price change
  ### wdw: the window for calculating s.d. in Day, e.g. wdw=60 means 3M s.d.
  
  # library("Rcpp")
  # rollSD = roll_sd(series, n = wdw, na.rm = FALSE, fill = NA)
  # return(c(rollSD))
  
  seriesLength = length(series)
  outputSD = c(rep(NA, wdw-1))
  
  for (n in wdw:seriesLength) {
    stddev = sd(series[(n-wdw+1):n])
    
    outputSD = c(outputSD, stddev)
  }
  
  return(outputSD)
}
yfinanceData <- function(ticker, fromDate='1960-01-01') {
  ### Download the data as xts object
  dataObject = loadSymbols(ticker, 
                           from = fromDate, 
                           verbose = TRUE, 
                           warnings = TRUE, 
                           auto.assign = FALSE)
  
  ### Convert the data to dataframe
  dataObject.df = as.data.frame(dataObject)
  colnames(dataObject.df) = c('Open', 'High', 'Low', 'Close', 'Volume', 'Adj Close')
  dataObject.df <- dataObject.df[, c('Open', 'High', 'Low', 'Close', 'Adj Close', 'Volume')]
  
  ### Put the date index to column
  dataObject.df = rownames_to_column(dataObject.df)
  names(dataObject.df)[names(dataObject.df) == 'rowname'] = 'Date'
  dataObject.df$Date = YMD(dataObject.df$Date)
  
  ### Remove rows with NA in any column
  dataObject.df = na.omit(dataObject.df)
  
  ### Save as csv file
  write.csv(dataObject.df, file = paste0(ticker, ".csv"), row.names = FALSE)
  }


### Data Extraction
### Collapse the list: Alt + L
### Expand the list: Alt + Shift + L
allTickers = c('^GSPC',   # S&P 500 Index
               'ZB=F',    # US 30-Year Treasury Bond Futures
               'SPHB',    # S&P 500 High Beta
               'SPLV',    # S&P 500 Low Volatility
               '^SPXEW',  # S&P EQUAL WEIGHT INDEX
               '^MOVE',   # ICE BofAML MOVE Index
               '^990100-USD-STRD'  # MSCI WORLD
               )

for (ticker in allTickers) {
  yfinanceData(ticker = ticker)
}


# schd_div = getDividends("SCHD",
#                         from = "1970-01-01",
#                         src = "yahoo",
#                         auto.assign = FALSE,
#                         auto.update = FALSE,
#                         verbose = FALSE,
#                         split.adjust = FALSE)