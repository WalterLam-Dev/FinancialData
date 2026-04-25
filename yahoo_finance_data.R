library(tidyquant)
library(ggplot2)
library(dplyr)
library(tibble)
library(Rcpp)
library(quantmod)
library(rvest)


### Set directory
setwd(dirname(rstudioapi::getSourceEditorContext()$path))
dl_complete = FALSE

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
yfinanceData <- function(ticker, fromDate='1920-01-01') {
  ### Download the data as xts object
  dataObject = loadSymbols(ticker, 
                           from = fromDate, 
                           verbose = TRUE, 
                           warnings = FALSE, 
                           auto.assign = FALSE)
  
  ### Convert the data to dataframe
  dataObject.df = as.data.frame(dataObject, check.names=FALSE)
  colnames(dataObject.df) = c('Open', 'High', 'Low', 'Close', 'Volume', 'Adj Close')
  dataObject.df <- dataObject.df[, c('Open', 'High', 'Low', 'Close', 'Adj Close', 'Volume')]
  
  ### Put the date index to column
  dataObject.df = tibble::rownames_to_column(dataObject.df)
  names(dataObject.df)[names(dataObject.df) == 'rowname'] = 'Date'
  dataObject.df$Date <- gsub("X", "", dataObject.df$Date) # remove the char 'X'
  dataObject.df$Date = YMD(dataObject.df$Date)
  
  ### Remove rows with NA in any column
  dataObject.df = na.omit(dataObject.df)
  
  ### Read the csv file saved before (if applicable)
  tryCatch({
    ### check.names=FALSE to avoid converting ' ' in column names to '.'
    df_original = read.csv(file = paste0(ticker, ".csv"), header = TRUE, check.names=FALSE)
    df_original$Date <- as.Date(df_original$Date, format = "%Y-%m-%d")
    
    # Append the new data if it is not in the file saved before
    dataObject.df = dataObject.df[!dataObject.df$Date %in% df_original$Date, ]
    dataObject.df = rbind(df_original, dataObject.df)
    
    ### Rearrange the rows by date in ascending order
    dataObject.df <- dataObject.df[order(dataObject.df$Date), ]
  }, error=function(e){cat("ERROR:",conditionMessage(e), "\n")})
  
  ### Reset row index
  rownames(dataObject.df) <- NULL
  
  ### Save as csv file
  write.csv(dataObject.df, file = paste0(ticker, ".csv"), row.names = FALSE)
}
getDiv_AASTOCKS <- function(ticker = "SCHD") {
  library(polite)  # respectful webs craping
  
  ### Make our intentions known to the website
  ### Extract data from AASTOCKS (only cash dividend: "filter=C")
  url_aastocks = paste0("http://www.aastocks.com/en/usq/analysis/dividend.aspx?symbol=", 
                        ticker, 
                        "&filter=C")
  bow <- bow(
    url = url_aastocks,  # base URL
    user_agent = "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/17.5 Safari/605.1.15",  # identify ourselves
    force = TRUE)
  
  print(bow)
  
  ### save the page (as html) after scraping
  page = scrape(bow)
  
  ### extract the table from the page
  ### the data is saved under: <table class="...">, so we put "table" as input variable
  table_node = html_nodes(page, "table")
  
  ### run html_table(table_node) to find the index (23) of the table needed
  table_i = 0
  for (i in 1:length(table_node)) {
    if (names(html_table(table_node, header = TRUE)[[i]])[1] == "Ex-date") {
      table_i = i
    }
  }
  table_content <- html_table(table_node, header = TRUE)[[table_i]]
  
  ### Clean the data before saving
  ### Rearrange the rows by date in ascending order
  ### Website date format: YYYY/mm/dd
  for (c in c("Ex-date", "Record date", "Pay date")) {
    table_content[[c]] <- as.Date(table_content[[c]], format = "%Y/%m/%d")
  }
  table_content <- table_content[order(table_content$`Ex-date`),]
  
  ### Rearrange the columns
  table_content = table_content[, c("Ex-date", "Amount", "Record date", 
                                    "Div. Type", "Pay date")]
  
  ### Convert the dividend column as numeric
  ### "sub(".*\\ ", "", x)" means substitute everything until a space " " as ""
  table_content$Amount <- as.numeric(sub(".*\\ ", "", table_content$Amount))
  
  ### Rename the "Particular" as the ticker
  names(table_content)[names(table_content) == "Amount"] <- paste0(ticker, "_Div")
  
  ### Read the csv file saved before (if applicable)
  tryCatch({
    ### check.names=FALSE to avoid converting ' ' in column names to '.'
    df_original = read.csv(file = paste0(ticker, "_Div_aastocks.csv"), 
                           header = TRUE, 
                           check.names=FALSE)
    
    # Excel date format: d/m/YYYY
    # R csv read date format: YYYY-MM-DD
    for (c in c("Ex-date", "Record date", "Pay date")) {
      df_original[[c]] <- as.Date(df_original[[c]], format = "%Y-%m-%d")
    }
    
    # Append the new data if it is not in the file saved before
    table_content = table_content[!table_content$`Ex-date` %in% df_original$`Ex-date`, ]
    table_content = rbind(df_original, table_content)
    
    ### Rearrange the rows by date in ascending order
    table_content <- table_content[order(table_content$`Ex-date`), ]
  }, error=function(e){cat("ERROR:",conditionMessage(e), "\n")})
  
  ### Reset row index
  rownames(table_content) <- NULL
  
  ### Save the file
  write.csv(table_content, 
            file = paste0(ticker, "_Div_aastocks.csv"), 
            row.names = FALSE)
} # for US stocks only
getDiv_nasdaq <- function(ticker = "SCHD", append = TRUE, delay = 5) {
  library(polite)  # respectful webs craping
  
  ### Input variables
  if (append == TRUE && file.exists(paste0(ticker, "_Div_nasdaq.csv"))) { 
    original_div = read_csv(paste0(ticker, "_Div_nasdaq.csv"), 
                            show_col_types = FALSE)
    
    ### Convert the date column from Excel/system format "mm/dd/yyyy" to readable dates
    dateCol = c("exOrEffDate", "paymentDate", "declarationDate", "recordDate")
    for (c in dateCol) {
      ### Convert excel date (dd/mm/yyyy) to R date
      original_div[[c]] <- as.Date(original_div[[c]], format = "%d/%m/%Y")
    }
  }
  assetclass = "etf"
  ua = paste0("Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) ", 
                      "AppleWebKit/605.1.15 (KHTML, like Gecko) Version/", 
                      "17.5 Safari/605.1.15")
  url_nasdaq = paste0("https://api.nasdaq.com/api/quote/", ticker, 
                      "/dividends?assetclass=", assetclass)
  
  ### Make our intentions known to the website
  ### Extract data from nasdaq.com
  bow <- bow(url = url_nasdaq,  # base URL
             user_agent = ua,  # identify ourselves
             force = TRUE, 
             delay = delay+abs(rnorm(1)))
  
  print(bow)
  
  ### Save the page (as html) after scraping
  page = scrape(bow)
  
  ### Change the assetclass if "etf" doesn't work
  if (is.null(page$data)) {
    assetclass = "stocks"
    url_nasdaq = paste0("https://api.nasdaq.com/api/quote/", ticker, 
                        "/dividends?assetclass=", assetclass)
    bow = bow(url = url_nasdaq, user_agent = ua, force = TRUE, delay = delay+abs(rnorm(1)))
    page = scrape(bow)
  }
  
  ### Convert the table from the page as data.frame()
  ### Source: https://www.geeksforgeeks.org/convert-list-of-lists-to-dataframe-in-r/
  div_table = data.frame(t(unlist(page$data$dividends$rows[[1]])))
  for (i in 2:length(page$data$dividends$rows)) {
    div_table[i, ] = t(unlist(page$data$dividends$rows[[i]]))
  }
  
  ### Clean the data before saving
  ### Rearrange the columns
  div_table = div_table[, c("exOrEffDate", "amount", "type", "currency", 
                            "paymentDate", "declarationDate", "recordDate")]
  
  ### Convert the dividend column as numeric
  ### "sub(".*\\$", "", x)" means substitute everything until a space "$" as ""
  div_table$amount <- as.numeric(sub(".*\\$", "", div_table$amount))
  
  ### Convert the date column from NASDAQ format "mm/dd/yyyy" to readable dates
  dateCol = c("exOrEffDate", "paymentDate", "declarationDate", "recordDate")
  for (c in dateCol) {
    div_table[[c]] <- as.Date(div_table[[c]], format = "%m/%d/%Y")
  }
  
  ### Rename the column "amount" as the ticker
  names(div_table)[names(div_table) == "amount"] <- paste0(ticker, "_Div")
  
  ### Append the new dividend record to original file
  if (append == TRUE && file.exists(paste0(ticker, "_Div_nasdaq.csv"))) {
    div_to_append = div_table[!div_table$exOrEffDate %in% original_div$exOrEffDate, ]
    div_table = rbind(original_div, div_to_append)
  }
  
  ### Rearrange the rows by date in ascending order
  div_table <- div_table[order(div_table$exOrEffDate), ]
  row.names(div_table) = NULL
  
  ### Save the file
  write.csv(div_table, 
            file = paste0(ticker, "_Div_nasdaq.csv"), 
            row.names = FALSE)
} # for US stocks only

### Data Extraction
### Collapse the list: Command + option + L
### Expand the list: Command + option + Shift + L
### Stock price list
allTickers = c(
               '^SPXEW',    # S&P Equal Weight Index
               'RSP', 
               'HG=F',      # Copper Price
               'GC=F',      # Gold Price
               '^BCOM',     # BBG Commodity
               '^GSPE',     # Energy
               '^SP500-15', # Materials
               '^SP500-20', # Industrials
               '^SP500-25', # Consumer Discretionary
               '^SP500-30', # Consumer Staples
               '^SP500-35', # Healthcare
               '^SP500-40', # Financials
               '^SP500-45', # Information Technology
               '^SP500-50', # Communication Services
               '^SP500-55', # Utilities
               '^SP500-60', # Real Estate
               '^RUT',      # Russell 2000
               '^RUI',      # Russell 1000 
               '^OEX',      # S&P 100
               '^MID',      # S&P Mid-Cap 400
               '^SP600',    # S&P SmallCap 600
               '^GSPC',     # S&P 500 Index
               '^SP500TR',  # S&P 500 Total Return Index
               '^IXIC',     # NASDAQ Composite
               '^HSI',      # Hang Seng Index
               'UPRO',
               '^W5000',    # Wilshire 5000 Total Market Index
               '^990100-USD-STRD', # MSCI World Index
               'EEM',       # iShares MSCI Emerging Markets ETF
               'USMV',      # iShares MSCI USA Min Vol Factor ETF
               'SPLV',      # Invesco S&P 500 Low Volatility ETF
               'SPHB',      # Invesco S&P 500 High Beta ETF
               'SCHD', 
               '^IRX',      # US 13W Treasury Yield (3M)
               '^TNX',      # US 10Y Treasury Yield
               '^TYX',      # US 30Y Treasury Yield
               'EDV',       # Vanguard Extended Duration Treasury Index Fund ETF Shares
               # '^DJGSP',  # Dow Jones Precious Metals Index
               'DBB',       # Invesco DB Base Metals Fund (Proxy of Industrial Metals)
               'DBP',       # Invesco DB Precious Metals Fund (Proxy of Precious Metals)
               '^XAU',      # Philadelphia Gold and Silver Index
               'DX-Y.NYB',  # US Dollar Index
               '^SP500-1010',    # Energy
               '^SP500-1510',    # Materials
               '^SP500-2010',    # Capital Goods
               '^SP500-2020',    # Commercial & Professional Services
               '^SP500-2030',    # Transportation
               '^SP500-2510',    # Automobiles & Components
               '^SP500-2520',    # Consumer Durables & Apparel
               '^SP500-2530',    # Consumer Services
               '^SP500-2550',    # Retailing
               '^SP500-3010',    # Food & Staples Retailing
               '^SP500-3020',    # Food Beverage & Tobacco
               '^SP500-3030',    # Household & Personal Products
               '^SP500-3510',    # Health Care Equipment & Services
               '^SP500-3520',    # Pharmaceuticals, Biotechnology & Life Sciences
               '^SP500-4010',    # Banks
               '^SP500-4020',    # Diversified Financials
               '^SP500-4030',    # Insurance
               '^SP500-4510',    # Software & Services
               '^SP500-4520',    # Technology Hardware & Equipment
               '^SP500-4530',    # Semiconductors & Semiconductor Equipment
               '^SP500-5010',    # Telecommunication Services
               '^SP500-5020',    # Media & Entertainment
               '^SP500-5510',    # Utilities
               '^SP500-6010',    # Real Estate
               'GC=F',           # Gold Futures
               'HG=F',           # Copper Futures
               'ZN=F',           # 10-Year T-Note Futures
               'ZB=F',           # 30-Year T-Bond Futures
               'UB=F',           # Ultra U.S. Treasury Bond Future (>25Y)6CC
               'VOOG',           # Vanguard S&P 500 Growth ETF
               'VUG',            # Vanguard Growth ETF
               'VIGRX',          # Vanguard Growth Index
               'VIVAX',          # Vanguard Value Index
               '^SHORTVOL'       # Short VIX Futures Index
               )
### Div list
divTickers = c("SCHD")

for (ticker in allTickers) {
  dl_complete = FALSE
  
  ### Try until the download is complete
  while (dl_complete == FALSE) {
    tryCatch({
      yfinanceData(ticker = ticker)
      dl_complete = TRUE
    }, error=function(e){cat("ERROR :",conditionMessage(e), "\n")})
  }
  ### To avoid IP address being blocked
  # Sys.sleep(5)
  
  ### if the download is complete
  # dl_complete = TRUE
}

for (ticker in divTickers) {
  getDiv_AASTOCKS(ticker)
  # getDiv_nasdaq(ticker)
}

print(paste("Files were saved at ", getwd(), ".", sep = ""))