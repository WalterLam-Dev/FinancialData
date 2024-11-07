#############################
###### MSCI FUNCTIONS #######
#############################

library(quantmod)
library(readxl)
library(lubridate)

# The script will download data from MSCI's end of day history app, example as below:
# https://app2.msci.com/products/service/index/indexmaster/downloadLevelData?output=INDEX_LEVELS&currency_symbol=USD&index_variant=STRD&start_date=20201106&end_date=20241106&data_frequency=DAILY&baseValue=false&index_codes=704866,704865

setwd("C:/Users/cwlam/Documents/R Projects")
options(HTTPUserAgent = paste0("Mozilla/5.0 (Windows NT 10.0; Win64; x64) ", 
                               "AppleWebKit/537.36 (KHTML, like Gecko) ", 
                               "Chrome/130.0.0.0 Safari/537.36 Edg/130.0.0.0"))

# Download data
msci_eod_dl <- function(output = "INDEX_LEVELS", 
                        currency = "USD", 
                        idx_var = "STRD", 
                        start_dt = "19691231", 
                        end_dt = format(Sys.Date(), format = "%Y%m%d"), 
                        freq = "DAILY", 
                        basevalue = "false", 
                        idx_code = 704866
                       ) {
  
  # Transform the list of index id
  if (length(idx_code) > 1) {
    idx_code = paste(idx_code, collapse = ",") 
  }
  
  # Construct the URL
  URL = paste0("https://app2.msci.com/products/service/index/indexmaster/downloadLevelData?", 
               "output=", output, 
               "&currency_symbol=", currency, 
               "&index_variant=", idx_var, 
               "&start_date=", start_dt, 
               "&end_date=", end_dt, 
               "&data_frequency=", freq, 
               "&baseValue=", basevalue, 
               "&index_codes=", idx_code)
  
  # Download the .xls file
  file_name = paste0("MSCI_", idx_code)
  download.file(url = URL, 
                destfile = "MSCI_temp.xls", 
                mode = "wb", 
                quiet = TRUE)
  
  ### Data cleaning and save again
  # Read the .xls file
  df = read_excel("MSCI_temp.xls")
  df = df[-(1:(which(df$...1=="Date")-1)), ] # only keeps rows after 'Date'
  names(df) = df[1,] # put first row as column names
  df = df[-1,] # remove the first row
  df = na.omit(df) # remove row with any NA
  
  # Change the first column as date format
  df$Date = mdy(df$Date)
  
  # Save the file as csv
  write.csv(df, paste0(file_name, ".csv"), row.names = FALSE)
}

# To check the index code:
# 1. Select the indices from the MSCI website (https://www.msci.com/end-of-day-data-search)
# 2. Go to Developer (Ctrl + Shift + I) -> Network
# 3. Click the "Download Data" button in the website
# 4. Check the name of the new item after clicking the button 
index_code = list(c(704866,704865,984000), # Cyclical, Defensive, All Strd
                  c(984000,129858,139133,703025), # All Strd, EqWgt, MinVol, Mmt
                  704866)

for (id in index_code) {
  msci_eod_dl(idx_code = id)
}
