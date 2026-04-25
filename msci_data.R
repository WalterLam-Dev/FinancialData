#############################
###### MSCI FUNCTIONS #######
#############################

library(quantmod)
library(readxl)
library(lubridate)

# The script will download data from MSCI's end of day history app, example as below:
# https://app2.msci.com/products/service/index/indexmaster/downloadLevelData?output=INDEX_LEVELS&currency_symbol=USD&index_variant=STRD&start_date=20201106&end_date=20241106&data_frequency=DAILY&baseValue=false&index_codes=704866,704865

# setwd("C:/Users/cwlam/Documents/R Projects")
setwd(dirname(rstudioapi::getSourceEditorContext()$path))
options(HTTPUserAgent = paste0("Mozilla/5.0 (Windows NT 10.0; Win64; x64) ", 
                               "AppleWebKit/537.36 (KHTML, like Gecko) ", 
                               "Chrome/130.0.0.0 Safari/537.36 Edg/130.0.0.0"))

# Download data
msci_eod_dl <- function(idx_code = 704866, 
                        appendMode = TRUE, 
                        output = "INDEX_LEVELS", 
                        currency = "USD", 
                        idx_var = "STRD", 
                        start_dt = "20000103", # 19691231 fails, it cannot be later than launch date
                        end_dt = format(Sys.Date(), format = "%Y%m%d"), 
                        freq = "DAILY", 
                        basevalue = "false"
                       ) {
  # Transform the list of index id
  if (length(idx_code) > 1) {
    idx_code = paste(idx_code, collapse = ",") 
  }
  
  # Put the date as Friday if it is Sat/Sun
  if (end_dt == format(Sys.Date(), format = "%Y%m%d")) {
    if (wday(Sys.Date()) == 7) { # 7: Saturday
      end_dt = format(Sys.Date()-1, format = "%Y%m%d")
      } else if (wday(Sys.Date()) == 1) { # 1: Sunday
        end_dt = format(Sys.Date()-2, format = "%Y%m%d")
      }
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
  print(URL)
  
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
  df = data.frame(df)
  df$Date = mdy(df$Date)
  df$Date <- as.Date(df$Date, format = "%m/%d/%Y")
  # Remove the comma in the numbers and read as numeric
  for (i in 2:ncol(df)) {
    df[, i] = gsub(",", "", df[, i])
    df[, i] = as.numeric(df[, i])
  }
  
  # Append the new data to the original .csv file
  if (appendMode == TRUE && file.exists(paste0(file_name, ".csv"))) { 
    original_df = read_csv(paste0(file_name, ".csv"), show_col_types = FALSE)
    # Read the date column according to sys/excel date format
    original_df$Date <- as.Date(original_df$Date, format = "%d/%m/%Y")
    
    # Append the new data if it is not in the file saved before
    df = df[!df$Date %in% original_df$Date, ]
    df = rbind(original_df, df)
    
    ### Rearrange the rows by date in ascending order
    df <- df[order(df$Date), ]
  }
  
  # Save the file as csv
  write.csv(df, paste0(file_name, ".csv"), row.names = FALSE)
}

# To check the index code:
# 1. Select the indices from the MSCI website (https://www.msci.com/end-of-day-data-search)
# 2. Go to Developer (Ctrl + Shift + I) -> Network
# 3. Click the "Download Data" button in the website
# 4. Check the name of the new item after clicking the button 
msci_idx_code = list(c(704866,704865), # Cyclical, Defensive
                     c(105825,105826), # Growth, Value
                     891800, # MSCI Emerging Markets Index
                     105825 # MSCI USA Growth Index
                     )

for (id in msci_idx_code) {
  msci_eod_dl(idx_code = id)
}
