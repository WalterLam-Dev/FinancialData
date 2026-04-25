library(quantmod)
library(readxl)
library(writexl)
library(lubridate)
library(tidyverse)

# The script will download data from S&P Global end of day history, example as below:
# https://www.spglobal.com/spdji/en/indices/dividends-factors/sp-500-high-beta-index/#overview

# setwd("C:/Users/cwlam/Documents/R Projects")
setwd(dirname(rstudioapi::getSourceEditorContext()$path))
# custom_userAgent = paste0("Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) ", 
#                           "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/", 
#                           "126.0.0.0 Safari/537.36"
#                           )
# Below header can be obtained by:
# Safari -> right click in the website -> Inspect Element
# Network -> click the download button in the website
# find "collect" under "Network" -> select the tab "Headers"
# vary Accept-Encoding,User-Agent
headers = c("Accept" = "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7", 
            "Accept-Encoding" = "gzip,deflate,br,zstd",
            "accept-language" = "zh-TW,zh;q=0.9", 
            "dnt" = "1", 
            "priority" = "u=0,i", 
            # "sec-ch-ua" = "\"Not;A=Brand\";v=\"99\", \"Google Chrome\";v=\"139\", \"Chromium\";v=\"139\"", 
            "sec-ch-ua-mobile" = "?0", 
            "sec-ch-ua-platform" = "\"macOS\"", 
            "Sec-Fetch-Dest" = "document", 
            "Sec-Fetch-Mode" = "navigate", 
            "Sec-Fetch-Site" = "none", 
            "sec-fetch-user" = "?1", 
            "upgrade-insecure-requests" = "1", 
            "User-Agent" = "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/139.0.0.0 Safari/537.36"
            )

# Download data
# https://www.spglobal.com/spdji/en/idsexport/file.xls?
# hostIdentifier=48190c8c-42c4-46af-8d1a-0cd5db894797
# &redesignExport=true&languageId=1
# &selectedModule=PerformanceGraphView
# &selectedSubModule=Graph
# &yearFlag=tenYearFlag
# &indexId=5475130
spg_eod_dl <- function(idx_code = 5475130, 
                       appendMode = TRUE, 
                       yearFlag = "oneYearFlag" # oneYearFlag, fiveYearFlag, tenYearFlag, etc.
                       ) {
  
  # Construct the URL
  # Use Chrome to check
  URL = paste0("https://www.spglobal.com/spdji/en/idsexport/file.xls?", 
               "hostIdentifier=48190c8c-42c4-46af-8d1a-0cd5db894797", 
               "&redesignExport=true", 
               "&languageId=1", 
               "&selectedModule=PerformanceGraphView", 
               "&selectedSubModule=Graph", 
               "&yearFlag=", yearFlag, 
               "&indexId=", as.character(idx_code)
               )
  
  # Download the .xls file
  file_name = paste0("SPG_", idx_code)
  # Don't know why the following isn't working
  # download.file(url = URL, 
  #               destfile = "SPG_temp.xls", 
  #               mode = "ab", 
  #               quiet = FALSE, 
  #               headers = headers
  #               )
  
  # Download the file with browser
  browseURL(URL)
  # By default, the name of file downloaded is PerformanceGraphExport.xls
  original_file_path <- "/Users/WalterLam/Downloads/PerformanceGraphExport.xls"
  new_file_path <- "/Users/WalterLam/Documents/Investment/Signal/Raw Data/Data by R/SPG_temp.xls"
  while (!file.exists(original_file_path)) {
    # Print a message indicating waiting status
    # message("Waiting for file to be downloaded...")
    
    # Pause execution for 1 second
    Sys.sleep(0.1) 
  }
  file.copy(from = original_file_path, to = new_file_path, overwrite = TRUE)
  file.remove(original_file_path)
  
  ### Data cleaning and save again
  # Read the .xls file
  df = read_excel("SPG_temp.xls")
  df = na.omit(df) # remove row with any NA
  names(df) = df[1,] # put first row as column names
  df = df[-1,] # remove the first row
  
  # Change the first column as date format
  df = data.frame(df)
  df$Effective.date = as.numeric(df$Effective.date)
  df$Effective.date <- as.Date(df$Effective.date, origin = "1899-12-30")
  # Read index as numeric
  df[, -1] = as.numeric(df[, -1])
  
  # Append the new data to the original .csv file
  if (appendMode == TRUE && file.exists(paste0(file_name, ".csv"))) { 
    original_df = read_csv(paste0(file_name, ".csv"), show_col_types = FALSE)
    # Read the date column according to sys/excel date format
    original_df$Effective.date <- as.Date(original_df$Effective.date, format = "%d/%m/%Y")
    
    # Append the new data if it is not in the file saved before
    df = df[!df$Effective.date %in% original_df$Effective.date, ]
    df = rbind(original_df, df)
    
    ### Rearrange the rows by date in ascending order
    df <- df[order(df$Effective.date), ]
  }
  
  # Save the file as csv
  print(paste0("Saving ", idx_code, ": ", colnames(df)[2]))
  write.csv(df, paste0(file_name, ".csv"), row.names = FALSE)
  
  # Sys.sleep(1)
}

# To check the index code:
# 1. Select the indices from the S&P Global website
# 2. Go to Developer (Ctrl + Shift + I) -> Network
# 3. Click the "Export" button in the website
# 4. Check the name of the new item after clicking the button 
spg_idx_code = list(5475130, # S&P 500 High Beta Index
                    5475134, # S&P 500 Low Volatility Index
                    10003477, # S&P GSCI Industrial Metals
                    10003812, # S&P GSCI Precious Metals
                    2029, # S&P 500 Growth
                    92319526, # S&P U.S. High Yield Corporate Bond Index
                    92030168, # S&P 500 AAA Investment Grade Corporate Bond Index
                    92347751 # S&P U.S. Government & Corporate AAA Bond Index
                    )

for (id in spg_idx_code) {
  spg_eod_dl(idx_code = id)
}


### US State Coincident Indexes from Federal Reserve Bank of Philadelphia
# Download the file from "Revised Data" from the "Latest Data" section
# Source: https://www.philadelphiafed.org/surveys-and-data/regional-economic-analysis/state-coincident-indexes
dl_file_name = "coincident-revised.xls"
COI_url = paste0("https://www.philadelphiafed.org/-/media/FRBP/Assets/Surveys-And-Data/coincident/", 
                 dl_file_name)

# Download the file with browser
browseURL(COI_url)
COI_dl_file_path <- paste0("/Users/WalterLam/Downloads/", 
                           dl_file_name)
COI_new_file_path <- paste0("/Users/WalterLam/Documents/Investment/Signal/Raw Data/Data by R/", 
                        "coincident-revised.xls")
while (!file.exists(COI_dl_file_path)) {
  # Print a message indicating waiting status
  # message("Waiting for file to be downloaded...")
  
  # Pause execution for 1 second
  Sys.sleep(0.1) 
}

# 1. Get the names of all spreadsheets in the file
COI_sheet_names <- excel_sheets(COI_dl_file_path)

# 2. Read each spreadsheet into a list of data frames
# set_names ensures the list elements match the original sheet names
COI_all_sheets <- setNames(lapply(COI_sheet_names, 
                                  function(x) read_excel(COI_dl_file_path, sheet = x)), 
                           COI_sheet_names)
                                                                                                   
# 3. Write the entire list to a new .xlsx file
write_xlsx(COI_all_sheets, paste0("/Users/WalterLam/Documents/Investment/Signal/Raw Data/Data by R/", 
                                  "coincident-revised.xlsx"))

file.copy(from = COI_dl_file_path, to = COI_new_file_path, overwrite = TRUE)
file.remove(COI_dl_file_path)