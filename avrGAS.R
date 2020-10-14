#' Author: Ich Hung Ngo
#' Gmail: hung.ni96hn@gmail.com
#' ***********************************
library(readr)
library(readxl)
library(plyr)
library(dplyr)
library(lubridate)
library(openair)
library(openxlsx)
#' ***********************************
#' This file is used to calculate 1-hour average and 24-hour average (1-day) for gas data
#' 1-hour average is calculated as follows: 0 minute 10 seconds of the previous hour --> 0 minute 0 seconds of next hour
#' For example: 1-hour average at 6 o'clock = average from 5:00:10 to 6:00:00

#' 24-hour average is calculated as follows: 0 hour 0 minute 10 seconds of today --> 0 hour 0 minute 0 seconds of the next day
#' For example: 24-hour average of 15-08-2020 = average from 0:00:10 of 15-08-2020 to 0:00:00 of 16-08-2020
#' This method is referenced from the automatic environmental quality monitoring solutions of Hanoi City and Bac Ninh Province
#' *********************************************
#' CO (ppm), NO-NO2-NOX (ppb), O3 (ppb)
#' If all = FALSE then import the file according to start & end
#' If all = TRUE, import all the files in the directory
#' The GAS data is an instantaneous value about every 10 seconds
#' Default directory
path <- "D:/GitHub/lab/" ## edit folder
#' Config
config <- read_excel(paste0(path,"configGAS.xlsx"))
import.gas <- function(start, end, all,
                       path){
  if (all == TRUE) {
    file <- sort(list.files(path = paste0(path,"Input/GAS/")))
    df <- read_csv(paste0(path,"Input/GAS/",file[1]))
    df <- df %>% dplyr::mutate(date = (TheTime - 25569.2916666667)*86400,
                               NO = CH1*200 - 0.4,
                               NO2 = CH2*200 - 0.4,
                               NOX = CH3*200 - 0.4,
                               O3 = CH6*200,
                               CO = (CH7*5-0.45)/0.78)
    df <- df[,c("date", "NO", "NO2", "NOX", "O3", "CO")]
    gas <- df
    for (i in length(file):2) {
      df <- read_csv(paste0(path,"Input/GAS/",file[i]))
      df <- df %>% dplyr::mutate(date = (TheTime - 25569.2916666667)*86400,
                                 NO = CH1*200 - 0.4,
                                 NO2 = CH2*200 - 0.4,
                                 NOX = CH3*200 - 0.4,
                                 O3 = CH6*200,
                                 CO = (CH7*5-0.45)/0.78)
      df <- df[,c("date", "NO", "NO2", "NOX", "O3", "CO")]
      gas <- dplyr::union(gas,df)
    }
    return(gas)
  }
  else {
    start.row <- paste0(start,".csv")
    end.row <- paste0(end,".csv")
    file <- sort(list.files(path = paste0(path,"Input/GAS/")))
    file <- file[which(file == start.row):which(file == end.row)]
    df <- read_csv(paste0(path,"Input/GAS/",file[1]))
    df <- df %>% dplyr::mutate(date = (TheTime - 25569.2916666667)*86400,
                               NO = CH1*200 - 0.4,
                               NO2 = CH2*200 - 0.4,
                               NOX = CH3*200 - 0.4,
                               O3 = CH6*200,
                               CO = (CH7*5-0.45)/0.78)
    df <- df[,c("date", "NO", "NO2", "NOX", "O3", "CO")]
    gas <- df
    for (i in length(file):2) {
      df <- read_csv(paste0(path,"Input/GAS/",file[i]))
      df <- df %>% dplyr::mutate(date = (TheTime - 25569.2916666667)*86400,
                                 NO = CH1*200 - 0.4,
                                 NO2 = CH2*200 - 0.4,
                                 NOX = CH3*200 - 0.4,
                                 O3 = CH6*200,
                                 CO = (CH7*5-0.45)/0.78)
      df <- df[,c("date", "NO", "NO2", "NOX", "O3", "CO")]
      gas <- dplyr::union(gas,df)
    }
    return(gas)
  }
}
gas <- import.gas(start = config$startgas[1], end = config$endgas[1], all = config$allgas[1], path = path) ## edit folder
#' Validate data
#' Set lower limit and upper limit
gas$NO[gas$NO <= config$limitnox[1] | gas$NO >= config$limitnox[2] | is.na(gas$NO)] <- NA
gas$NO2[gas$NO2 <= config$limitnox[1] | gas$NO2 >= config$limitnox[2] | is.na(gas$NO2)] <- NA
gas$NOX[gas$NOX <= config$limitnox[1] | gas$NOX >= config$limitnox[2] | is.na(gas$NOX)] <- NA
gas$O3[gas$O3 <= config$limito3[1] | gas$O3 >= config$limito3[2] | is.na(gas$O3)] <- NA
gas$CO[gas$CO <= config$limitco[1] | gas$CO >= config$limitco[2] | is.na(gas$CO)] <- NA
#' R-openair calculates an hourly average based on the number of hour: for example R calculates the average at 6 o'clock from 6:00:00 to 6:59:59,
#' Similar to 24-hour, R will count for all hours of that date
#' This makes a difference with respect to the time of observation
date <- c(NA, gas$date[-length(gas$date)])
gas$date <- date
gas <- gas %>% dplyr::filter(!is.na(date))
gas$date <- as.POSIXct(gas$date, origin = "1970-01-01", tz = "Asia/Ho_Chi_Minh")
#' Average hour
avg.hour <- openair::timeAverage(gas, avg.time = "hour", statistic = "mean")
date <- avg.hour$date + 3600
avg.hour$date <- date
avg.hour$NO[is.nan(avg.hour$NO)] <- NA
avg.hour$NO2[is.nan(avg.hour$NO2)] <- NA
avg.hour$NOX[is.nan(avg.hour$NOX)] <- NA
avg.hour$O3[is.nan(avg.hour$O3)] <- NA
avg.hour$CO[is.nan(avg.hour$CO)] <- NA
avg.hour <- avg.hour[-c(1,length(avg.hour$date)),]
#' Average day
avg.day <- openair::timeAverage(gas, avg.time = "day", statistic = "mean")
avg.day$NO[is.nan(avg.day$NO)] <- NA
avg.day$NO2[is.nan(avg.day$NO2)] <- NA
avg.day$NOX[is.nan(avg.day$NOX)] <- NA
avg.day$O3[is.nan(avg.day$O3)] <- NA
avg.day$CO[is.nan(avg.day$CO)] <- NA
avg.day <- avg.day[-c(1,length(avg.day$date)),]
#' Tao workbook excel
wb <- createWorkbook()
addWorksheet(wb, "GASHOUR")
addWorksheet(wb, "GASDAY")
# Format excel
options("openxlsx.borderColour" = "#4F80BD")
hs.hearder <- createStyle(textDecoration = "BOLD", fontColour = "#FFFFFF", fgFill = "#4F80BD", border = c("top", "bottom", "left", "right"), borderColour = "#FFFFFF", halign = "center", valign = "center")

# Save file excel
writeData(wb, "GASHOUR", avg.hour, colNames = TRUE, borders = "all", headerStyle = hs.hearder)
writeData(wb, "GASDAY", avg.day, colNames = TRUE, borders = "all", headerStyle = hs.hearder)
saveWorkbook(wb, file = paste0(path,"Output/", "GAS.xlsx"), overwrite = TRUE)
# Quit and save
q(save = "yes")
