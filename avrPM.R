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
#'
#' 1-hour average is calculated as follows: 0 minute 10 seconds of the previous hour --> 0 minute 0 seconds of next hour
#' For example: 1-hour average at 6 o'clock = average from 5:00:10 to 6:00:00
#'
#' 24-hour average is calculated as follows: 0 hour 0 minute 10 seconds of today --> 0 hour 0 minute 0 seconds of the next day
#' For example: 24-hour average of 15-08-2020 = average from 0:00:10 of 15-08-2020 to 0:00:00 of 16-08-2020
#'
#' This method is referenced from the automatic environmental quality monitoring solutions of Hanoi City and Bac Ninh Province
#' *********************************************
#' PM25 (ug/m3)
#' The PM25 data is an instantaneous value about every 10 seconds
#' Default directory
path <- "D:/GitHub/lab/" ## edit folder
#' Config
config <- read_excel(paste0(path,"configPM.xlsx"))

#' If all = FALSE then import the file according to start and end
#' If all = TRUE, import all the files in the directory
import.pm <- function(start, end, all,
                      path){ 
  if (all == TRUE) {
    file <- sort(list.files(path = paste0(path,"Input/PM25/")))
    #' First file
    df <- read_csv(paste0(path,"Input/PM25/",file[1]), col_names = FALSE)
    df <- df %>% dplyr::mutate(date = paste(df$X2,df$X3))
    df$date <-  as.POSIXct(strptime(df$date, format = "%Y-%m-%d %H:%M:%S", tz = "Asia/Ho_Chi_Minh"))
    #' Set limits
    df$X4[df$X4 <= config$limitpm[1] | df$X4 >= config$limitpm[2] | is.na(df$X4)] <- NA
    df$X5[df$X5 <= config$limitpm[1] | df$X5 >= config$limitpm[2] | is.na(df$X5)] <- NA
    df$X6[df$X6 <= config$limitpm[1] | df$X6 >= config$limitpm[2] | is.na(df$X6)] <- NA
    df$X7[df$X7 <= config$limitpm[1] | df$X7 >= config$limitpm[2] | is.na(df$X7)] <- NA
    #' Average PM25 = median
    df$PM25 <- apply(df[,c("X4","X5","X6","X7")],1,median, na.rm=TRUE)
    df$PM25[is.nan(df$PM25)] <- NA
    df <- df[,c("date","PM25")]
    pm <- df
    for (i in length(file):2) {
      df <- read_csv(paste0(path,"Input/PM25/",file[i]), col_names = FALSE)
      df <- df %>% dplyr::mutate(date = paste(df$X2,df$X3))
      df$date <-  as.POSIXct(strptime(df$date, format = "%Y-%m-%d %H:%M:%S", tz = "Asia/Ho_Chi_Minh"))
      df$X4[df$X4 <= config$limitpm[1] | df$X4 >= config$limitpm[2] | is.na(df$X4)] <- NA
      df$X5[df$X5 <= config$limitpm[1] | df$X5 >= config$limitpm[2] | is.na(df$X5)] <- NA
      df$X6[df$X6 <= config$limitpm[1] | df$X6 >= config$limitpm[2] | is.na(df$X6)] <- NA
      df$X7[df$X7 <= config$limitpm[1] | df$X7 >= config$limitpm[2] | is.na(df$X7)] <- NA
      df$PM25 <- apply(df[,c("X4","X5","X6","X7")],1,median, na.rm=TRUE)
      df$PM25[is.nan(df$PM25)] <- NA
      df <- df[,c("date","PM25")]
      pm <- dplyr::union(pm,df)
    }
    return(pm)
  }
  else {
    start.row <- paste0("PM25_",start,".CSV")
    end.row <- paste0("PM25_",end,".CSV")
    file <- sort(list.files(path = paste0(path,"Input/PM25/")))
    file <- file[which(file == start.row):which(file == end.row)]
    df <- read_csv(paste0(path,"Input/PM25/",file[1]), col_names = FALSE)
    df <- df %>% dplyr::mutate(date = paste(df$X2,df$X3))
    df$date <-  as.POSIXct(strptime(df$date, format = "%Y-%m-%d %H:%M:%S", tz = "Asia/Ho_Chi_Minh"))
    df$X4[df$X4 <= config$limitpm[1] | df$X4 >= config$limitpm[2] | is.na(df$X4)] <- NA
    df$X5[df$X5 <= config$limitpm[1] | df$X5 >= config$limitpm[2] | is.na(df$X5)] <- NA
    df$X6[df$X6 <= config$limitpm[1] | df$X6 >= config$limitpm[2] | is.na(df$X6)] <- NA
    df$X7[df$X7 <= config$limitpm[1] | df$X7 >= config$limitpm[2] | is.na(df$X7)] <- NA
    df$PM25 <- apply(df[,c("X4","X5","X6","X7")],1,mean, na.rm=TRUE)
    df$PM25[is.nan(df$PM25)] <- NA
    df <- df[,c("date","PM25")]
    pm <- df
    for (i in length(file):2) {
      df <- read_csv(paste0(path,"Input/PM25/",file[i]), col_names = FALSE)
      df <- df %>% dplyr::mutate(date = paste(df$X2,df$X3))
      df$date <-  as.POSIXct(strptime(df$date, format = "%Y-%m-%d %H:%M:%S", tz = "Asia/Ho_Chi_Minh"))
      df$X4[df$X4 <= config$limitpm[1] | df$X4 >= config$limitpm[2] | is.na(df$X4)] <- NA
      df$X5[df$X5 <= config$limitpm[1] | df$X5 >= config$limitpm[2] | is.na(df$X5)] <- NA
      df$X6[df$X6 <= config$limitpm[1] | df$X6 >= config$limitpm[2] | is.na(df$X6)] <- NA
      df$X7[df$X7 <= config$limitpm[1] | df$X7 >= config$limitpm[2] | is.na(df$X7)] <- NA
      df$PM25 <- apply(df[,c("X4","X5","X6","X7")],1,mean, na.rm=TRUE)
      df$PM25[is.nan(df$PM25)] <- NA
      df <- df[,c("date","PM25")]
      pm <- dplyr::union(pm,df)
    }
    return(pm)
  }
}
pm <- import.pm(start = config$startpm[1], end = config$endpm[1], all = config$allpm[1], path = path) ## edit folder

#' Prepare date
date <- c(NA, pm$date[-length(pm$date)])
pm$date <- date
pm <- pm %>% dplyr::filter(!is.na(date))
pm$date <- as.POSIXct(pm$date, origin = "1970-01-01", tz = "Asia/Ho_Chi_Minh")

#' Average hour
pm.hour <- openair::timeAverage(pm, avg.time = "hour", statistic = "mean")
date <- pm.hour$date + 3600
pm.hour$date <- date
pm.hour$PM25[is.nan(pm.hour$PM25)] <- NA
pm.hour <- pm.hour[-c(1,length(pm.hour$date)),]

#' Average day
pm.day <- openair::timeAverage(pm, avg.time = "day", statistic = "mean")
pm.day$PM25[is.nan(pm.day$PM25)] <- NA
pm.day <- pm.day[-c(1,length(pm.day$date)),]

#' Tao workbook excel
wb <- createWorkbook()
addWorksheet(wb, "PM25HOUR")
addWorksheet(wb, "PM25DAY")
# Format excel
options("openxlsx.borderColour" = "#4F80BD")
hs.hearder <- createStyle(textDecoration = "BOLD", fontColour = "#FFFFFF", fgFill = "#4F80BD", border = c("top", "bottom", "left", "right"), borderColour = "#FFFFFF", halign = "center", valign = "center")

# Save file excel
writeData(wb, "PM25HOUR", pm.hour, colNames = TRUE, borders = "all", headerStyle = hs.hearder)
writeData(wb, "PM25DAY", pm.day, colNames = TRUE, borders = "all", headerStyle = hs.hearder)
saveWorkbook(wb, file = paste0(path,"Output/", "PM25.xlsx"), overwrite = TRUE)
# Quit and save
q(save = "yes")