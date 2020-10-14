##* Author: Ich Hung Ngo
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
#' Default directory
path <- "D:/GitHub/lab/" ## edit folder
#' Merge pm & gas
PM25hour <- read_excel(paste0(path,"Output/PM25.xlsx"), sheet = "PM25HOUR")
PM25day <- read_excel(paste0(path,"Output/PM25.xlsx"), sheet = "PM25DAY")

GAShour <- read_excel(paste0(path,"Output/GAS.xlsx"), sheet = "GASHOUR")
GASday <- read_excel(paste0(path,"Output/GAS.xlsx"), sheet = "GASDAY")

avg.hour <- merge(PM25hour, GAShour, by = "date", all.x = T, all.y = T)
avg.day  <- merge(PM25day, GASday, by = "date", all.x = T, all.y = T)

##* Create sheet
wb <- createWorkbook()
addWorksheet(wb, "avgHour")
addWorksheet(wb, "avgDay")
##* Format sheet
options("openxlsx.borderColour" = "#4F80BD")
hs.hearder <- createStyle(textDecoration = "BOLD", fontColour = "#FFFFFF", fgFill = "#4F80BD", border = c("top", "bottom", "left", "right"), borderColour = "#FFFFFF", halign = "center", valign = "center")
##* Save excel file
##* If only pm then replace avg.hour = pm.hour, avg.day = pm.day
writeData(wb, "avgHour", avg.hour, colNames = TRUE, borders = "all", headerStyle = hs.hearder)
writeData(wb, "avgDay", avg.day, colNames = TRUE, borders = "all", headerStyle = hs.hearder)
saveWorkbook(wb, file = paste0(path,"Output/", "DATA.xlsx"), overwrite = TRUE)
# Quit and save
q(save = "yes")
