library(devtools)
library(curl)
library(RCurl)
library(rga)
library(xlsx)

rga.open(instance = "ga")

# assign gaid, ASUS Themes (UA-52407334-12)
gaid_ori <- "103299109"
#assign gaid, ASUS Themes user activity (UA-52407334-19)
gaid_10 <- "121989812"

segment_new_mechanism <- "sessions::condition::ga:appVersion[]1.6.0.38_161017|1.6.0.39_161027|1.6.0.42_161122|1.6.0.46_161209|1.6.0.52_161227|1.6.0.56_170103|1.6.0.58_170117|1.6.0.59_170120|1.6.0.60_170222"

#gaid_ori
##201610
gaid_ori_data_201610 <- ga$getData(gaid_ori, start.date = as.Date("2016-10-01"), end.date=as.Date("2016-10-31"), metrics = "ga:sessions, ga:screenviews, ga:screenviewsPerSession, ga:avgSessionDuration", dimensions = "ga:date")

gaid_ori_mau_total_201610 <- ga$getData(gaid_ori, start.date = as.Date("2016-10-31"), end.date=as.Date("2016-10-31"), metrics = "ga:30dayUsers", dimensions = "ga:date")

#201611
gaid_ori_data_201611 <- ga$getData(gaid_ori, start.date = as.Date("2016-11-01"), end.date=as.Date("2016-11-30"), metrics = "ga:sessions, ga:screenviews, ga:screenviewsPerSession, ga:avgSessionDuration", dimensions = "ga:date")

gaid_ori_mau_total_201611 <- ga$getData(gaid_ori, start.date = as.Date("2016-11-30"), end.date=as.Date("2016-11-30"), metrics = "ga:30dayUsers", dimensions = "ga:date")

#201612
gaid_ori_data_201612 <- ga$getData(gaid_ori, start.date = as.Date("2016-12-01"), end.date=as.Date("2016-12-31"), metrics = "ga:sessions, ga:screenviews, ga:screenviewsPerSession, ga:avgSessionDuration", dimensions = "ga:date")

gaid_ori_mau_total_201612 <- ga$getData(gaid_ori, start.date = as.Date("2016-12-31"), end.date=as.Date("2016-12-31"), metrics = "ga:30dayUsers", dimensions = "ga:date")

#201701
gaid_ori_data_201701 <- ga$getData(gaid_ori, start.date = as.Date("2017-01-01"), end.date=as.Date("2017-01-31"), metrics = "ga:sessions, ga:screenviews, ga:screenviewsPerSession, ga:avgSessionDuration", dimensions = "ga:date")

gaid_ori_mau_total_201701 <- ga$getData(gaid_ori, start.date = as.Date("2017-01-31"), end.date=as.Date("2017-01-31"), metrics = "ga:30dayUsers", dimensions = "ga:date")

#201702
gaid_ori_data_201702 <- ga$getData(gaid_ori, start.date = as.Date("2017-02-01"), end.date=as.Date("2017-02-28"), metrics = "ga:sessions, ga:screenviews, ga:screenviewsPerSession, ga:avgSessionDuration", dimensions = "ga:date")

gaid_ori_mau_total_201702 <- ga$getData(gaid_ori, start.date = as.Date("2017-02-28"), end.date=as.Date("2017-02-28"), metrics = "ga:30dayUsers", dimensions = "ga:date")

#gaid_10
#201610
gaid_10_data_201610 <- ga$getData(gaid_10, start.date = as.Date("2016-10-01"), end.date=as.Date("2016-10-31"), metrics = "ga:sessions, ga:screenviews, ga:screenviewsPerSession, ga:avgSessionDuration", dimensions = "ga:date")

gaid_10_mau_total_201610 <- ga$getData(gaid_10, start.date = as.Date("2016-10-31"), end.date=as.Date("2016-10-31"), metrics = "ga:30dayUsers", dimensions = "ga:date")

gaid_10_mau_new_mechanism_201610 <- ga$getData(gaid_10, start.date = as.Date("2016-10-31"), end.date=as.Date("2016-10-31"), metrics = "ga:30dayUsers", dimensions = "ga:date", segment = segment_new_mechanism)

#201611
gaid_10_data_201611 <- ga$getData(gaid_10, start.date = as.Date("2016-11-01"), end.date=as.Date("2016-11-30"), metrics = "ga:sessions, ga:screenviews, ga:screenviewsPerSession, ga:avgSessionDuration", dimensions = "ga:date")

gaid_10_mau_total_201611 <- ga$getData(gaid_10, start.date = as.Date("2016-11-30"), end.date=as.Date("2016-11-30"), metrics = "ga:30dayUsers", dimensions = "ga:date")

gaid_10_mau_new_mechanism_201611 <- ga$getData(gaid_10, start.date = as.Date("2016-11-30"), end.date=as.Date("2016-11-30"), metrics = "ga:30dayUsers", dimensions = "ga:date", segment = segment_new_mechanism)

#201612
gaid_10_data_201612 <- ga$getData(gaid_10, start.date = as.Date("2016-12-01"), end.date=as.Date("2016-12-31"), metrics = "ga:sessions, ga:screenviews, ga:screenviewsPerSession, ga:avgSessionDuration", dimensions = "ga:date")

gaid_10_mau_total_201612 <- ga$getData(gaid_10, start.date = as.Date("2016-12-31"), end.date=as.Date("2016-12-31"), metrics = "ga:30dayUsers", dimensions = "ga:date")

gaid_10_mau_new_mechanism_201612 <- ga$getData(gaid_10, start.date = as.Date("2016-12-31"), end.date=as.Date("2016-12-31"), metrics = "ga:30dayUsers", dimensions = "ga:date", segment = segment_new_mechanism)

#201701
gaid_10_data_201701 <- ga$getData(gaid_10, start.date = as.Date("2017-01-01"), end.date=as.Date("2017-01-31"), metrics = "ga:sessions, ga:screenviews, ga:screenviewsPerSession, ga:avgSessionDuration", dimensions = "ga:date")

gaid_10_mau_total_201701 <- ga$getData(gaid_10, start.date = as.Date("2017-01-31"), end.date=as.Date("2017-01-31"), metrics = "ga:30dayUsers", dimensions = "ga:date")

gaid_10_mau_new_mechanism_201701 <- ga$getData(gaid_10, start.date = as.Date("2017-01-31"), end.date=as.Date("2017-01-31"), metrics = "ga:30dayUsers", dimensions = "ga:date", segment = segment_new_mechanism)

#201702
gaid_10_data_201702 <- ga$getData(gaid_10, start.date = as.Date("2017-02-01"), end.date=as.Date("2017-02-28"), metrics = "ga:sessions, ga:screenviews, ga:screenviewsPerSession, ga:avgSessionDuration", dimensions = "ga:date")

gaid_10_mau_total_201702 <- ga$getData(gaid_10, start.date = as.Date("2017-02-28"), end.date=as.Date("2017-02-28"), metrics = "ga:30dayUsers", dimensions = "ga:date")

gaid_10_mau_new_mechanism_201702 <- ga$getData(gaid_10, start.date = as.Date("2017-02-28"), end.date=as.Date("2017-02-28"), metrics = "ga:30dayUsers", dimensions = "ga:date", segment = segment_new_mechanism)

#create a new workbook
wb <- createWorkbook()
#create a sheet named Data
sheet1 <- createSheet(wb, sheetName = "gaid_10_data")
sheet2 <- createSheet(wb, sheetName = "gaid_10_mau_total")
sheet3 <- createSheet(wb, sheetName = "gaid_10_mau_new_mechanism")
sheet4 <- createSheet(wb, sheetName = "gaid_ori_data")
sheet5 <- createSheet(wb, sheetName = "gaid_ori_mau_total")

#gaid_ori
##combine gaData_sessions
gaid_ori_data <- rbind(gaid_ori_data_201610, gaid_ori_data_201611, gaid_ori_data_201612, gaid_ori_data_201701, gaid_ori_data_201702)

#combine mau_total
gaid_ori_mau_total <- rbind(gaid_ori_mau_total_201610, gaid_ori_mau_total_201611, gaid_ori_mau_total_201612, gaid_ori_mau_total_201701, gaid_ori_mau_total_201702)

#gaid_10
#combine gaData_sessions
gaid_10_data <- rbind(gaid_10_data_201610, gaid_10_data_201611, gaid_10_data_201612, gaid_10_data_201701, gaid_10_data_201702)

#combine mau_total
gaid_10_mau_total <- rbind(gaid_10_mau_total_201610, gaid_10_mau_total_201611, gaid_10_mau_total_201612, gaid_10_mau_total_201701, gaid_10_mau_total_201702)

#combine mau_new_mechanism
gaid_10_mau_new_mechanism <- rbind(gaid_10_mau_new_mechanism_201610, gaid_10_mau_new_mechanism_201611, gaid_10_mau_new_mechanism_201612, gaid_10_mau_new_mechanism_201701, gaid_10_mau_new_mechanism_201702)

#add your data to the Data sheet
addDataFrame(gaid_10_data, sheet1, row.names = FALSE)
addDataFrame(gaid_10_mau_total, sheet2, row.names = FALSE)
addDataFrame(gaid_10_mau_new_mechanism, sheet3, row.names = FALSE)
addDataFrame(gaid_ori_data, sheet4, row.names = FALSE)
addDataFrame(gaid_ori_mau_total, sheet5, row.names = FALSE)
#don't forget to save your workbook
#this will save to your working directory
saveWorkbook(wb, "ThemeGADataAnalytics.xlsx")
