library(xlsx)
library(stringr)
source("src/tinhgiavon.R")

# ## T9 2014
# datetime <- as.Date("30/09/2014", "%d/%m/%Y")
# time_title <- "THÁNG 9/2014 - SEPTEMBER 2014"
# signature_date <- "Vũng Tàu, ngày 30/09/2014"
# output_dir <- "T9 2014"
# 
# data.giavon <- read.giavon("data/T9 2014/VAS Accounting Book Sep 14.xlsx", 9, 1361)  
# data.vtu.1540001 <- read.vtu("data/T9 2014/TK 154 Sep 2014.xlsx", "1540001", 11, 879)
# data.vtu.1540002 <- read.vtu("data/T9 2014/TK 154 Sep 2014.xlsx", "1540002", 11, 918)
# 
# run.giavon(output_dir, datetime, data.giavon, data.vtu.1540001, data.vtu.1540002)
# 
# ## T10 2014
# datetime <- as.Date("31/10/2014", "%d/%m/%Y")
# time_title <- "THÁNG 10/2014 - OCTOBER 2014"
# signature_date <- "Vũng Tàu, ngày 31/10/2014"
# output_dir <- "T10 2014"
# 
# data.giavon <- read.giavon("data/T10 2014/VAS Accounting Book Oct 14.xlsx", 9, 1509)
# data.vtu.1540001 <- read.vtu("data/T10 2014/TK 154 Oct 2014.xlsx", "1540001", 11, 1186)
# data.vtu.1540002 <- read.vtu("data/T10 2014/TK 154 Oct 2014.xlsx", "1540002", 11, 1084)
# run.giavon(output_dir, datetime, data.giavon, data.vtu.1540001, data.vtu.1540002)
# 
# ## T11 2014
# datetime <- as.Date("30/11/2014", "%d/%m/%Y")
# data.giavon <- read.giavon("data/T11 2014/VAS Accounting Book Nov 14.xlsx", 9, 1446)
# 
# time_title <- "THÁNG 11/2014 - NOVEMBER 2014"
# signature_date <- "Vũng Tàu, ngày 30/11/2014"
# output_dir <- "T11 2014"
# 
# data.vtu.1540001 <- read.vtu("data/T11 2014/TK 154 Nov 2014.xlsx", "1540001", 11, 1444)
# data.vtu.1540002 <- read.vtu("data/T11 2014/TK 154 Nov 2014.xlsx", "1540002", 11, 1257)
# run.giavon(output_dir, datetime, data.giavon, data.vtu.1540001, data.vtu.1540002)
# 
# ## T12 2014
# datetime <- as.Date("31/12/2014", "%d/%m/%Y")
# data.giavon <- read.giavon("data/T12 2014/VAS Accounting Book Dec 14.xlsx", 9, 1501)
# 
# time_title <- "THÁNG 12/2014 - December 2014"
# signature_date <- "Vũng Tàu, ngày 31/12/2014"
# output_dir <- "T12 2014"
# 
# data.vtu.1540001 <- read.vtu("data/T12 2014/TK 154 Dec 2014.xlsx", "1540001", 11, 1742)
# data.vtu.1540002 <- read.vtu("data/T12 2014/TK 154 Dec 2014.xlsx", "1540002", 11, 1385)
# 
# run.giavon(output_dir, datetime, data.giavon, data.vtu.1540001, data.vtu.1540002)
# 
# ## T1 2015
# datetime <- as.Date("31/01/2015", "%d/%m/%Y")
# data.giavon <- read.giavon("data/T1 2015/VAS Accounting Book Jan 15.xlsx", 9, 1387)
# 
# time_title <- "THÁNG 1/2015 - January 2014"
# signature_date <- "Vũng Tàu, ngày 31/01/2015"
# output_dir <- "T1 2015"
# 
# 
# data.vtu.1540001 <- read.vtu("data/T1 2015/TK 154 Jan 2015.xlsx", "1540001", 11, 1986)
# data.vtu.1540002 <- read.vtu("data/T1 2015/TK 154 Jan 2015.xlsx", "1540002", 11, 1514)
# 
# run.giavon(output_dir, datetime, data.giavon, data.vtu.1540001, data.vtu.1540002)
# 
# # T2 2015
# datetime <- as.Date("28/02/2015", "%d/%m/%Y")
# time_title <- "THÁNG 2/2015 - February 2014"
# signature_date <- "Vũng Tàu, ngày 28/02/2015"
# output_dir <- "T2 2015"
# 
# data.giavon <- read.giavon("data/T2 2015/VAS Accounting Book Feb 15.xlsx", 9, 1194)
# 
# data.vtu.1540001 <- read.vtu("data/T2 2015/TK 154 Feb 2015.xlsx", "1540001", 11, 2196)
# data.vtu.1540002 <- read.vtu("data/T2 2015/TK 154 Feb 2015.xlsx", "1540002", 11, 1601)
# 
# run.giavon(output_dir, datetime, data.giavon, data.vtu.1540001, data.vtu.1540002)

# ## T3 2015
# data.giavon <- read.giavon("T3 2015/VAS Accounting Book Mar 15.xlsx", 9, 1401)
# View(data.giavon)
# 
# time_title <- "THÁNG 3/2015 - March 2014"
# signature_date <- "Vũng Tàu, ngày 31/03/2015"
# output_dir <- "T3 2015"
# 
# data.vtu.1540001 <- read.vtu("T3 2015/TK 154 Mar 2015.xlsx", "1540001", 11, 2450)
# data.vtu.1540002 <- read.vtu("T3 2015/TK 154 Mar 2015.xlsx", "1540002", 11, 1740)
# data.vtu.1540001$No <- as.numeric(data.vtu.1540001$No)
# data.vtu.1540002$No <- as.numeric(data.vtu.1540002$No)
# 
# run.giavon(output_dir, data.giavon, data.vtu.1540001, data.vtu.1540002)


# ## T4 2015
# datetime <- as.Date("30/04/2015", "%d/%m/%Y")
# time_title <- "THÁNG 4/2015 - April 2014"
# signature_date <- "Vũng Tàu, ngày 30/04/2015"
# output_dir <- "T4 2015"
# 
# data.giavon <- read.giavon("data/T4 2015/VAS Accounting Book Apr 15.xlsx", 9, 1497)
# 
# data.vtu.1540001 <- read.vtu("data/T4 2015/TK 154 Apr 2015.xlsx", "1540001", 11, 2745)
# data.vtu.1540002 <- read.vtu("data/T4 2015/TK 154 Apr 2015.xlsx", "1540002", 11, 1881)
# 
# run.giavon(output_dir, datetime, data.giavon, data.vtu.1540001, data.vtu.1540002)