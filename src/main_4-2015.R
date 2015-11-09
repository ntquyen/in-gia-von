library(xlsx)
library(stringr)
source("src/tinhgiavon.R")


datetime <- as.Date("30/04/2015", "%d/%m/%Y")
time_title <- "THÁNG 4/2015 - April 2014"
signature_date <- "Vũng Tàu, ngày 30/04/2015"
output_dir <- "T4 2015"

data.giavon <- read.giavon("data/T4 2015/VAS Accounting Book Apr 15.xlsx", 9, 1497)

data.vtu.1540001 <- read.vtu("data/T4 2015/TK 154 Apr 2015.xlsx", "1540001", 11, 2745)
data.vtu.1540002 <- read.vtu("data/T4 2015/TK 154 Apr 2015.xlsx", "1540002", 11, 1881)

run.giavon(output_dir, datetime, data.giavon, data.vtu.1540001, data.vtu.1540002)