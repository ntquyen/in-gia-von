library(xlsx)
library(stringr)
source("src/tinhgiavon.R")


datetime <- as.Date("31/07/2015", "%d/%m/%Y")
time_title <- "THÁNG 7/2015 - July 2015"
signature_date <- "Vũng Tàu, ngày 31/07/2015"
output_dir <- "T7 2015"

data.giavon <- read.giavon("data/T7 2015/VAS Accounting Book July 15.xlsx", 9, 1302)

data.vtu.1540001 <- read.vtu("data/T7 2015/TK 154 July 2015.xlsx", "1540001", 11, 375)
data.vtu.1540002 <- read.vtu("data/T7 2015/TK 154 July 2015.xlsx", "1540002", 11, 450)

run.giavon(output_dir, datetime, data.giavon, data.vtu.1540001, data.vtu.1540002)


