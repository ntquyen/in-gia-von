library(xlsx)
library(stringr)
source("src/tinhgiavon.R")


datetime <- as.Date("31/08/2015", "%d/%m/%Y")
time_title <- "THÁNG 8/2015 - August 2015"
signature_date <- "Vũng Tàu, ngày 31/08/2015"
output_dir <- "T8 2015"

data.giavon <- read.giavon("data/T8 2015/VAS Accounting Book Aug 15.xlsx", 9, 1587)

data.vtu.1540001 <- read.vtu("data/T8 2015/TK 154 Aug 2015.xlsx", "1540001", 11, 744)
data.vtu.1540002 <- read.vtu("data/T8 2015/TK 154 Aug 2015.xlsx", "1540002", 11, 563)

run.giavon(output_dir, datetime, data.giavon, data.vtu.1540001, data.vtu.1540002)


