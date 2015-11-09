library(xlsx)
library(stringr)
source("src/tinhgiavon.R")


datetime <- as.Date("30/09/2015", "%d/%m/%Y")
time_title <- "THÁNG 9/2015 - September 2015"
signature_date <- "Vũng Tàu, ngày 30/09/2015"
output_dir <- "T9 2015"

data.giavon <- read.giavon("data/T9 2015/VAS Accounting Book Sep 15.xlsx", 9, 1218)

data.vtu.1540001 <- read.vtu("data/T9 2015/TK 154 Sep 2015.xlsx", "1540001", 11, 1032)
data.vtu.1540002 <- read.vtu("data/T9 2015/TK 154 Sep 2015.xlsx", "1540002", 11, 733)

run.giavon(output_dir, datetime, data.giavon, data.vtu.1540001, data.vtu.1540002)


