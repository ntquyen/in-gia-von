library(xlsx)
library(stringr)
source("src/tinhgiavon.R")


datetime <- as.Date("30/06/2015", "%d/%m/%Y")
time_title <- "THÁNG 6/2015 - June 2015"
signature_date <- "Vũng Tàu, ngày 30/06/2015"
output_dir <- "T6 2015"

data.giavon <- read.giavon("data/T6 2015/VAS Accounting Book June 15.xlsx", 9, 1823)

data.vtu.1540001 <- read.vtu("data/T6 2015/TK 154 Jun 2015.xlsx", "1540001", 11, 3417)
data.vtu.1540002 <- read.vtu("data/T6 2015/TK 154 Jun 2015.xlsx", "1540002", 11, 2178)

run.giavon(output_dir, datetime, data.giavon, data.vtu.1540001, data.vtu.1540002)


