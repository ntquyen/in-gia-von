library(xlsx)
library(stringr)
source("src/tinhgiavon.R")


datetime <- as.Date("31/05/2015", "%d/%m/%Y")
time_title <- "THÁNG 5/2015 - May 2015"
signature_date <- "Vũng Tàu, ngày 31/05/2015"
output_dir <- "T5 2015"

data.giavon <- read.giavon("data/T5 2015/VAS Accounting Book May 15.xlsx", 9, 1442)

data.vtu.1540001 <- read.vtu("data/T5 2015/TK 154 May 2015.xlsx", "1540001", 11, 3093)
data.vtu.1540002 <- read.vtu("data/T5 2015/TK 154 May 2015.xlsx", "1540002", 11, 2009)

run.giavon(output_dir, datetime, data.giavon, data.vtu.1540001, data.vtu.1540002)
