read.giavon <- function(file, startRow, endRow) {
  data <- read.xlsx2(file, sheetName = "NKC", stringsAsFactors = FALSE,
                     startRow=startRow, endRow=endRow, colIndex=1:14)
  
  colnames(data) <- c("thang", "chungtu", "date", "no", "desc", 
                      "dr", "cr", "currency", "fx", "vnd", "note", 
                      "id", "inv", "control_number")

  data.giavon <- data[data$cr %in% c(1540001, 1540002), ]
  return(data.giavon)
}

read.vtu <- function(file, cr, startRow, endRow) {
  data <- read.xlsx2(file, sheetName = cr, stringsAsFactors = FALSE,
                     startRow = startRow, endRow = endRow, colIndex=1:13)
  
  colnames(data) <- c("Thang", "Ngay", "So", "Dien_giai", 
                      "tk", "No", "Co", "Control_Number", 
                      "Note", "Marine_Energy", "Labour_General", "External_Intragroup", "Notes")
  
  data$No <- as.numeric(data$No)
  
  data$Thang <- as.numeric(data$Thang)
  data$Thang <- as.Date(data$Thang-25569, origin="1970-01-01")
  
  data$Ngay <- as.numeric(data$Ngay)
  data$Ngay <- as.Date(data$Ngay-25569, origin="1970-01-01")
  
  return(data)
}