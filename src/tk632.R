source("src/writeLib.R")

size.column <- function (sheet) {
  setColumnWidth(sheet, 1, 10)
  setColumnWidth(sheet, c(2, 3), 11)
  setColumnWidth(sheet, 5, 12)
  setColumnWidth(sheet, 4, 44)
  setColumnWidth(sheet, c(6, 7), 15)
  setColumnWidth(sheet, 8, 23)
}

setPrintSetupTk632 <- function (sheet) {
  printSetup(sheet, copies = 1, paperSize = "A4_PAPERSIZE")
  sheet$setFitToPage(TRUE)
  sheet$setHorizontallyCenter(TRUE)
  
  sheet$setMargin(.jshort(sheet$LeftMargin), 0.3)
  sheet$setMargin(.jshort(sheet$RightMargin), 0.6)
  sheet$setMargin(.jshort(sheet$BottomMargin), 0.5)
  sheet$setMargin(.jshort(sheet$TopMargin), 0.6)
}

# 6321101: INTRAGROUP

write.tk632.energy <- function(output_dir, datetime, data, vtu, je, tk) {
  
  file <- paste("/Users/quyennt/Projects/in-gia-von/output/", output_dir, "/", je, "_tk632.xlsx", sep="");
  print(file)
  
  wb <- createWorkbook(type = "xlsx")
  sheet <- createSheet(wb, sheetName = vtu)
  setPrintSetupTk632(sheet)
  
  BORDER_STYLE_THIN <- Border(color = "Black", position = c("BOTTOM", "LEFT", "TOP", "RIGHT"))
  BORDER_STYLE_DOTTED <- Border(color = "Black", position = c("BOTTOM", "LEFT", "TOP", "RIGHT"), 
                                pen = c("BORDER_DOTTED", "BORDER_THIN", "BORDER_DOTTED", "BORDER_THIN"))
  ALIGN_RIGHT <- Alignment(h="ALIGN_RIGHT", v="VERTICAL_CENTER", wrapText = TRUE)
  ALIGN_CENTER <- Alignment(h="ALIGN_CENTER", v="VERTICAL_CENTER", wrapText = TRUE)
  ALIGN_LEFT <- Alignment(h="ALIGN_LEFT", v="VERTICAL_CENTER", wrapText = TRUE)
  
  ADDRESS_STYLE <- CellStyle(wb) + 
    Font(wb, heightInPoints=9, name = "Arial")
  TITLE_STYLE <- CellStyle(wb) + 
    Font(wb, heightInPoints=20, name = "Arial", isBold = TRUE) +
    Alignment(h="ALIGN_CENTER",  wrapText = TRUE)
  JE_STYLE <- CellStyle(wb) + 
    Font(wb, heightInPoints=12, name = "Arial", isBold = TRUE) +
    Alignment(h="ALIGN_RIGHT",  wrapText = TRUE)
  TIME_STYLE <- CellStyle(wb) + 
    Font(wb, heightInPoints=13, name = "Arial", isBold = TRUE) +
    Alignment(h="ALIGN_CENTER", wrapText = TRUE)
  
  TABLE_HEADER_STYLE <- CellStyle(wb) + 
    Font(wb, heightInPoints=12, name = "Arial", isBold = TRUE) +
    ALIGN_CENTER + BORDER_STYLE_THIN
  
  TABLE_CONTENT_STYLE_BOLD <- CellStyle(wb) + 
    Font(wb, heightInPoints=12, name = "Arial", isBold = TRUE) +
    ALIGN_LEFT + BORDER_STYLE_THIN
  
  TABLE_CONTENT_STYLE <- CellStyle(wb) + 
    Font(wb, heightInPoints=12, name = "Arial") +
    ALIGN_LEFT + BORDER_STYLE_DOTTED
  
  TABLE_SUMMARY_STYLE <- CellStyle(wb) + 
    Font(wb, heightInPoints=12, name = "Arial") + ALIGN_LEFT + BORDER_STYLE_THIN
  
  size.column(sheet)
  
  nextRowIndex <- write.address(sheet, 1, ADDRESS_STYLE)
  nextRowIndex <- xlsx.writeCell(sheet, rowIndex = nextRowIndex + 1, colIndex = 1, textStyle = TITLE_STYLE,
                 rowHeight = 1.7, mergedRegion = c(nextRowIndex + 1, nextRowIndex + 1, 1, 8),
                 texts = c(paste("GIÁ VỐN DỊCH VỤ TK", tk)))
  nextRowIndex <- xlsx.writeCell(sheet, rowIndex = nextRowIndex, colIndex = 1, textStyle = TITLE_STYLE,
                                 rowHeight = 1.7, mergedRegion = c(nextRowIndex, nextRowIndex, 1, 8),
                                 texts = c("COST OF SERVICE SOLD - ENERGY"))
  
  nextRowIndex <- xlsx.writeCell(sheet, rowIndex = nextRowIndex + 1, colIndex=1, textStyle = TITLE_STYLE,
                 rowHeight = 1.7, mergedRegion=c(nextRowIndex + 1, nextRowIndex + 1, 1, 8),
                 texts = c(vtu))
  nextRowIndex <- xlsx.writeCell(sheet, rowIndex = nextRowIndex, colIndex=1, textStyle = TITLE_STYLE,
                                 rowHeight = 1.7, texts = c(""))
  nextRowIndex <- write.tableHeader(sheet, TABLE_HEADER_STYLE, rowIndex = nextRowIndex:(nextRowIndex+1))
  
  n <- nrow(data)
  for (i in 1:n) {
    if (!is.na(data$No[i]) && data$No[i] > 0) {
      no <- format(data$No[i], big.mark=",", scientific = FALSE)
      thang <- format(data$Thang[i], format = "%b-%y")
      ngay <- format(data$Ngay[i], format = "%d/%m/%y")
      
      nextRowIndex <- xlsx.writeCell(sheet, rowIndex = nextRowIndex, colIndex=1:8, rowHeight = 3, 
                     texts = c(thang, ngay, data$So[i], data$Dien_giai[i], data$tk[i], no, data$Co[i], data$Note[i]),
                     textStyle = list(TABLE_CONTENT_STYLE + ALIGN_CENTER, 
                                      TABLE_CONTENT_STYLE + ALIGN_CENTER, 
                                      TABLE_CONTENT_STYLE + ALIGN_CENTER,
                                      TABLE_CONTENT_STYLE,
                                      TABLE_CONTENT_STYLE + ALIGN_CENTER,
                                      TABLE_CONTENT_STYLE + ALIGN_RIGHT,
                                      TABLE_CONTENT_STYLE,
                                      TABLE_CONTENT_STYLE),
                     multiStyles = TRUE)
    }
    
  }
  
  total_cost <- sum(data$No, na.rm = TRUE)
  total_cost.formatted <- format(total_cost, big.mark=",", scientific = FALSE)
  thang <- format(datetime, format = "%b-%y")
  ngay <- format(datetime, format = "%d/%m/%y")
  
  nextRowIndex <- xlsx.writeCell(sheet, rowIndex = nextRowIndex, colIndex=1:8, 
                 rowHeight = 3,
                 texts = c("", "", "", paste("Total Cost ", vtu, " - Energy", sep=""), "", total_cost.formatted, "", ""),
                 textStyle = list(TABLE_CONTENT_STYLE_BOLD, TABLE_CONTENT_STYLE_BOLD, TABLE_CONTENT_STYLE_BOLD,
                                  TABLE_CONTENT_STYLE_BOLD, TABLE_CONTENT_STYLE_BOLD,
                                  TABLE_CONTENT_STYLE_BOLD + ALIGN_RIGHT, TABLE_CONTENT_STYLE_BOLD, TABLE_CONTENT_STYLE_BOLD),
                 multiStyles = TRUE)
  
  nextRowIndex <- xlsx.writeCell(sheet, rowIndex = nextRowIndex, colIndex=1:8, 
                 rowHeight = 3,
                 texts = c(thang, ngay, "", paste("Giá vốn dịch vụ trong kỳ / Cost of Service - Energy ", vtu, sep=""), 
                           "9110000", "", total_cost.formatted, ""),
                 textStyle = list(TABLE_SUMMARY_STYLE + ALIGN_CENTER, TABLE_SUMMARY_STYLE + ALIGN_CENTER, TABLE_SUMMARY_STYLE,
                                  TABLE_SUMMARY_STYLE, TABLE_SUMMARY_STYLE + ALIGN_CENTER,
                                  TABLE_SUMMARY_STYLE, TABLE_CONTENT_STYLE_BOLD + ALIGN_RIGHT, TABLE_SUMMARY_STYLE),
                 multiStyles = TRUE)
  
  
  setPrintArea(wb, 1, 1, 8, 1, nextRowIndex)
  saveWorkbook(wb, file)
}

write.tk632.marine <- function(output_dir, datetime, data, vtu, je, tk) {
  file <- paste("/Users/quyennt/Projects/in-gia-von/output/", output_dir, "/", je, "_tk632.xlsx", sep="");
  print(file)
  
  wb<-createWorkbook(type = "xlsx")
  sheet <- createSheet(wb, sheetName = vtu)
  setPrintSetupTk632(sheet)
  
  BORDER_STYLE_THIN <- Border(color = "Black", position = c("BOTTOM", "LEFT", "TOP", "RIGHT"))
  BORDER_STYLE_DOTTED <- Border(color = "Black", position = c("BOTTOM", "LEFT", "TOP", "RIGHT"), 
                                pen = c("BORDER_DOTTED", "BORDER_THIN", "BORDER_DOTTED", "BORDER_THIN"))
  ALIGN_RIGHT <- Alignment(h="ALIGN_RIGHT", v="VERTICAL_CENTER", wrapText = TRUE)
  ALIGN_CENTER <- Alignment(h="ALIGN_CENTER", v="VERTICAL_CENTER", wrapText = TRUE)
  ALIGN_LEFT <- Alignment(h="ALIGN_LEFT", v="VERTICAL_CENTER", wrapText = TRUE)
  
  ADDRESS_STYLE <- CellStyle(wb) + 
    Font(wb, heightInPoints=9, name = "Arial")
  TITLE_STYLE <- CellStyle(wb) + 
    Font(wb, heightInPoints=20, name = "Arial", isBold = TRUE) +
    Alignment(h="ALIGN_CENTER",  wrapText = TRUE)
  JE_STYLE <- CellStyle(wb) + 
    Font(wb, heightInPoints=12, name = "Arial", isBold = TRUE) +
    Alignment(h="ALIGN_RIGHT",  wrapText = TRUE)
  TIME_STYLE <- CellStyle(wb) + 
    Font(wb, heightInPoints=13, name = "Arial", isBold = TRUE) +
    Alignment(h="ALIGN_CENTER", wrapText = TRUE)
  
  TABLE_HEADER_STYLE <- CellStyle(wb) + 
    Font(wb, heightInPoints=12, name = "Arial", isBold = TRUE) +
    ALIGN_CENTER + BORDER_STYLE_THIN
  
  TABLE_CONTENT_STYLE_BOLD <- CellStyle(wb) + 
    Font(wb, heightInPoints=12, name = "Arial", isBold = TRUE) +
    ALIGN_LEFT + BORDER_STYLE_THIN
  
  TABLE_CONTENT_STYLE <- CellStyle(wb) + 
    Font(wb, heightInPoints=12, name = "Arial") +
    ALIGN_LEFT + BORDER_STYLE_DOTTED
  
  TABLE_SUMMARY_STYLE <- CellStyle(wb) + 
    Font(wb, heightInPoints=12, name = "Arial") + ALIGN_LEFT + BORDER_STYLE_THIN
  
  size.column(sheet)
  
  nextRowIndex <- write.address(sheet, 1, ADDRESS_STYLE)
  
  title <- "COST OF SERVICE SOLD - MARINE"
  title.tk <- tk
  if (tk == "6321101") {
    title.tk <- paste(tk, "Intragroup", sep = ":")
    title <- "INTRAGROUP COST OF SERVICE SOLD - MARINE"
  }
  print(title.tk)
  nextRowIndex <- xlsx.writeCell(sheet, rowIndex = nextRowIndex + 1, colIndex=1, textStyle = TITLE_STYLE,
                 rowHeight = 1.7, mergedRegion=c(nextRowIndex + 1, nextRowIndex + 1, 1, 8),
                 texts = c(paste("GIÁ VỐN DỊCH VỤ TK", tk)))

  nextRowIndex <- xlsx.writeCell(sheet, rowIndex = nextRowIndex, colIndex=1, textStyle = TITLE_STYLE,
                 rowHeight = 1.6, mergedRegion=c(nextRowIndex, nextRowIndex, 1, 8),
                 texts = c(title))

  nextRowIndex <- xlsx.writeCell(sheet, rowIndex = nextRowIndex + 1, colIndex=1, textStyle = TITLE_STYLE,
                 rowHeight = 1.7, mergedRegion=c(nextRowIndex + 1, nextRowIndex + 1, 1, 8),
                 texts = c(vtu))
  nextRowIndex <- xlsx.writeCell(sheet, rowIndex = nextRowIndex, colIndex=1, textStyle = TITLE_STYLE,
                                 rowHeight = 1.7, texts = c(""))
  
  nextRowIndex <- write.tableHeader(sheet, TABLE_HEADER_STYLE, rowIndex = (nextRowIndex):(nextRowIndex + 1))
  
  n <- nrow(data)
  for (i in 1:n) {
    if (!is.na(data$No[i]) && data$No[i] > 0) {
      no <- format(data$No[i], big.mark=",", scientific = FALSE)
      thang <- format(data$Thang[i], format = "%b-%y")
      ngay <- format(data$Ngay[i], format = "%d/%m/%y")
      
      nextRowIndex <- xlsx.writeCell(sheet, rowIndex = nextRowIndex, colIndex=1:8, rowHeight = 3, 
                     texts = c(thang, ngay, data$So[i], data$Dien_giai[i], data$tk[i], no, "", data$Note[i]),
                     textStyle = list(TABLE_CONTENT_STYLE + ALIGN_CENTER, 
                                      TABLE_CONTENT_STYLE + ALIGN_CENTER, 
                                      TABLE_CONTENT_STYLE + ALIGN_CENTER,
                                      TABLE_CONTENT_STYLE,
                                      TABLE_CONTENT_STYLE + ALIGN_CENTER,
                                      TABLE_CONTENT_STYLE + ALIGN_RIGHT,
                                      TABLE_CONTENT_STYLE,
                                      TABLE_CONTENT_STYLE),
                     multiStyles = TRUE)
    }
    
  }
  
  total_cost <- sum(data$No, na.rm = TRUE)
  total_cost.formatted <- format(total_cost, big.mark=",", scientific = FALSE)
  thang <- format(datetime, format = "%b-%y")
  ngay <- format(datetime, format = "%d/%m/%y")
  nextRowIndex <- xlsx.writeCell(sheet, rowIndex = nextRowIndex, colIndex=1:8, 
                 rowHeight = 3,
                 texts = c("", "", "", paste("Total Cost ", vtu, " - Marine", sep=""), "", total_cost.formatted, "", ""),
                 textStyle = list(TABLE_CONTENT_STYLE_BOLD, TABLE_CONTENT_STYLE_BOLD, TABLE_CONTENT_STYLE_BOLD,
                                  TABLE_CONTENT_STYLE_BOLD, TABLE_CONTENT_STYLE_BOLD,
                                  TABLE_CONTENT_STYLE_BOLD + ALIGN_RIGHT, TABLE_CONTENT_STYLE_BOLD, TABLE_CONTENT_STYLE_BOLD),
                 multiStyles = TRUE)
  
  title <- paste("Giá vốn dịch vụ trong kỳ / Cost of Service - Marine ", vtu, sep="")
  if (tk == "6321101") {
    title <- paste("Giá vốn dịch vụ trong kỳ Intragroup Cost of Service - Marine ", vtu, sep="")
  }
  nextRowIndex <- xlsx.writeCell(sheet, rowIndex = nextRowIndex, colIndex=1:8, 
                 rowHeight = 3,
                 texts = c(thang, ngay, "", title, 
                           "9110000", "", total_cost.formatted, ""),
                 textStyle = list(TABLE_SUMMARY_STYLE + ALIGN_CENTER, TABLE_SUMMARY_STYLE + ALIGN_CENTER, TABLE_SUMMARY_STYLE,
                                  TABLE_SUMMARY_STYLE, TABLE_SUMMARY_STYLE + ALIGN_CENTER,
                                  TABLE_SUMMARY_STYLE, TABLE_CONTENT_STYLE_BOLD + ALIGN_RIGHT, TABLE_SUMMARY_STYLE),
                 multiStyles = TRUE)
  
  setPrintArea(wb, 1, 1, 8, 1, nextRowIndex)
  saveWorkbook(wb, file)
}

write.tableHeader <- function (sheet, textStyle, rowIndex) {
  rows <-createRow(sheet, rowIndex = rowIndex)
  setRowHeight(rows, multiplier = 2)
  
  cells <-createCell(rows, 1:8)
  
  addMergedRegion(sheet, rowIndex[1], rowIndex[2], 1, 1)
  setCellValue(cells[[1, 1]], "Tháng")
  
  addMergedRegion(sheet, rowIndex[1], rowIndex[1], 2, 3)
  setCellValue(cells[[1, 2]], "Chứng từ gốc")
  
  setCellValue(cells[[2, 2]], "Ngày")
  
  setCellValue(cells[[2, 3]], "Số")
  
  addMergedRegion(sheet, rowIndex[1], rowIndex[2], 4, 4)
  setCellValue(cells[[1, 4]], "Diễn giải")
  
  addMergedRegion(sheet, rowIndex[1], rowIndex[2], 5, 5)
  setCellValue(cells[[1, 5]], "TK")
  
  addMergedRegion(sheet, rowIndex[1], rowIndex[1], 6, 7)
  setCellValue(cells[[1, 6]], "Số tiền")
  
  setCellValue(cells[[2, 6]], "Nợ")
  
  setCellValue(cells[[2, 7]], "Có")
  
  addMergedRegion(sheet, rowIndex[1], rowIndex[2], 8, 8)
  setCellValue(cells[[1, 8]], "Note")
  
  for(i in 1:2) {
    for(j in 1:8) {
      setCellStyle(cells[[i, j]], textStyle)
    }
  }
  
  return(rowIndex[2] + 1)
}