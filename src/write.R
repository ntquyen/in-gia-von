source("src/writeLib.R")

setPrintSetupTk154 <- function (sheet) {
  printSetup(sheet, copies = 1, paperSize = "A4_PAPERSIZE")
  sheet$setFitToPage(TRUE)
  sheet$setHorizontallyCenter(TRUE)
  
  sheet$setMargin(.jshort(sheet$LeftMargin), 0.7)
  sheet$setMargin(.jshort(sheet$RightMargin), 0.3)
}

write.giathanh <- function(output_dir, vtu, je, time, signature_date, actual.total.giavon,
                           labour_cost, indirect_cost,
                           airfare, taxi, minibus, transportation, ferry, entertainment,
                           hotel, subsistence, courier, igcost, other) {
  file <- paste("/Users/quyennt/Projects/in-gia-von/output/", output_dir, "/",
                je, "_giathanh.xlsx", sep="")
  print(file)
  tableRowHeight <- 1.2
  emptyRowHeight <- 1
  
  direct_cost <- airfare + taxi + minibus + transportation + ferry + entertainment +
    hotel + subsistence + courier + igcost + other
  general_cost = direct_cost + indirect_cost
  total_cost <- labour_cost + direct_cost + indirect_cost
  
  wb<-createWorkbook(type = "xlsx")
  sheet <- createSheet(wb, sheetName = vtu)
  setPrintSetupTk154(sheet)
  
  BORDER_STYLE_THIN <- Border(color = "Black", position = c("BOTTOM", "LEFT", "TOP", "RIGHT"))
  BORDER_STYLE_DOTTED <- Border(color = "Black", position = c("BOTTOM", "LEFT", "TOP", "RIGHT"), 
                                pen = c("BORDER_DOTTED", "BORDER_THIN", "BORDER_DOTTED", "BORDER_THIN"))
  
  ALIGN_RIGHT <- Alignment(h="ALIGN_RIGHT", v="VERTICAL_CENTER", wrapText = TRUE)
  ALIGN_CENTER <- Alignment(h="ALIGN_CENTER", v="VERTICAL_CENTER", wrapText = TRUE) 
  ALIGN_LEFT <- Alignment(h="ALIGN_LEFT", v="VERTICAL_CENTER", wrapText = TRUE) 
  
  ADDRESS_STYLE <- CellStyle(wb) + 
    Font(wb, heightInPoints=9, name = "Arial")
  TITLE_STYLE <- CellStyle(wb) + 
    Font(wb, heightInPoints=16, name = "Times New Roman", isBold = TRUE) + ALIGN_CENTER
  JE_STYLE <- CellStyle(wb) + 
    Font(wb, heightInPoints=12, name = "Times New Roman", isBold = TRUE) + ALIGN_RIGHT
  TIME_STYLE <- CellStyle(wb) + 
    Font(wb, heightInPoints=13, name = "Times New Roman", isBold = TRUE) + ALIGN_CENTER
  
  TABLE_HEADER_STYLE <- CellStyle(wb) + 
    Font(wb, heightInPoints=12, name = "Times New Roman", isBold = TRUE) +
    ALIGN_CENTER + BORDER_STYLE_THIN
  
  TABLE_CONTENT_STYLE_BOLD <- CellStyle(wb) + 
    Font(wb, heightInPoints=12, name = "Times New Roman", isBold = TRUE) + BORDER_STYLE_THIN
  
  TABLE_CONTENT_STYLE <- CellStyle(wb) + 
    Font(wb, heightInPoints=12, name = "Times New Roman") + BORDER_STYLE_DOTTED
  
  TABLE_SUM_STYLE <- CellStyle(wb) + 
    Font(wb, heightInPoints=14, name = "Times New Roman", isBold = TRUE) + BORDER_STYLE_THIN
  
  DATE_STYLE <- CellStyle(wb) + 
    Font(wb, heightInPoints=12, name = "Times New Roman") +
    Alignment(h="ALIGN_CENTER")
  SIGNATURE_STYLE <- CellStyle(wb) + 
    Font(wb, heightInPoints=12, name = "Times New Roman")
  SIGNATURE_CENTER_STYLE <- CellStyle(wb) + 
    Font(wb, heightInPoints=12, name = "Times New Roman") + 
    Alignment(h="ALIGN_CENTER")
  
  setColumnWidth(sheet, 1, 9)
  setColumnWidth(sheet, 2, 48.7)
  setColumnWidth(sheet, c(3, 5, 6), 18)
  
  nextRowIndex <- write.address(sheet, 1, ADDRESS_STYLE)
  
  title <- paste("PHIẾU TÍNH GIÁ VỐN DỊCH VỤ - ", vtu, "\n", "COST OF SERVICE SOLD SLIP", sep="")
  nextRowIndex <- xlsx.writeCell(sheet, rowIndex = nextRowIndex + 1, colIndex=1, textStyle = TITLE_STYLE,
                 rowHeight = 3, mergedRegion=c(nextRowIndex + 1, nextRowIndex + 1, 1, 3),
                 texts = c(title))
  
  nextRowIndex <- xlsx.writeCell(sheet, rowIndex = nextRowIndex, colIndex = 3, textStyle = JE_STYLE, 
                 texts = c(je))
  nextRowIndex <- xlsx.writeCell(sheet, rowIndex = nextRowIndex, colIndex = 2, textStyle = TIME_STYLE, 
                 texts = c(time))
  
  nextRowIndex <- xlsx.writeCell(sheet, rowIndex = nextRowIndex + 1, colIndex = 1:3, rowHeight = tableRowHeight, textStyle = TABLE_HEADER_STYLE,
               texts = c("STT/No", "Diễn giải / Description", "Số tiền / Amount"))
  nextRowIndex <- xlsx.writeCell(sheet, rowIndex = nextRowIndex, colIndex = 1:3, rowHeight = emptyRowHeight, textStyle = TABLE_HEADER_STYLE,
               texts = c("", "", ""))
  
  labour_cost.formatted <- format(labour_cost, big.mark=",", scientific = FALSE)
  nextRowIndex <- xlsx.writeCell(sheet, rowIndex = nextRowIndex, colIndex = 1:3, rowHeight = tableRowHeight, multiStyles = TRUE,
                 textStyle = list(TABLE_CONTENT_STYLE_BOLD + ALIGN_CENTER, TABLE_CONTENT_STYLE_BOLD + ALIGN_LEFT, TABLE_CONTENT_STYLE_BOLD + ALIGN_RIGHT),
               texts = c("1", "Chi phí nhân công trực tiếp - Direct Labour Cost", 
                         labour_cost.formatted))
  
  nextRowIndex <- xlsx.writeCell(sheet, rowIndex = nextRowIndex, colIndex = 1:3, rowHeight = emptyRowHeight, textStyle = TABLE_CONTENT_STYLE, texts = c("", "", ""))
  
  general_cost.formatted <- format(general_cost, big.mark=",", scientific = FALSE)
  nextRowIndex <- xlsx.writeCell(sheet, rowIndex = nextRowIndex, colIndex = 1:3, rowHeight = tableRowHeight, multiStyles = TRUE,
                 textStyle = list(TABLE_CONTENT_STYLE_BOLD + ALIGN_CENTER, TABLE_CONTENT_STYLE_BOLD + ALIGN_LEFT, TABLE_CONTENT_STYLE_BOLD + ALIGN_RIGHT),
               texts = c("2", "Chi phí sản xuất chung - General Cost", 
                         general_cost.formatted))

  direct_cost.formatted <- format(direct_cost, big.mark=",", scientific = FALSE)
  nextRowIndex <- xlsx.writeCell(sheet, rowIndex = nextRowIndex, colIndex = 1:3, rowHeight = tableRowHeight, multiStyles = TRUE,
                 textStyle = list(TABLE_CONTENT_STYLE_BOLD + ALIGN_CENTER, TABLE_CONTENT_STYLE_BOLD + ALIGN_LEFT, TABLE_CONTENT_STYLE_BOLD + ALIGN_RIGHT),
               texts = c("2.1", "Chi phí trực tiếp - Direct expenses", direct_cost.formatted))

  TABLE_DETAIL_STYLE <- list(TABLE_CONTENT_STYLE, TABLE_CONTENT_STYLE + ALIGN_LEFT, TABLE_CONTENT_STYLE + ALIGN_RIGHT)
  if (airfare > 0) {
    airfare.formatted <- format(airfare, big.mark=",", scientific = FALSE)
    nextRowIndex <- xlsx.writeCell(sheet, rowIndex = nextRowIndex, colIndex = 1:3, rowHeight = tableRowHeight, multiStyles = TRUE,
                   textStyle = TABLE_DETAIL_STYLE, texts = c("", "Vé máy bay - Airfare", airfare.formatted))
  }
  if (taxi > 0) {
    taxi.formatted <- format(taxi, big.mark=",", scientific = FALSE)  
  } else {
    taxi.formatted <- ""
  }
  
  nextRowIndex <- xlsx.writeCell(sheet, rowIndex = nextRowIndex, colIndex = 1:3, rowHeight = tableRowHeight, multiStyles = TRUE,
                 textStyle = TABLE_DETAIL_STYLE, texts = c("", "Phí Taxi - Taxi Fee", taxi.formatted))
  
  if (minibus > 0) {
    minibus.formatted <- format(minibus, big.mark=",", scientific = FALSE)
    nextRowIndex <- xlsx.writeCell(sheet, rowIndex = nextRowIndex, colIndex = 1:3, rowHeight = tableRowHeight, multiStyles = TRUE,
                   textStyle = TABLE_DETAIL_STYLE, texts = c("", "Vé xe khách - Minibus", minibus.formatted))
  }
  
  if (transportation) {
    transportation.formatted <- format(transportation, big.mark=",", scientific = FALSE)
    nextRowIndex <- xlsx.writeCell(sheet, rowIndex = nextRowIndex, colIndex = 1:3, rowHeight = tableRowHeight,  multiStyles = TRUE,
                   textStyle = TABLE_DETAIL_STYLE, texts = c("", "CP Vận chuyển - Transportation", transportation.formatted))
  }

  if (ferry > 0) {
    ferry.formatted <- format(ferry, big.mark=",", scientific = FALSE)
    nextRowIndex <- xlsx.writeCell(sheet, rowIndex = nextRowIndex, colIndex = 1:3, rowHeight = tableRowHeight, multiStyles = TRUE,
                   textStyle = TABLE_DETAIL_STYLE, texts = c("", "Vé tàu cánh ngầm - Ferry", ferry.formatted))
  }
  
  if (hotel > 0) {
    hotel.formatted <- format(hotel, big.mark=",", scientific = FALSE)
    nextRowIndex <- xlsx.writeCell(sheet, rowIndex = nextRowIndex, colIndex = 1:3, rowHeight = tableRowHeight, multiStyles = TRUE,
                   textStyle = TABLE_DETAIL_STYLE, texts = c("", "CP Khách sạn - Hotel Charge", hotel.formatted))
  }

  if (subsistence > 0) {
    subsistence.formatted <- format(subsistence, big.mark=",", scientific = FALSE)
    nextRowIndex <- xlsx.writeCell(sheet, rowIndex = nextRowIndex, colIndex = 1:3, rowHeight = tableRowHeight, multiStyles = TRUE,
                   textStyle = TABLE_DETAIL_STYLE, texts = c("", "Phụ cấp lưu trú - Subsistence", subsistence.formatted))
  }
  
  if (entertainment > 0) {
    entertainment.formatted <- format(entertainment, big.mark=",", scientific = FALSE)
    nextRowIndex <- xlsx.writeCell(sheet, rowIndex = nextRowIndex, colIndex = 1:3, rowHeight = tableRowHeight, multiStyles = TRUE,
                   textStyle = TABLE_DETAIL_STYLE, texts = c("", "CP Tiếp khách - Client Entertainment", entertainment.formatted))
  }

  if (courier > 0) {
    courier.formatted <- format(courier, big.mark=",", scientific = FALSE)
    nextRowIndex <- xlsx.writeCell(sheet, rowIndex = nextRowIndex, colIndex = 1:3, rowHeight = tableRowHeight, multiStyles = TRUE,
                  textStyle = TABLE_DETAIL_STYLE, texts = c("", "Cước dịch vụ chuyển phát - Courier charge", courier.formatted))
  }

  if (igcost > 0) {
    igcost.formatted <- format(igcost, big.mark=",", scientific = FALSE)
    nextRowIndex <- xlsx.writeCell(sheet, rowIndex = nextRowIndex, colIndex = 1:3, rowHeight = tableRowHeight, multiStyles = TRUE,
                  textStyle = TABLE_DETAIL_STYLE, texts = c("", "Phê duyệt thiết kế - Design appraisal", igcost.formatted))
  }

  if (other > 0) {
    other.formatted <- format(other, big.mark=",", scientific = FALSE)  
  } else {
    other.formatted <- ""
  }
  nextRowIndex <- xlsx.writeCell(sheet, rowIndex = nextRowIndex, colIndex = 1:3, rowHeight = tableRowHeight, multiStyles = TRUE,
                                 textStyle = TABLE_DETAIL_STYLE,
                                 texts = c("", "Chi phí khác - Other Expenses", other.formatted))

  indirect_cost.formatted <- format(indirect_cost, big.mark=",", scientific = FALSE)

  nextRowIndex <- xlsx.writeCell(sheet, rowIndex = nextRowIndex, colIndex = 1:3, rowHeight = tableRowHeight, textStyle = TABLE_CONTENT_STYLE,
               texts = c("", "", ""))
  nextRowIndex <- xlsx.writeCell(sheet, rowIndex = nextRowIndex, colIndex = 1:3, rowHeight = tableRowHeight, multiStyles = TRUE,
                 textStyle = list(TABLE_CONTENT_STYLE_BOLD + ALIGN_CENTER, TABLE_CONTENT_STYLE_BOLD + ALIGN_LEFT, TABLE_CONTENT_STYLE_BOLD + ALIGN_RIGHT),
                 texts = c("2.2", "Chi phí gián tiếp phân bổ - Indirect cost allocation", 
                           indirect_cost.formatted))
  
  total_cost.formatted <- format(total_cost, big.mark=",", scientific = FALSE)
  actual.total.giavon.formatted <- format(actual.total.giavon, big.mark=",", scientific = FALSE)
  total.dif <- total_cost - actual.total.giavon
  total.dif <- format(total.dif, big.mark=",", scientific = FALSE)
  
  nextRowIndex <- xlsx.writeCell(sheet, rowIndex = nextRowIndex, colIndex = 1:3, rowHeight = emptyRowHeight, textStyle = TABLE_CONTENT_STYLE,
               texts = c("", "", ""))
  
  nextRowIndex <- xlsx.writeCell(sheet, rowIndex = nextRowIndex, colIndex = 1:6, multiStyles = TRUE,
                                 rowHeight = 2.5,
                 textStyle = list(TABLE_CONTENT_STYLE_BOLD, 
                                  TABLE_SUM_STYLE + ALIGN_CENTER, 
                                  TABLE_SUM_STYLE + ALIGN_RIGHT, 
                                  SIGNATURE_STYLE, SIGNATURE_STYLE, SIGNATURE_STYLE),
                 ignoredCell = 4,
               texts = c("", "TỔNG GIÁ VỐN DỊCH VỤ\nCOST OF SERVICE SOLD", 
                         total_cost.formatted, "", actual.total.giavon.formatted, total.dif))
  
  nextRowIndex <- xlsx.writeCell(sheet, rowIndex = nextRowIndex + 3, colIndex = 3, textStyle = DATE_STYLE, 
               texts = c(signature_date))
  
  xlsx.writeCell(sheet, rowIndex = nextRowIndex, colIndex = 1:3, 
                 textStyle = list(SIGNATURE_STYLE, SIGNATURE_STYLE, SIGNATURE_CENTER_STYLE),
                 multiStyles = TRUE,
                 ignoredCell = 2,
               texts = c("Người lập phiếu / Prepared by", "", "Kế toán trưởng / Chief Accountant"))
  
  nextRowIndex <- xlsx.writeCell(sheet, rowIndex = nextRowIndex + 6, colIndex = 1:3, textStyle = SIGNATURE_STYLE,
                 ignoredCell = 2,
               mergedRegion=c(nextRowIndex + 6, nextRowIndex + 6, 1, 2),
               texts = c("Nguyễn Thị Thu Thủy", "", "Lê Thị Thanh Thủy"))
  
  setPrintArea(wb, 1, 1, 4, 1, 40)
  saveWorkbook(wb, file)
}