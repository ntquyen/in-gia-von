write.address <- function (sheet, rowIndex = 1, textStyle) {
  nextRowIndex <- rowIndex
  nextRowIndex <- xlsx.writeCell(sheet, rowIndex = nextRowIndex, textStyle = textStyle,
                                 texts = c("CÔNG TY TNHH LLOYD'S REGISTER ASIA (VIETNAM)"))
  nextRowIndex <- xlsx.writeCell(sheet, rowIndex = nextRowIndex, textStyle = textStyle,
                                 texts = c("P450 PetroVietnam Tower, 8 Hoàng Diệu, P1, TP. Vũng Tàu"))
  nextRowIndex <- xlsx.writeCell(sheet, rowIndex = nextRowIndex, textStyle = textStyle,
                                 texts = c("MST: 3502101783"))
  return(nextRowIndex)
}

xlsx.writeCell <- function(sheet, 
                           rowIndex = 1, 
                           colIndex = 1,
                           ignoredCell,
                           rowHeight, 
                           mergedRegion,
                           texts, textStyle, multiStyles = FALSE) {
  
  rows <-createRow(sheet, rowIndex = rowIndex)
  cells <-createCell(rows, colIndex = colIndex)
  
  if (!missing(rowHeight)) {
    setRowHeight(rows, multiplier = rowHeight)
  }

  if (!missing(mergedRegion)) {
    addMergedRegion(sheet, mergedRegion[1], mergedRegion[2], mergedRegion[3], mergedRegion[4])
  }
  
  for(i in 1:length(cells)) {
    if (missing(ignoredCell) || ignoredCell != i) {
      setCellValue(cells[[1, i]], texts[i])
      
      if (!multiStyles) {
        setCellStyle(cells[[1, i]], textStyle)
      } else {
        setCellStyle(cells[[1, i]], textStyle[[i]])
      }
    }
    
  }
  return(rowIndex + 1)
}