source("src/write.R")
source("src/read.R")
source("src/tk632.R")
source("src/detailcost.R")

run.giavon <- function(output_dir, datetime, data.giavon, data.vtu.1540001, data.vtu.1540002) {
  je.prev <- ""
  n <- nrow(data.giavon)
  data.giavon$vnd <- as.numeric(data.giavon$vnd)
  month <- as.numeric(format(datetime, "%m"))
  print(month)
  i <- 170
  for(i in 1:n) {
    print(paste(i, "/", n))
    cr <- data.giavon$cr[i]
    vtu <- data.giavon$control_number[i]
    je <- data.giavon$no[i]
    print(paste(cr, vtu, je))
    actual.total.giavon <- sum(data.giavon[data.giavon$no == je, ]$vnd)
  
    if (je != je.prev) {
      data.vtu <- NULL
      if (cr == "1540001") {
        data.vtu <- data.vtu.1540001[data.vtu.1540001$Control_Number == vtu, ]
        co.indices <- which(is.na(data.vtu$No) | data.vtu$No == 0)
        data.co <- data.vtu[co.indices, ]
        tk <- data.co$tk[1]
        print(paste("tk =", tk))
        if (length(co.indices) > 2 || (length(co.indices) == 2 && co.indices[2] - co.indices[1] > 1)) {
          data.vtu <- data.vtu[data.vtu$Notes == month & !is.na(data.vtu$No) & data.vtu$No > 0, ]
        } else {
          data.vtu <- data.vtu[!is.na(data.vtu$No) & data.vtu$No != 0, ]
        }
        write.tk632.marine(output_dir, datetime, data.vtu, vtu, je, tk)
      } else {
        data.vtu <- data.vtu.1540002[data.vtu.1540002$Control_Number == vtu, ]
        co.indices <- which(is.na(data.vtu$No) | data.vtu$No == 0)
        data.co <- data.vtu[co.indices, ]
        n.co <- nrow(data.co)
        tk <- data.co$tk[1]
        print(paste("tk =", tk))
        if (length(co.indices) > 2 || (length(co.indices) == 2 && co.indices[2] - co.indices[1] > 1)) {
          data.vtu <- data.vtu[data.vtu$Notes == month & !is.na(data.vtu$No) & data.vtu$No > 0, ]
        } else {
          data.vtu <- data.vtu[!is.na(data.vtu$No) & data.vtu$No != 0, ]
        }
        write.tk632.energy(output_dir, datetime, data.vtu, vtu, je, tk)
      }

      data.vtu$No[is.na(data.vtu$No)] <- 0
      if (actual.total.giavon != sum(data.vtu$No)) {
        print(paste("############ Chuyen ket sai, actual =", actual.total.giavon, " | Calculated =", sum(data.vtu$No), "##############"))
      }
      data.vtu.labour <- data.vtu[data.vtu$Labour_General == "Labour", ]
      labour_cost <- sum(data.vtu.labour$No, na.rm = TRUE)
      
      data.vtu.indirectcost <- data.vtu[data.vtu$Note == "Indirect cost allocation", ]
      indirect_cost <- sum(data.vtu.indirectcost$No, na.rm = TRUE)
      
      airfare <- detailcost(data.vtu, "[Aa]irfare")
      taxi <- detailcost(data.vtu, "[Tt]axi")
      minibus <- detailcost(data.vtu, "[Mm]inibus")
      transportation <- detailcost(data.vtu, "[Tt]ransportation")
      ferry <- detailcost(data.vtu, "[Ff]erry")
      hotel <- detailcost(data.vtu, "[Hh]otel")
      subsistence <- detailcost(data.vtu, "[Ss]ubsistence")
      courier <- detailcost(data.vtu, "[Cc]ourier")
      entertainment <- detailcost(data.vtu, "[Ee]ntertainment")
      igcost <- detailcost(data.vtu, "IG [Cc]ost")
      other <- detailcost.other(data.vtu)
      
      write.giathanh(output_dir, vtu, je, time_title, signature_date, actual.total.giavon,
                     labour_cost, indirect_cost,
                     airfare, taxi, minibus, transportation, ferry, entertainment,
                     hotel, subsistence, courier, igcost, other)
      
      je.prev <- je 
    }
  }
}