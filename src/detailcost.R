detailcost <- function(df, pattern) {
  result <- 0
  df.detail <- df[grepl(pattern, df$Note), ]
  if (nrow(df.detail) > 0) {
    result <- sum(df.detail$No, na.rm = TRUE)
  }
  return(result)
}

detailcost.other <- function(df) {
  result <- 0
  df.other <- df[grepl("[Oo]ther", df$Note), ]
  
  df.other2 <- df[!grepl("[Oo]ther", df$Note) &
                    !grepl("[Aa]irfare", df$Note) &
                    !grepl("[Tt]axi", df$Note) &
                    !grepl("[Mm]inibus", df$Note) &
                    !grepl("[Tt]ransportation", df$Note) &
                    !grepl("[Ff]erry", df$Note) &
                    !grepl("[Hh]otel", df$Note) &
                    !grepl("[Ss]ubsistence", df$Note) &
                    !grepl("[Cc]ourier", df$Note) &
                    !grepl("IG [Cc]ost", df$Note) &
                    !grepl("Direct Labour Cost", df$Note) &
                    !grepl("[Ee]ntertainment", df$Note) &
                    !grepl("Indirect cost allocation", df$Note), ]
  
  if (nrow(df.other) > 0 || nrow(df.other2) > 0) {
    result <- sum(df.other$No, na.rm = TRUE) + sum(df.other2$No, na.rm = TRUE)
  }
  return(result)
}