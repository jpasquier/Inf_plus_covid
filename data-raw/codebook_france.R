#   _          __                           _      _
#  (_) _ __   / _|   _     ___  ___ __   __(_)  __| |
#  | || '_ \ | |_  _| |_  / __|/ _ \\ \ / /| | / _` |
#  | || | | ||  _||_   _|| (__| (_) |\ V / | || (_| |
#  |_||_| |_||_|    |_|   \___|\___/  \_/  |_| \__,_|
#
#                   _        _                    _
#    ___  ___    __| |  ___ | |__    ___    ___  | | __
#   / __|/ _ \  / _` | / _ \| '_ \  / _ \  / _ \ | |/ /
#  | (__| (_) || (_| ||  __/| |_) || (_) || (_) ||   <
#   \___|\___/  \__,_| \___||_.__/  \___/  \___/ |_|\_\
#

library(readxl)
library(writexl)

# Working directory
setwd("~/Projects/LaSource/Ortoleva - Inf+covid/data-raw")

# Import codebook
cb0 <- read_xlsx("EtudeINFCOVID-19_Data_France_NumeroModalites.xlsx",
                 sheet = "codebook")
cb0 <- as.data.frame(cb0)

# Codebook (variables)
cb1 <- cb0[c("Variable", "Type", "Libellé")]
names(cb1)[names(cb1) == "Libellé"] <- "label"
names(cb1) <- tolower(names(cb1))
types <- c(Date = "date", `Fermée Echelle` = "integer",
           `Fermée Multiple` = "character", `Fermée Unique` = "integer",
           `Numérique` = "numeric", Texte = "character")
cb1$type <- types[cb1$type]
rm(types)

# Codebook (values)
cb2 <- do.call(rbind, lapply(cb0$Variable, function(x) {
  i <- which(cb0$Variable == x)
  lab <- cb0[i, "Modalités"]
  if (!is.na(lab)) {
    lab <- trimws(strsplit(lab, ";")[[1]])
    val <- trimws(strsplit(cb0[i, "Codes et Informations"], ";")[[1]])
    val <- val[1:length(lab)]
    r <- data.frame(variable = x, value = val, label = lab)
  } else {
    r <- NULL
  }
  return(r)
}))
if (any(!(grepl("^[0-9]+$", cb2$value)))) {
  stop("check values")
} else {
  cb2$value <- as.integer(cb2$value)
}

# Codebook
codebook_france <- list(variables = cb1, values = cb2)

# Export
write_xlsx(codebook_france, "codebook_france.xlsx")
save(codebook_france, file = "codebook_france.rda", compress = "xz")

