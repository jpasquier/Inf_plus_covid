#  _        __                      _     _
# (_)_ __  / _|  _    ___ _____   _(_) __| |
# | | '_ \| |_ _| |_ / __/ _ \ \ / / |/ _` |
# | | | | |  _|_   _| (_| (_) \ V /| | (_| |
# |_|_| |_|_|   |_|  \___\___/ \_/ |_|\__,_|
#
#                                                   _
#  _ __  _ __ ___ _ __  _ __ ___   ___ ___  ___ ___(_)_ __   __ _
# | '_ \| '__/ _ \ '_ \| '__/ _ \ / __/ _ \/ __/ __| | '_ \ / _` |
# | |_) | | |  __/ |_) | | | (_) | (_|  __/\__ \__ \ | | | | (_| |
# | .__/|_|  \___| .__/|_|  \___/ \___\___||___/___/_|_| |_|\__, |
# |_|            |_|                                        |___/

library(readxl)
library(labelled)

# Set working directory
setwd("~/Projects/LaSource/Ortoleva - Inf+covid")

# Import data and codebook
file_name <- "data-raw/EtudeINFCOVID-19_Data_France_NumeroModalites.xlsx"
sheet_name <- "EtudeINFCOVID-19_FR"
dta <- read_xlsx(file_name, sheet = sheet_name)
dta <- as.data.frame(dta)
dta <- dta[!(grepl("^(DATE_(SAISIE|ENREG|MODIF)|E_mail(2)?)$", names(dta)))]
load("data-raw/codebook_france.rda")
cb1 <- codebook_france$variables
cb2 <- codebook_france$values

# Recoding
dta2 <- as.data.frame(lapply(setNames(names(dta), names(dta)), function(v) {
  x <- dta[[v]]
  b1 <- grepl(paste0("^", v, "$"), cb1$variable)
  b2 <- grepl(paste0("^", v, "$"), cb2$variable)
  if (any(b1)) {
    var_label(x) <- cb1$label[b1]
    if (cb1$type[b1] == "integer") {
      x <- as.integer(x)
      if (any(b2)) {
        x <- labelled(x, setNames(cb2$value[b2], cb2$label[b2]))
      }
    }
  }
  return(x)
}))
