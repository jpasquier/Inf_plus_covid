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
dta <- as.data.frame(lapply(setNames(names(dta), names(dta)), function(v) {
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

# --------------------------------- WHOQOL ---------------------------------- #

# Missing values
V <- grep("^WHOQOL_[0-9]+$", names(dta))
apply(is.na(dta[V]), 2, sum)
table(apply(is.na(dta[V]), 1, sum))
rm(V)

# Reverse items
dta$WHOQOL_3r <- 6 - dta$WHOQOL_3
dta$WHOQOL_4r <- 6 - dta$WHOQOL_4
dta$WHOQOL_26r <- 6 - dta$WHOQOL_26

# Transforms scores
whoqol_table4 <- lapply(list(7:35, 6:30, 3:15, 8:40), function(z) {
  w <- cbind(z, round((z - min(z)) / (max(z) - min(z)) * 16) + 4, NA)
  w[, 3] <- round((w[, 2] - 4) / 16 * 100 + 0.001)
  w
})

# Computes dimension scores
v <- list(c("3r", "4r", 10, 15:18), c(5:7, 11, 19, "26r"),
          20:22, c(8:9, 12:14, 23:25))
v <- lapply(v, function(z) paste0("WHOQOL_", z))
for(i in 1:4) {
  x0 <- paste0("WHOQOL_d", i, "_raw_mean")
  x1 <- paste0("WHOQOL_d", i, "_s20")
  x2 <- paste0("WHOQOL_d", i, "_s100")
  dta[[x0]] <- ifelse(apply(is.na(dta[v[[i]]]), 1, sum) <= c(2, 2, 1, 2)[i],
                      apply(dta[v[[i]]], 1, mean, na.rm = TRUE), NA)
  dta[[x1]] <- round(4 * (dta[[x0]]))
  dta[[x2]] <- round((dta[[x1]] - 4) / 16 * 100 + 0.001)
}
rm(v, i, x0, x1, x2)
table(apply(is.na(dta[grep("WHOQOL_d[1-4]_raw_mean", names(dta))]), 1, sum))

# ---------------------------------- PSS14 ---------------------------------- #

# Missing values
V <- grep("^PSS_[0-9]+$", names(dta))
apply(is.na(dta[V]), 2, sum)
table(apply(is.na(dta[V]), 1, sum))
rm(V)

# Computes PSS14 score (total)
dta$PSS14_score <- apply(dta[paste0("PSS_", c(1:3, 8, 11:12, 14))], 1, sum) +
  apply(4 - dta[paste0("PSS_", c(4:7, 9:10, 13))], 1, sum)

# ------------------------------- BienEtrPro -------------------------------- #

# Missing values
V <- grep("^BienEtrPro_[1-8]$", names(dta))
apply(is.na(dta[V]), 2, sum)
table(apply(is.na(dta[V]), 1, sum))
rm(V)

# Computes BienEtrPro score (mean)
dta$BienEtrPro_score <-
  apply(dta[grep("^BienEtrPro_[1-8]$", names(dta))], 1, mean)

# ------------------------------- Brief COPE -------------------------------- #

# Missing values
V <- grep("^Brief_COPE_[0-9]+$", names(dta))
apply(is.na(dta[V]), 2, sum)
table(apply(is.na(dta[V]), 1, sum))
rm(V)

# Dimensions (french version, Muller 2003)
dim_BC <- list(
  BC_Coping_actif   = c( 2, 20),
  BC_Planification  = c(13, 24),
  BC_Soutien_instru = c(10, 19),
  BC_Soutien_emotio = c( 5, 14),
  BC_Expr_sentiment = c( 9, 18),
  BC_Reinterpr_posi = c(11, 26),
  BC_Acceptation    = c( 8, 23),
  BC_Deni           = c( 3, 21),
  BC_Blame          = c(12, 25),
  BC_Humour         = c(16, 28),
  BC_Religion       = c( 7, 27),
  BC_Distraction    = c( 1, 17),
  BC_Utili_substanc = c( 4, 22),
  BC_Deseng_comport = c( 6, 15)
)

# Computes Brief COPE scores (totals)
for (x in names(dim_BC)) {
  v <- paste0("Brief_COPE_", dim_BC[[x]])
  dta[[x]] <- apply(dta[v], 1, sum)
}
rm(x, v)

# --------------------------------- PTGI-SF --------------------------------- #

# Missing values
V <- grep("^PTGI_SP[1-8]$", names(dta))
apply(is.na(dta[V]), 2, sum)
table(apply(is.na(dta[V]), 1, sum))
rm(V)

# Computes BienEtrPro score (total)
dta$PTGI_SF_score <- apply(dta[grep("^PTGI_SP[1-8]$", names(dta))], 1, sum)

# --------------------------------- CD-RISC --------------------------------- #

# Missing values
V <- grep("^CD_RISC_[0-9]+$", names(dta))
apply(is.na(dta[V]), 2, sum)
table(apply(is.na(dta[V]), 1, sum))
rm(V)

# Computes BienEtrPro score (total)
dta$CD_RISC_score <- apply(dta[grep("^CD_RISC_[0-9]+$", names(dta))], 1, sum)

# ---------------------------------- MSPSS ---------------------------------- #

# Missing values
V <- grep("^MSPSS_[0-9]+$", names(dta))
apply(is.na(dta[V]), 2, sum)
table(apply(is.na(dta[V]), 1, sum))
rm(V)

# Computes MSPSS scores (means)
dta$MSPSS_signif_other <- apply(dta[paste0("MSPSS_", c(1:2, 5, 10))], 1, mean)
dta$MSPSS_family <- apply(dta[paste0("MSPSS_", c(3:4, 8, 11))], 1, mean)
dta$MSPSS_friends <- apply(dta[paste0("MSPSS_", c(6:7, 9, 12))], 1, mean)
dta$MSPSS_score <- apply(dta[paste0("MSPSS_", 1:12)], 1, mean)

# --------------------------------- COPSOQ ---------------------------------- #

# Missing values
V <- grep("^COPSOQ_[1-8]$", names(dta))
apply(is.na(dta[V]), 2, sum)
table(apply(is.na(dta[V]), 1, sum))
rm(V)

# Preview
sapply(dta[grep("^COPSOQ_[1-8]$", names(dta))], table)

# -------------------------------- Save data -------------------------------- #

saveRDS(dta, file = "data/data_france.rds", compress = "xz")
