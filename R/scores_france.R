#  _          __                           _      _
# (_) _ __   / _|   _     ___  ___ __   __(_)  __| |
# | || '_ \ | |_  _| |_  / __|/ _ \\ \ / /| | / _` |
# | || | | ||  _||_   _|| (__| (_) |\ V / | || (_| |
# |_||_| |_||_|    |_|   \___|\___/  \_/  |_| \__,_|
#                                       __   __
#  ___   ___  ___   _ __  ___  ___     / /  / _| _ __  __ _  _ __    ___  ___
# / __| / __|/ _ \ | '__|/ _ \/ __|   / /  | |_ | '__|/ _` || '_ \  / __|/ _ \
# \__ \| (__| (_) || |  |  __/\__ \  / /   |  _|| |  | (_| || | | || (__|  __/
# |___/ \___|\___/ |_|   \___||___/ /_/    |_|  |_|   \__,_||_| |_| \___|\___|

library(writexl)

# Set working directory
setwd("~/Projects/LaSource/Ortoleva - Inf+covid")

# Import preprocessed data
dta <- readRDS("data/data_france.rds")

# Scores
X <- c("WHOQOL_1", "WHOQUOL_2", "WHOQOL_d1_s100", "WHOQOL_d2_s100",
       "WHOQOL_d3_s100", "WHOQOL_d4_s100", "PSS14_score", "BienEtrPro_score",
       "BC_Coping_actif", "BC_Planification", "BC_Soutien_instru",
       "BC_Soutien_emotio", "BC_Expr_sentiment", "BC_Reinterpr_posi",
       "BC_Acceptation", "BC_Deni", "BC_Blame", "BC_Humour", "BC_Religion",
       "BC_Distraction", "BC_Utili_substanc", "BC_Deseng_comport",
       "PTGI_SF_score", "CD_RISC_score", "MSPSS_signif_other", "MSPSS_family",
       "MSPSS_friends", "MSPSS_score")
scores <- do.call(rbind, lapply(X, function(x) {
  u <- dta[[x]]
  data.frame(
    score = x,
    n = sum(!is.na(u)),
    nmiss = sum(is.na(u)),
    mean = mean(u, na.rm = TRUE),
    sd = sd(u, na.rm = TRUE),
    median = median(u, na.rm = TRUE),
    q25 = quantile(u, .25, na.rm = TRUE),
    q75 = quantile(u, .75, na.rm = TRUE),
    min = min(u, na.rm = TRUE),
    max = max(u, na.rm = TRUE)
  )
}))
rownames(scores) <- NULL

# Export
write_xlsx(scores, "results/scores_france_20210322.xlsx")
