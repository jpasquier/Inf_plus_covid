library(parallel)
library(labelled)
library(ggplot2)
library(rlang)
library(writexl)

options(mc.cores=detectCores())

# Set working directory
setwd("~/Projects/LaSource/Ortoleva - Inf+covid")

# Import preprocessed data
dta <- readRDS("data/data_france.rds")

# Recoding of the response variable
for (y in c("WHOQOL_1", "WHOQUOL_2")) dta[[y]] <- as.numeric(dta[[y]])
rm(y)

# Outcomes
Y <- c("WHOQOL_1", "WHOQUOL_2", "WHOQOL_d1_s100", "WHOQOL_d2_s100",
       "WHOQOL_d3_s100", "WHOQOL_d4_s100", "BienEtrPro_score")

# Exposure
X <- "PSS14_score"

# Moderators
Z <- c("BC_Coping_actif", "BC_Planification", "BC_Soutien_instru",
       "BC_Soutien_emotio", "BC_Expr_sentiment", "BC_Reinterpr_posi",
       "BC_Acceptation", "BC_Deni", "BC_Blame", "BC_Humour", "BC_Religion",
       "BC_Distraction", "BC_Utili_substanc", "BC_Deseng_comport",
       "PTGI_SF_score", "CD_RISC_score", "MSPSS_signif_other", "MSPSS_family",
       "MSPSS_friends", "MSPSS_score", "COPSOQ_soutien_superieur",
       "COPSOQ_soutien_collegues", "COPSOQ_satisf_qualite")

# Strata
M <- c("Tous", "Salarié-e du publique", "Salarié-e du privé", "Libéral",
       "Mixte", "Autre")
names(M) <-  tolower(iconv(gsub(" ", "_", sub("é-e", "e", M)), from = "UTF-8",
                           to = "ASCII//TRANSLIT"))

# Correlation matrices
cormat <- mclapply(M, function(m) {
  if (m == "Tous") {
    ttl <- "Toutes les observations"
  } else {
    ttl <- m
    dta <- dta[!is.na(dta$Mode_exerc) & to_factor(dta$Mode_exerc) == m, ]
  }
  V <- c(X, Y, Z)
  mat <- do.call(rbind, mclapply(1:(length(V) - 1), function(i) {
    do.call(rbind, mclapply((i + 1):length(V), function(j) {
      var1 <- V[i]
      var2 <- V[j]
      x <- dta[[var1]]
      y <- dta[[var2]]
      b <- !is.na(x) & !is.na(y)
      x <- x[b]
      y <- y[b]
      n <- sum(b)
      if (n <= 3) return(NULL)
      delta <- qnorm(0.975) / sqrt(n - 3)
      cor_pearson <- cor(x, y)
      cor_pearson_lwr <- tanh(atanh(cor_pearson) - delta)
      cor_pearson_upr <- tanh(atanh(cor_pearson) + delta)
      cor_spearman <- cor(x, y, method="spearman")
      cor_spearman_lwr <- tanh(atanh(cor_spearman) - delta)
      cor_spearman_upr <- tanh(atanh(cor_spearman) + delta)
      data.frame(
        var1 = var1,
        var2 = var2,
        n = n,
        cor_pearson = cor_pearson,
        cor_pearson_lwr = cor_pearson_lwr,
        cor_pearson_upr = cor_pearson_upr,
        cor_spearman = cor_spearman,
        cor_spearman_lwr = cor_spearman_lwr,
        cor_spearman_upr = cor_spearman_upr
      )
    }))
  }))
  figs <- mclapply(1:2, function(k) {
    sttl <- paste("Coefficient de corrélation de", c("Pearson", "Spearman")[k])
    z <- c("cor_pearson", "cor_spearman")[k]
    fig <- ggplot(mat, aes(factor(var1, V), factor(var2, rev(V)),
                               fill=!!sym(z))) +
      geom_tile(color="white") +
      scale_fill_gradient2(low="blue", high="red", mid="white", midpoint=0,
                           limit=c(-1, 1), space="Lab", name="") +
      theme_minimal() +
      theme(axis.text.x=element_text(angle=90, vjust=0, hjust=1)) +
      coord_fixed() +
      labs(x="", y="", title = ttl, subtitle = sttl)
    return(fig)
  })
  attr(mat, "figs") <- figs
  return(mat)
})

# Export results
write_xlsx(cormat, "results/correlation_matrices_20210427.xlsx")
pdf("results/correlation_matrices_20210427.pdf")
for (mat in cormat) {
  print(attr(mat, "figs")[[1]])
  print(attr(mat, "figs")[[2]])
}
dev.off()
rm(mat)
