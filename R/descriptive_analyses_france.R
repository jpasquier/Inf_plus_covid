#  _          __                           _      _
# (_) _ __   / _|   _     ___  ___ __   __(_)  __| |
# | || '_ \ | |_  _| |_  / __|/ _ \\ \ / /| | / _` |
# | || | | ||  _||_   _|| (__| (_) |\ V / | || (_| |
# |_||_| |_||_|    |_|   \___|\___/  \_/  |_| \__,_|

library(labelled)
library(binom)
library(ggplot2)
library(XLConnect)

# Set working directory
setwd("~/Projects/LaSource/Ortoleva - Inf+covid")

# Mode d'exercice
mode_exerc_list <-
  c("Tous", "Salarié-e du publique", "Salarié-e du privé", "Libéral",
    "Mixte", "Autre")
me <- gsub(" ", "_", sub("é-e", "e", mode_exerc_list))
me <- tolower(iconv(me, from = "UTF-8", to = "ASCII//TRANSLIT"))
names(mode_exerc_list) <- me
rm(me)

# Import preprocessed data
dta <- readRDS("data/data_france.rds")

# Descriptive statistics
fig_descr <- list()
tbl_descr <- lapply(names(mode_exerc_list), function(me) {
  mode_exerc <- mode_exerc_list[me]
  fig_descr[[me]] <<- list()
  names(dta)[sapply(dta, function(x) any(class(x) == "character")) ]
  X <- names(dta)
  X <- X[!(X %in% c("N.Obs", "Sit_fam_Autre", "Specialites",
                    "Specialites_Autre", "Mode_exerc_Autre",
                    "Type_autre_dipl_Autre", "CLE"))]
  if (mode_exerc != "Tous") {
    X <- X[X != "Mode_exerc"]
    b <- !is.na(dta$Mode_exerc) & to_factor(dta$Mode_exerc) == mode_exerc
    dta <- dta[b, ]
    rm(b)
  }
  X <- X[apply(!is.na(dta[X]), 2, sum) > 0]
  do.call(rbind, lapply(X, function(x) {
    if (any(class(dta[[x]]) == "haven_labelled")) dta[[x]] <- to_factor(dta[[x]])
    type <- paste(class(dta[[x]]), collapse = " ")
    z <- na.omit(dta[[x]])
    if (any(class(dta[[x]]) %in% c("factor", "logical"))) {
      if (any(class(z) == "factor")) z <- droplevels(z)
      tab <- table(z)
      ci <- t(sapply(tab, function(x) {
        y <- binom.confint(x = x, n = sum(tab), method = "wilson")
        unlist(y[, c("lower", "upper")])
      }))
      tab <- cbind(tab, prop.table(tab), ci)
      tab <- data.frame(rownames(tab), tab)
      if (any(nchar(as.character(z)) > 15)) {
        rotate <- TRUE
        levels(dta[[x]]) <- sapply(levels(dta[[x]]), function(w) {
          if (nchar(w) > 15) w <- paste0(substr(w, 1, 12), "...")
          w
        })
      } else {
        rotate <- FALSE
      }
      cap <- paste("Mode d'exercice :", mode_exerc, "/ N :",
                   sum(!is.na(dta[[x]])))
      fig <- ggplot(dta[!is.na(dta[[x]]), ], aes_string(x = x)) +
        geom_bar(aes(y = (..count..)/sum(..count..)), fill = "steelblue") +
        scale_y_continuous(labels=scales::percent) +
        labs(y = "Fréquences relatives", caption = cap) +
        theme_classic()
      if (rotate) {
        fig <- fig +
          theme(axis.text.x = element_text(angle = 45, hjust = 1))
      }
      fig_descr[[me]][[x]] <<- fig
    } else if (class(dta[[x]]) %in% c("numeric", "integer")) {
      if (length(z) > 1) {
        hci <- qt(0.975, length(z) - 1) * sd(z) / sqrt(length(z))
      } else {
        hci <- NA
      }
      tab <- data.frame(
        c("Mean/SD", "Median/IQR", "Min/Max"),
        c(mean(z), median(z), min(z)),
        c(sd(z), IQR(z), max(z)),
        c(mean(z) - hci, NA, NA),
        c(mean(z) + hci, NA, NA)
      )
      brx <- pretty(range(z), n = nclass.Sturges(z), min.n = 1)
      fig <- ggplot(dta[!is.na(dta[[x]]), ], aes_string(x = x)) +
        geom_histogram(color = "darkgray", fill = "steelblue", breaks = brx)
      y_range <- ggplot_build(fig)$layout$panel_params[[1]]$y.range
      fig <- fig +
        geom_boxplot(aes(y = y_range[2] * 1.1),
                     width = sum(y_range * c(-1, 1)) / 10) +
        theme_classic()
      fig_descr[[me]][[x]] <<- fig
    } else {
      stop(paste0(x, ": unsupported type (", type, ")"))
    }
    tab <- cbind(x, type, sum(is.na(dta[[x]])), tab)
    names(tab) <- 
      c("Variable", "Type", "Nmiss", "Value", "(N)", "(Prop)", "2.5 %", "97.5 %")
    rownames(tab) <- NULL
    return(tab)
  }))
})
names(tbl_descr) <- names(mode_exerc_list)

# Export directory
output_dir <- "results/descriptive_analyses_20210424"
if (!dir.exists(output_dir)) dir.create(output_dir)

# Export results in an Excel file
f <- file.path(output_dir, "tables.xlsx")
if (file.exists(f)) file.remove(f)
wb <- loadWorkbook(f, create = TRUE)
for (tbl in names(tbl_descr)) {
  ## Init WB
  r <- tbl_descr[[tbl]]
  s <- tbl
  createSheet(wb, name = s)
  writeWorksheet(wb, r, sheet = s)
  ## Merged cells
  mc_var <- list(r$Variable)
  mc_list <- c("B", "C")
  mc <- aggregate(1:nrow(r), mc_var, function(x) {
    if(length(x) > 1) {
      x <- x + 1
      out <- paste0("A", min(x), ":A", max(x))
    } else {
      out <- NA
    }
    return(out)
  })[, length(mc_var) + 1]
  mc <- mc[!is.na(mc)]
  mc <- c(mc, sapply(mc_list, function(z) gsub("A", z, mc)))
  mergeCells(wb, sheet = s, reference = mc)
  ## Borders
  cs <- createCellStyle(wb)
  setBorder(cs, side = "top", type = XLC$"BORDER.THIN",
            color = XLC$"COLOR.BLACK")
  for(i in aggregate(1:nrow(r), mc_var, min)[, length(mc_var) + 1] + 1) {
    setCellStyle(wb, sheet = s, row = i, col = 1:ncol(r), cellstyle = cs)
  }
  ## Colors
  cs <- createCellStyle(wb)
  setFillPattern(cs, fill = XLC$"FILL.SOLID_FOREGROUND")
  setFillForegroundColor(cs, color = XLC$"COLOR.YELLOW")
}
## Save WB
saveWorkbook(wb)
## Remove unused variables
rm(f, wb, tbl, r, s, mc_var, mc_list, mc, cs, i)

# Export figures
for (me in names(fig_descr)) {
  f <- file.path(output_dir, paste0("figs_", me, ".pdf"))
  cairo_pdf(f, onefile = TRUE)
  for (fig in fig_descr[[me]]) print(fig)
  garbage <- dev.off()
}
rm(me, f, fig, garbage)

# Session infos
sink(file.path(output_dir, "sessionInfo.txt"))
print(sessionInfo(), locale = FALSE)
sink()
