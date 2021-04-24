
# ----------------- Function which creates Rmd child files ------------------ #

rmd_txt <- function(z) {
  paste0("
```{r}
m <- M[['", z, "']]
```\n
## Modérateur : ", z, "\n
### Régressions {.tabset}\n
#### Non ajustée\n
Nombre d observations considérées : `r nrow(m$m0$model)`\n
```{r}
summary(m$m0)
f <- '/tmp/coef_", z, "_unadjusted.xlsx'
write_xlsx(tidy(m$m0), f)
embed_file(f, text = 'coefficients (xlsx)')
f <- file.remove(f)
```\n
#### Ajustée\n
Nombre d observations considérées : `r nrow(m$m1$model)`\n
```{r}
summary(m$m1)
f <- '/tmp/coef_", z, "_adjusted.xlsx'
write_xlsx(tidy(m$m1), f)
embed_file(f, text = 'coefficients (xlsx)')
f <- file.remove(f)
```\n
### Graphiques {.tabset}\n
#### Interaction (non ajustée)\n
```{r}
print(m$fig$int)
```\n
#### Droite de régression (réponse ~ modérateur)\n
```{r}
print(m$fig$reg)
```\n
#### Q-Q plot (non ajustée)\n
```{r}
print(m$fig$qq0)
```\n
#### Q-Q plot (ajustée)\n
```{r}
print(m$fig$qq1)
```
"
  )
}

# ----------------------- Generate rmarkdown reports ------------------------ #

# List of parameters
mode_exerc_list <- 
  c("Tous", "Salarié-e du publique", "Salarié-e du privé", "Libéral",
    "Mixte", "Autre")
outcome_list <- c("WHOQOL_1", "WHOQUOL_2", "WHOQOL_d1_s100", "WHOQOL_d2_s100",
                  "WHOQOL_d3_s100", "WHOQOL_d4_s100", "BienEtrPro_score")

# Initialize parameters
parms <- list(outcome = NA, mode_exerc = NA, date = "24.04.2021")

# Loop
for (outcome in outcome_list) {
  for (mode_exerc in mode_exerc_list) {
    parms$mode_exerc <- mode_exerc
    parms$outcome <- outcome
    d <- sub("([0-9]+)\\.([0-9]+)\\.([0-9]+)", "\\3\\2\\1", parms$date)
    output_dir <- paste0("../results/moderator_effects_france_", d)
    if (!dir.exists(output_dir)) dir.create(output_dir)
    me <- gsub(" ", "_", sub("é-e", "e", parms$mode_exerc))
    me <- tolower(iconv(me, from = "UTF-8", to = "ASCII//TRANSLIT"))
    output_file <- paste0(tolower(parms$outcome), "_", me, ".html")
    output_file <- path.expand(file.path(output_dir, output_file))
    rmarkdown::render(
      input = "moderator_effects_france.Rmd",
      params = parms,
      output_file = output_file
    )
  }
}
