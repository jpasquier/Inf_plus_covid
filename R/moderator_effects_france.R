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

d <- format(Sys.time(), "%Y%m%d")
f <- paste0("../results/moderator_effects_france_", d, ".html")
rmarkdown::render("moderator_effects_france.Rmd", output_file = f)
