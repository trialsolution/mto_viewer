library(readxl)
library(tidyverse)
library(xlsx)

# get the 'cube' with the full list - mergefile
load("C:/Users/himicmi/Downloads/meltdown2026/data/cube_EUNmerge_20261104_10h47.xlsx_2026-11-04-11h21.Rdata")

# get the code lists prepared in advance in Excel

regions <- c("E14",
             "EUN",
             "NMS",
             "WLD")

products <- read_excel("./drafting the report/emissions_codelist.xlsx", sheet = "products", range = cell_cols("A:A"), col_names = FALSE)
colnames(products) <- "products"

variables <- c("ECOE..TOT",
               "CO2QP")


# filter big cube
emission_results <- cube %>% filter(region %in% regions, product %in% products$products, attribute %in% variables)


# save to Excel
time_of_run <- "final"

write.xlsx(as.data.frame(emission_results), file = paste("./drafting the report/emission_results_", time_of_run, ".xlsx", sep = ""),
           row.names = FALSE, col.names = TRUE, sheetName = "emissions",
           showNA = TRUE)
