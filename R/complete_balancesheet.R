# script to read the huge merge file of the Aglink runs

library(readxl)
library(tidyverse)
library(xlsx)


# EUN

fdp_eun <- read_excel("Y:/work/EUNDAIRY.xlsx", sheet = "EUNDAIRYDB", range = "M457:BD483", 
                      col_names = FALSE, col_types = c(rep("text",3), rep("numeric",41)))
fdp_eun <- as_tibble(fdp_eun)

# name first columns
colnames(fdp_eun)[1] <- "region"
colnames(fdp_eun)[2] <- "product"
colnames(fdp_eun)[3] <- "attribute"

for (i in 4:44) {
  colnames(fdp_eun)[i] <- 1996+i
}

# NMS

fdp_nms <- read_excel("Y:/work/EUNDAIRY.xlsx", sheet = "EUNDAIRYDB", range = "M487:BD516", 
                      col_names = FALSE, col_types = c(rep("text",3), rep("numeric",41)))
fdp_nms <- as_tibble(fdp_nms)

# name first columns
colnames(fdp_nms)[1] <- "region"
colnames(fdp_nms)[2] <- "product"
colnames(fdp_nms)[3] <- "attribute"

for (i in 4:44) {
  colnames(fdp_nms)[i] <- 1996+i
}


# E14

fdp_e14 <- read_excel("Y:/work/EUNDAIRY.xlsx", sheet = "EUNDAIRYDB", range = "M523:BD555", 
                      col_names = FALSE, col_types = c(rep("text",3), rep("numeric",41)))
fdp_e14 <- as_tibble(fdp_e14)

# name first columns
colnames(fdp_e14)[1] <- "region"
colnames(fdp_e14)[2] <- "product"
colnames(fdp_e14)[3] <- "attribute"

for (i in 4:44) {
  colnames(fdp_e14)[i] <- 1996+i
}

fdp <- rbind(fdp_eun, fdp_e14, fdp_nms)

# save the data on fresh dairy commodities in R format
time_of_run <- format(Sys.time(), "%Y-%m-%d-%Hh%M")
save(fdp, file = paste("data/fdp_", time_of_run, ".Rdata", sep = ""))



# Alternatively, load FDP products data from the .RData file
# (if previously processed)
#load("data/fdp_2025-10-16-12h15.Rdata")

# time stamp
time_of_run <- format(Sys.time(), "%Y-%m-%d-%Hh%M")



# calculate QC = QP-NT
x <- fdp %>% pivot_longer(starts_with("20"),names_to = "year",values_to = "value")

# QP calculated for EUN (sum of NMS and E14)
x_qp <- x %>% filter(region != "EUN", attribute == "QP") %>% group_by(product,year) %>% summarise( value = sum(value))
x_qp <- ungroup(x_qp)
# NT calculated for EUN (sum of NMS and E14)
x_nt <- x %>% filter(region != "EUN", attribute == "NT") %>% group_by(product,year) %>% summarise( value = sum(value))
x_nt <- ungroup(x_nt)

x_qc <- x_qp %>% full_join(x_nt, by = c("product","year"))
colnames(x_qc)[3] <- "qp"
colnames(x_qc)[4] <- "nt"

x_qc <- x_qc %>% mutate(qc = qp - nt)


# write QC and QP to Excel

# QC
to_excel <- x_qc %>% select(-nt,-qp) %>% pivot_wider(names_from = year, values_from = qc) 

# save to Excel
write.xlsx(as.data.frame(to_excel), file = paste("reporting/fdps_qc_", time_of_run, ".xlsx", sep = ""),
           row.names = FALSE, col.names = TRUE, sheetName = "qc",
           showNA = TRUE)

# QP
to_excel <- x_qc %>% select(-nt,-qc) %>% pivot_wider(names_from = year, values_from = qp) 

# save to Excel
write.xlsx(as.data.frame(to_excel), file = paste("reporting/fdps_qp_", time_of_run, ".xlsx", sep = ""),
           row.names = FALSE, col.names = TRUE, sheetName = "qp",
           showNA = TRUE)


# Calculate FOA, by adding waste to FO
# this would be TRUE..FO + share * FDP_WST..DIST
# the shares are NMS/E14 specific

x <- fdp %>% pivot_longer(starts_with("20"),names_to = "year",values_to = "value")

# Get TRUE..FO
x_truefo <- x %>% filter(region != "EUN", attribute == "TRUE..FO") %>% 
  group_by(region,product,year) %>% summarise( value = sum(value))

x_truefo <- ungroup(x_truefo)

# Get WST..DIST


# Create waste for the fresh dairy products
waste_cream <- x_wstdist
waste_cream$product <- "Cream"
waste_cream <- waste_cream %>% mutate(value = 0.6*value)

calculate_waste <- function(fdp_data = fdp, fdp_product = "Cream", waste_share = 0.08, eu_region){
  
  x_wstdist <- x %>% filter(region != "EUN", attribute == "WST..DIST", region == eu_region) %>% 
    group_by(region,product,year) %>% summarise( value = sum(value))
  x_wstdist <- ungroup(x_wstdist)  
  
  waste <- x_wstdist
  waste$product <- fdp_product
  waste <- waste %>% mutate(value = waste_share*value)
  
  return(waste)  
}

waste_dmilk_nms <- calculate_waste(fdp_data = fdp, fdp_product = "Milk",  waste_share = 0.6, eu_region = "NMS")
waste_cream_nms <- calculate_waste(fdp_data = fdp, fdp_product = "Cream", waste_share = 0.08, eu_region = "NMS")
waste_yogurt_nms <- calculate_waste(fdp_data = fdp, fdp_product = "Yogurt", waste_share = 0.3, eu_region = "NMS")

waste_dmilk_e14 <- calculate_waste(fdp_data = fdp, fdp_product = "Milk",  waste_share = 0.6, eu_region = "E14")
waste_cream_e14 <- calculate_waste(fdp_data = fdp, fdp_product = "Cream", waste_share = 0.05, eu_region = "E14")
waste_yogurt_e14 <- calculate_waste(fdp_data = fdp, fdp_product = "Yogurt", waste_share = 0.2, eu_region = "E14")

to_excel <- rbind(waste_cream_e14, waste_cream_nms, waste_dmilk_e14, waste_dmilk_nms, waste_yogurt_e14, waste_yogurt_nms)
to_excel <- to_excel %>% pivot_wider(names_from = year)

write.xlsx(as.data.frame(to_excel), file = paste("reporting/waste_fdps_", time_of_run, ".xlsx", sep = ""),
           row.names = FALSE, col.names = TRUE, sheetName = "wst..dist",
           showNA = TRUE)

# Calculate FOA per capita, by adding waste to FO..POP
# this would be FO..POP + (WST..DIST(calculated above) / population)
# Enough to do it for EUN

x <- fdp %>% pivot_longer(starts_with("20"),names_to = "year",values_to = "value")

# Get FO..POP
x_fopop <- x %>% filter(region == "EUN", attribute == "FO..POP") %>% 
  group_by(region,product,year) %>% summarise( value = sum(value))

x_fopop <- ungroup(x_fopop)

# Get FO
x_fo <- x %>% filter(region == "EUN", attribute == "FO") %>% 
  group_by(region,product,year) %>% summarise( value = sum(value))

x_fo <- ungroup(x_fo)


# Calculate population as population = FO / FO..POP

