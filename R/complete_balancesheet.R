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

# time stamp
time_of_run <- format(Sys.time(), "%Y-%m-%d-%Hh%M")


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

