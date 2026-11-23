# script to read the huge merge file of the Aglink runs

library(readxl)
library(tidyverse)
library(xlsx)


# EUN
# The Dairy master file is located on a shared drive (mapped to Y:)
# both JRC and AGRI has access to this drive, which is copied over to the virtual machines for Aglink runs
#
fdp_eun <- read_excel("./MF/EUNDAIRY final MTO 2026.xlsx", sheet = "EUNDAIRYDB", range = "M457:BD483", 
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

fdp_nms <- read_excel("./MF/EUNDAIRY final MTO 2026.xlsx", sheet = "EUNDAIRYDB", range = "M487:BD516", 
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

fdp_e14 <- read_excel("./MF/EUNDAIRY final MTO 2026.xlsx", sheet = "EUNDAIRYDB", range = "M523:BD555", 
                      col_names = FALSE, col_types = c(rep("text",3), rep("numeric",41)))
fdp_e14 <- as_tibble(fdp_e14)

# name first columns
colnames(fdp_e14)[1] <- "region"
colnames(fdp_e14)[2] <- "product"
colnames(fdp_e14)[3] <- "attribute"

for (i in 4:44) {
  colnames(fdp_e14)[i] <- 1996+i
}

# put the 3 regions into one tibble
fdp <- rbind(fdp_eun, fdp_e14, fdp_nms)
rm(fdp_e14, fdp_nms)


# save the data on fresh dairy commodities in R format
time_of_run <- format(Sys.time(), "%Y-%m-%d-%Hh%M")
save(fdp, file = paste("data/fdp_", time_of_run, ".Rdata", sep = ""))



# Alternatively, load FDP products data from the .RData file
# (if previously processed)
#load("data/fdp_2025-10-16-12h15.Rdata")

# time stamp
#time_of_run <- format(Sys.time(), "%Y-%m-%d-%Hh%M")



# calculate QC = QP-NT
x <- fdp %>% pivot_longer(starts_with("20"),names_to = "year",values_to = "value")

# QP calculated for EUN (sum of NMS and E14)
x_qp <- x %>% filter(region != "EUN", attribute == "QP") %>% group_by(product,year) %>% summarise( value = sum(value))
x_qp <- ungroup(x_qp)
# NT calculated for EUN (sum of NMS and E14)
x_nt <- x %>% filter(region != "EUN", attribute == "NT") %>% group_by(product,year) %>% summarise( value = sum(value))
x_nt <- ungroup(x_nt)




# check if and how FPD production and net trade are different in the MF and in the final merge results
# All this difference will be shifted to the 'Other' product, which will be re-calculated as a remainder: FDP - Cream - Yogurt - Milk

#
# load merge results for FDP
#
load("data/cube_EUNmerge_20261104_10h47.xlsx_2026-11-04-11h21.Rdata")

# load variable list from earlier runs -- my_selection
load(file = "data/my_selection.RData")

# filter the big cube and drop it
small_cube <- cube %>% right_join(my_selection, by = join_by(region,product,attribute))
rm(cube)

# replace NAs with zeroes
x <- small_cube %>% pivot_longer(starts_with("20"),names_to = "year",values_to = "value")
x <- x %>% replace(is.na(.),0)
small_cube <- x %>% pivot_wider(names_from = "year", values_from = "value")
rm(x)

# get FDP_QP and FDP_NT
merge_qp <- small_cube %>% filter(region == "EUN", product == "FDP", attribute == "QP")
merge_nt <- small_cube %>% filter(region == "EUN", product == "FDP", attribute == "NT")

m_qp <- merge_qp %>% pivot_longer(starts_with("20"), names_to = "year", values_to = "value")
m_qp <- m_qp %>% select(-region, -attribute)


# compare QP in merge vs. master file (MF)
m_qp <- m_qp %>% pivot_wider(names_from = "product")
colnames(m_qp)[2] <- "FDP_merge"

# get data from the MF
MF_qp <- x_qp %>% filter(product == "FDP") %>% pivot_wider(names_from = "product")

compare <- m_qp %>% full_join(MF_qp, by = join_by(year == year))

# calculate and show difference and relative difference
compare %>% mutate(diff = FDP_merge - FDP, perc_diff = (FDP_merge - FDP)/abs(FDP)*100) %>% view()


# get m_qp again from the raw results and filter as needed
m_qp <- merge_qp %>% pivot_longer(starts_with("20"), names_to = "year", values_to = "value")
m_qp <- m_qp %>% select(-region, -attribute)

# replace FDP in x_qp with the values from the merge file (simulation result)
x_qp <- x_qp %>% filter(product!="FDP")
x_qp <- rbind(x_qp, m_qp)

# correct QP of 'Other' to achieve consistency between FDP and the sub-categories
x_qp <- x_qp %>% pivot_wider(names_from = "product")
x_qp <- x_qp %>% mutate(Other = FDP - Cream - Milk - Yogurt)
x_qp <- x_qp %>% pivot_longer(cols = c("Cream", "Milk", "Other", "Yogurt", "FDP"), names_to = "product", values_to = "value")


# do the same for NT
m_nt <- merge_nt %>% pivot_longer(starts_with("20"), names_to = "year", values_to = "value")
m_nt <- m_nt %>% select(-region, -attribute)

# get MF data for NT
MF_nt <- x_nt %>% filter(product == "FDP") %>% pivot_wider(names_from = "product")

# replace FDP in x_nt with the values from the merge file (simulation result)
x_nt <- x_nt %>% filter(product!="FDP")
x_nt <- rbind(x_nt, m_nt)

# correct NT to achieve consistency between FDP and the sub-categories
x_nt <- x_nt %>% pivot_wider(names_from = "product")
x_nt <- x_nt %>% mutate(Milk = FDP - Cream - Other - Yogurt)
x_nt <- x_nt %>% pivot_longer(cols = c("Cream", "Milk", "Other", "Yogurt", "FDP"), names_to = "product", values_to = "value")


# calculate consumption

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



#
#
# PART II.
#
# Calculate FOA, by adding waste to FO
#
# this would be TRUE..FO + share * FDP_WST..DIST
# the shares are NMS/E14 specific
#
# Note that the final FDP_FO, FDP_FOA might differ from the ones in the master file
# => we need to correct the FO/FOA of the FDP-subcategires to achieve consistency (adding-up)
#

#
# Note that 'fdp' contains data from the MF
#
x <- fdp %>% pivot_longer(starts_with("20"),names_to = "year",values_to = "value")

# Get TRUE..FO
x_truefo <- x %>% filter(region != "EUN", attribute == "TRUE..FO") %>% 
  group_by(region,product,year) %>% summarise( value = sum(value))

x_truefo <- ungroup(x_truefo)

#TODO: recalculate TRUE..FO based on FDP_FO from the mergefile
# e.g. allocate difference between MF vs. mergefile for FDP_FO to 'Other'_TRUE..FO
m_fdp_fo <- cube %>% filter(region %in% c("E14","NMS"), product == "FDP", attribute == "FO")

m_fdp_fo <- m_fdp_fo %>% pivot_longer(starts_with("20"), names_to = "year", values_to = "value")

m_fdp_fo <- m_fdp_fo %>% select(-attribute)



# merge and recalulate 'Other' FO
x_truefo <- rbind(x_truefo, m_fdp_fo)

x_truefo <- x_truefo %>% pivot_wider(names_from = "product")

#error trap
x_truefo %>% mutate(Other = FDP - Cream - Milk -Yogurt) %>% filter(Other < 0)

# re-calculation
x_truefo <- x_truefo %>% mutate(Other = FDP - Cream - Milk -Yogurt) 



# Get WST..DIST from the Master file for FDP total
x_wstdist <- x %>% filter(region != "EUN", attribute == "WST..DIST") %>% 
  group_by(region,product,year) %>% summarise( value = sum(value))

x_wstdist <- ungroup(x_wstdist)



#TODO: WST..DIST for FDP and for E14/NMS can be read from the i.e. from mergefile
#      FOA       for FDP and for E14/NMS can be read from the i.e. from mergefile
# note that this is not always part of my_selection!
m_fdp_wst <- cube %>% filter(region %in% c("E14","NMS"), product == "FDP", attribute == "WST..DIST")

# convert to longer format
m_fdp_wst <- m_fdp_wst %>% pivot_longer(starts_with("20"), names_to = "year", values_to = "value")


#
# Calculate the waste for the fresh dairy products
# The waste share coefficients are those used in the EUNDAIRY master file
# The coefficients allocate the total waste of Fresh dairy to the different fresh dairy subcategories:
# -  Cream: 8%
# -  Drinking milk: 60%
# -  Yogurt: 30%
# -  The category 'Other' will close the balance, i.e. its waste share is 100% - sum(cream, drinking m., yogurt)
#
# These waste shares are hard coded in the MF for the calculation of TRUE..FO in the historical years
# At the same time, note that the MF file does not have waste calculated for the FDP-subcategories
#
# Be aware:  
#           - need to update this coeffs here manually if necessary
#           - coeffs are different for the EU regions
#

calculate_waste <- function(fdp_data = fdp, fdp_product = "Cream", waste_share = 0.08, eu_region){
  
  x_wstdist <- x %>% filter(region != "EUN", attribute == "WST..DIST", region == eu_region) %>% 
    group_by(region,product,year) %>% summarise( value = sum(value))
  x_wstdist <- ungroup(x_wstdist)  
  
  waste <- x_wstdist
  waste$product <- fdp_product
  waste <- waste %>% mutate(value = waste_share*value)
  
  return(waste)  
}

waste_dmilk_nms <- calculate_waste(fdp_data = m_fdp_wst, fdp_product = "Milk",  waste_share = 0.6, eu_region = "NMS")
waste_cream_nms <- calculate_waste(fdp_data = m_fdp_wst, fdp_product = "Cream", waste_share = 0.08, eu_region = "NMS")
waste_yogurt_nms <- calculate_waste(fdp_data = m_fdp_wst, fdp_product = "Yogurt", waste_share = 0.3, eu_region = "NMS")

waste_dmilk_e14 <- calculate_waste(fdp_data = m_fdp_wst, fdp_product = "Milk",  waste_share = 0.6, eu_region = "E14")
waste_cream_e14 <- calculate_waste(fdp_data = m_fdp_wst, fdp_product = "Cream", waste_share = 0.05, eu_region = "E14")
waste_yogurt_e14 <- calculate_waste(fdp_data = m_fdp_wst, fdp_product = "Yogurt", waste_share = 0.2, eu_region = "E14")


# put it together in one parameter
waste_fdp_products <- rbind(waste_dmilk_nms, waste_cream_nms, waste_yogurt_nms, 
                            waste_dmilk_e14, waste_cream_e14, waste_yogurt_e14)

rm(waste_dmilk_nms, waste_cream_nms, waste_yogurt_nms, 
   waste_dmilk_e14, waste_cream_e14, waste_yogurt_e14)


#calculating the waste of 'Other' we can use the remainder approach


#merge total FDP waste
waste_fdp_products <- rbind(waste_fdp_products, m_fdp_wst %>% select(-attribute))


#reformat and calculate 'Other'
waste_fdp_products <- waste_fdp_products %>% pivot_wider(names_from = "product")
waste_fdp_products <- waste_fdp_products %>% mutate(Other = FDP - Milk - Cream - Yogurt)

#error trap -- check if there are negatives
waste_fdp_products %>% filter(Other < 0)



#TODO: calculate FOA = TRUE..FO + WST..DIST
# using waste_fdp_products and x_truefo

colnames(x_truefo) <- c("region", "year", "Cream_FO", "Milk_FO", "Other_FO", "Yogurt_FO", "FDP_FO")
colnames(waste_fdp_products) <- c("region", "year", "Milk_WST", "Cream_WST", "Yogurt_WST", "FDP_WST", "Other_WST")

fdps_foa <- x_truefo %>% full_join(waste_fdp_products, by = join_by("region", "year"))

fdps_foa <- fdps_foa %>% mutate(Cream = Cream_FO+Cream_WST,
                    Milk = Milk_FO+Milk_WST,
                    Yogurt = Yogurt_FO+Yogurt_WST,
                    Other = Other_FO+Other_WST,
                    FDP = FDP_FO+FDP_WST)

fdps_foa <- fdps_foa %>% select(-Cream_FO,-Cream_WST,-Milk_FO,-Milk_WST,
                                -Yogurt_FO,-Yogurt_WST,-Other_FO,-Other_WST,
                                -FDP_FO,-FDP_WST)



#aggregate to EUN
fdps_foa <- fdps_foa %>% pivot_wider(names_from = product)
fdps_foa <- fdps_foa %>% group_by(year) %>% summarise(Cream=sum(Cream), Milk=sum(Milk), Other=sum(Other), Yogurt=sum(Yogurt), FDP=sum(FDP))
fdps_foa <- fdps_foa %>% mutate(region = "EUN")

fdps_foa <- fdps_foa %>% pivot_longer(cols = c("Cream", "Milk", "Other", "Yogurt", "FDP"), names_to = "product", values_to = "value")


#write to Excel
to_excel <- fdps_foa %>% pivot_wider(names_from = year)

write.xlsx(as.data.frame(to_excel), file = paste("reporting/fdps_foa_", time_of_run, ".xlsx", sep = ""),
           row.names = FALSE, col.names = TRUE, sheetName = "FOA",
           showNA = TRUE)



# Calculate food availability (FOA) per capita

eun_pop <- cube %>% filter(region == "EUN", product == "ME", attribute == "POP")
eun_pop <- eun_pop %>% pivot_longer(starts_with("20"),names_to = "year",values_to = "value")
eun_pop <- eun_pop %>% select(year, region, product, value)

fdps_foa_pop <- rbind(fdps_foa, eun_pop)


# calculate per capita FOA
fdps_foa_pop <- fdps_foa_pop %>% pivot_wider(names_from = product)

fdps_foa_pop <- fdps_foa_pop %>% mutate(Cream = Cream/ME*1000,
                                        Milk = Milk/ME*1000,
                                        Other = Other/ME*1000,
                                        Yogurt = Yogurt/ME*1000,
                                        FDP = FDP/ME*1000
                                        )

fdps_foa_pop <- fdps_foa_pop %>% select(-ME)

fdps_foa_pop <- fdps_foa_pop %>% 
  pivot_longer(cols = c("Cream", "Milk", "Other", "Yogurt", "FDP"), 
               names_to = "product", values_to = "value")

#write to Excel
to_excel <- fdps_foa_pop %>% pivot_wider(names_from = year)

write.xlsx(as.data.frame(to_excel), file = paste("reporting/fdps_foa_pop_", time_of_run, ".xlsx", sep = ""),
           row.names = FALSE, col.names = TRUE, sheetName = "FOA..POP",
           showNA = TRUE)


# calculate ST and PRST for butter and SMP
#----------------------------------------
load("C:/Users/himicmi/Downloads/meltdown2025/data/cube_EUN_EUNMERGE_20112025_16h50.xlsx_2025-11-24-15h14.Rdata")


# butter ending stock in 2005
bt_end_2005 <- 99.82


stocks <- cube %>% filter(region == "EUN", product %in% c("BT","SMP"), attribute %in% c("VST", "VST..PRST", "IST", "ST")) 

bt_stocks <- stocks %>% filter(product == "BT")

start_value <- bt_end_2005 -as.numeric(filter(bt_stocks, attribute == "IST")[,"2005"])-as.numeric(filter(bt_stocks, attribute == "VST..PRST")[,"2005"])

bt_stocks <- bt_stocks %>% pivot_longer(starts_with("20"), names_to = "year")
bt_stocks$year <- as.numeric(bt_stocks$year)
bt_stocks <- bt_stocks %>% filter(year>2004)

#replace NAs with zero for the calcuation
bt_stocks <- bt_stocks %>% replace(is.na(.),0)

bt_stocks <- bt_stocks %>% pivot_wider(names_from = "attribute")

bt_stocks <- bt_stocks %>% mutate(PRST = as.numeric(start_value) + cumsum(VST..PRST))


# check if calculation matches ST in the results 
bt_stocks <- bt_stocks %>% mutate(ST_new = IST + PRST)
bt_stocks <- bt_stocks %>% mutate(check = ST - ST_new)



# reformat before saving to Excel
bt_stocks <- bt_stocks %>% select(region, product, year, ST_new, IST, VST..PRST, PRST)

bt_stocks <- bt_stocks %>% pivot_longer(cols=c("ST_new","IST","VST..PRST","PRST"), names_to = "attribute", values_to = "value")

#write to Excel
to_excel <- bt_stocks %>% pivot_wider(names_from = year)

write.xlsx(as.data.frame(to_excel), file = paste("reporting/bt_prst", time_of_run, ".xlsx", sep = ""),
           row.names = FALSE, col.names = TRUE, sheetName = "bt_prst",
           showNA = TRUE)


# SMP stocks
# ---
# SMP_PRST(2005) = 72.40 (1000 t)

smp_prst_2005 <- 72.4

smp_stocks <- stocks %>% filter(product == "SMP")

start_value <- smp_prst_2005 -as.numeric(filter(smp_stocks, attribute == "VST..PRST")[,"2005"])

smp_stocks <- smp_stocks %>% pivot_longer(starts_with("20"), names_to = "year")
smp_stocks$year <- as.numeric(smp_stocks$year)
smp_stocks <- smp_stocks %>% filter(year>2004)

#replace NAs with zero for the calcuation
smp_stocks <- smp_stocks %>% replace(is.na(.),0)

smp_stocks <- smp_stocks %>% pivot_wider(names_from = "attribute")

smp_stocks <- smp_stocks %>% mutate(PRST = as.numeric(start_value) + cumsum(VST..PRST))


# check if calculation matches ST in the results 
smp_stocks <- smp_stocks %>% mutate(ST= IST + PRST)


# reformat before saving to Excel
smp_stocks <- smp_stocks %>% select(region, product, year, VST, IST, VST..PRST, PRST, ST)

smp_stocks <- smp_stocks %>% pivot_longer(cols=c("VST","IST","VST..PRST","PRST", "ST"), names_to = "attribute", values_to = "value")

#write to Excel
to_excel <- smp_stocks %>% pivot_wider(names_from = year)

write.xlsx(as.data.frame(to_excel), file = paste("reporting/smp_st", time_of_run, ".xlsx", sep = ""),
           row.names = FALSE, col.names = TRUE, sheetName = "SMP_ST",
           showNA = TRUE)




# EXAMPLE for cumsum (for the calculation of stocks)

df <- data.frame(
  year = 2020:2024,
  change = c(0, 5, -2, 8, 4) # 0 change in the first year
)

start_val <- 100

df <- df %>%
  mutate(value = start_val + cumsum(change))


