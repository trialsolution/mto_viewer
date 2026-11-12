# script to retrive the results on import dependency for selected countries and products
# (IM-EX)/(QC)		NETIMDEP

library(readxl)
library(tidyverse)
library(xlsx)


# note that you need to create the viewer file manually
# by duplicating and renaming the previous viewer
viewer_file <- "DAIRY_viewer_2025.11.20_16h50.xlsx"
merge_file <- "EUN_EUNMERGE_20112025_16h50.xlsx"

# time stamp of running this script
# this time stamp will be used for all output files
# so we can identify later when the files will have been created
time_of_run <- format(Sys.time(), "%Y-%m-%d-%Hh%M")


load("C:/Users/himicmi/Downloads/meltdown2025/data/cube_EUN_EUNMERGE_20112025_16h50.xlsx_2025-11-24-15h14.Rdata")

# I. import dependency
#----
x <- cube %>% filter(attribute %in% c("NETIMDEP"))

# replace NAs with zeroes
x <- x %>% pivot_longer(starts_with("20"),names_to = "year",values_to = "value")
x <- x %>% replace(is.na(.),0)
x <- x %>% pivot_wider(names_from = "year", values_from = "value")


# save to Excel
write.xlsx(as.data.frame(x), file = paste("reporting/import_dependency_", time_of_run, ".xlsx", sep = ""),
           row.names = FALSE, col.names = TRUE, sheetName = "import_dependency",
           showNA = TRUE)