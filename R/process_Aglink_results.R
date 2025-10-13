# script to read the huge merge file of the Aglink runs

library(readxl)
library(tidyverse)
library(xlsx)

viewer_file <- "DAIRY_viewer_2025.10.10_16h45.xlsx"
merge_file <- "EUN_EUNMERGE_10102025_16h45.xlsx"

# time stamp of running this script
# this time stamp will be used for all output files
# so we can identify later when the files will have been created
time_of_run <- format(Sys.time(), "%Y-%m-%d-%Hh%M")

# call functions for data processing
source("R/filter_results.R")


# read results
cube <- read_excel(paste("mergefiles/", merge_file, sep = ""), sheet = 1)

# name first columns
colnames(cube)[1] <- "region"
colnames(cube)[2] <- "product"
colnames(cube)[3] <- "attribute"

# remove years before 2000
cube <- as_tibble(cube)
cube <- cube %>% select(!starts_with("19"))

# convert code names to uppercase to avoid character mismatch later
cube <- cube %>% mutate(region=toupper(region), product=toupper(product), attribute=toupper(attribute))

# save the whole result cube for further use
save(cube, file = paste("data/cube_", merge_file, "_", time_of_run, ".Rdata", sep = ""))



# update variable list from the viewer, if needed 
#extract_variable_list(viewer = viewer_file)
# load variable list -- my_selection
load(file = "data/my_selection.RData")

# filter the big cube
small_cube <- cube %>% right_join(my_selection, by = join_by(region,product,attribute))

# replace NAs with zeroes
x <- small_cube %>% pivot_longer(starts_with("20"),names_to = "year",values_to = "value")
x <- x %>% replace(is.na(.),0)
small_cube <- x %>% pivot_wider(names_from = "year", values_from = "value")


# save to Excel
write.xlsx(as.data.frame(small_cube), file = paste("to_copy_", time_of_run, ".xlsx", sep = ""),
           row.names = FALSE, col.names = TRUE, sheetName = "merge",
           showNA = TRUE)

# add timestamp
timestamp <- paste("data copied from merge file ", merge_file, " on ", time_of_run, sep = "")
write.xlsx(timestamp, file = paste("to_copy_", time_of_run, ".xlsx", sep = ""), 
           row.names = FALSE, col.names = TRUE, sheetName = "timestamp_merge",
           showNA = TRUE, append = TRUE)

# append new data to the viewer file
# the new data will be added on a new, time-stamped sheet
write.xlsx(as.data.frame(small_cube), file = paste("viewers/", viewer_file, sep = ""),
           row.names = FALSE, col.names = TRUE, sheetName = paste("taken_", time_of_run, sep = ""),
           showNA = TRUE, append = TRUE)



# example code that also gives memory error with the viewer file...
#wb <- loadWorkbook(paste("viewers/", viewer_file, sep = ""))
#writeData(wb, sheet = "report", dfReport)
#saveWorkbook(wb,"Report.xlsx",overwrite = T)
