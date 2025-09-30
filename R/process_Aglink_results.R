# script to read the huge merge file of the Aglink runs

library(readxl)
library(tidyverse)
library(xlsx)

viewer_file <- "DAIRY_viewer_2025.09.29_16h16.xlsx"
merge_file <- "EUN_EUNMERGE_29092025_16h16.xlsx"

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


# function to extract Aglink  variable list from Excel
# only the variables in use are needed from the huge merge file
extract_variable_list <- function(viewer){
  
  # Read the variables to filter
  my_selection <- read_excel(paste("viewers/", viewer, sep = ""), sheet = "relevantRes", range = cell_cols("C:E"))
  my_selection <- as_tibble(my_selection)
  
  # remove empty lines
  my_selection <- my_selection %>% filter(!is.na(region))
  # remove calculations not available in the result cube
  my_selection <- my_selection %>% filter(region!="otherASIA")
  my_selection <- my_selection %>% filter(region!="NAFR")
  my_selection <- my_selection %>% filter(region!="TMLE")
  # convert names to uppercase to avoid character matching issues in the join
  my_selection <- my_selection %>% mutate(region=toupper(region), product=toupper(product), attribute=toupper(attribute))
  
  # save the list of relevant Aglink variables
  save(my_selection, file="data/my_selection.RData")
  
}

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
write.xlsx(as.data.frame(small_cube), file = paste("to_copy_", format(Sys.time(), "%Y-%m-%d"), ".xlsx", sep = ""),
           row.names = FALSE, col.names = TRUE, sheetName = "merge",
           showNA = TRUE)

# add timestamp
timestamp <- format(Sys.time(), "data copied from merge file on %Y.%m.%d-%H:%M:%S")
write.xlsx(timestamp, file = paste("to_copy_", format(Sys.time(), "%Y-%m-%d"), ".xlsx", sep = ""), 
           row.names = FALSE, col.names = TRUE, sheetName = "timestamp_merge",
           showNA = TRUE, append = TRUE)
