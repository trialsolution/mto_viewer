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
