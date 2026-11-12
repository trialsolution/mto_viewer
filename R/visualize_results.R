library(readxl)
library(tidyverse)


# Prepare a tibble in the right format for the visualizations
# Here we take the big cube, filter to relevant results for dairy and then restructure

# load the result cube of interest
load("C:/Users/himicmi/Downloads/meltdown2026/data/cube_EUNmerge_20260924_17h28.xlsx_2026-09-25-08h40.Rdata")
load("C:/Users/himicmi/Downloads/meltdown2026/data/cube_EUNmerge_20261007_17h59.xlsx_2026-10-08-11h11.Rdata")

# update variable list from the viewer, if needed 
#extract_variable_list(viewer = viewer_file)

# if no additional variables needed, simply: load variable list from earlier runs -- my_selection
load(file = "data/my_selection.RData")

# filter the big cube
small_cube <- cube %>% right_join(my_selection, by = join_by(region,product,attribute))

# replace NAs with zeroes
x <- small_cube %>% pivot_longer(starts_with("20"),names_to = "year",values_to = "value")
x <- x %>% replace(is.na(.),0)
small_cube <- x %>% pivot_wider(names_from = "year", values_from = "value")


visuals <- small_cube %>% pivot_longer(starts_with("20"), names_to = "year", values_to = "value")
visuals$year <- as.numeric(visuals$year)


plot_timeserie <- function(reg, prod, att, mydata=visuals){
  
  p <- ggplot(data = mydata %>% filter(region == toupper(reg), product == toupper(prod), attribute == toupper(att)), 
              aes(x=year, y=value)) + geom_line() + ylim(0,NA) + theme_minimal()
  
  return(p)
}


plot_more_timeseries <- function(reg, prod, att, mygroup=attribute, mydata=visuals){
  
  mydata <- mydata %>% filter(region %in% toupper(reg), product %in% toupper(prod), attribute %in% toupper(att))
  
  p <- ggplot(data = mydata, 
#              aes(x=year, y=value, colour = mygroup)) + geom_line() + ylim(0,NA) + theme_minimal()
               aes(x=year, y=value, colour = mygroup)) + geom_line() +  theme_minimal()

  return(p)
}

plot_timeserie("EUN","FDP","NT", mydata = filter(visuals, year > 2023))
