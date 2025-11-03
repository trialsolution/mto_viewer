# extract and analyse milk yields from the master file

library(readxl)
library(tidyverse)
library(xlsx)


# All EU regions
mkc_eu <- read_excel("Y:/work/EUNDAIRY.xlsx", sheet = "EUNDAIRYDB", range = "M583:BD701", 
                      col_names = FALSE, col_types = c(rep("text",3), rep("numeric",41)))
mkc_eu <- as_tibble(mkc_eu)

# name first columns
colnames(mkc_eu)[1] <- "region"
colnames(mkc_eu)[2] <- "product"
colnames(mkc_eu)[3] <- "attribute"

for (i in 4:44) {
  colnames(mkc_eu)[i] <- 1996+i
}

# remove empty lines
mkc_eu <- mkc_eu %>% filter(!is.na(region) | !is.na(product) | !is.na(attribute))


mkc_eu <- mkc_eu %>% mutate(code = paste(region,product,attribute, sep = "_"))
mkc_eu <- mkc_eu %>% select(-region,-product,-attribute)
mkc_eu <- mkc_eu %>% select(-`2000`)


# save raw milk-related projections from the master file
time_of_run <- format(Sys.time(), "%Y-%m-%d-%Hh%M")
save(mkc_eu, file = paste("data/MF_mkc_", time_of_run, ".Rdata", sep = ""))


# convert to time series
x <- mkc_eu %>% pivot_longer(starts_with("20"), names_to = "year", values_to = "value")

# have a look at EUN milk yields
yield <- x %>% filter(code == "EUN_MKC_YLD")

# check in E14 and NMS by replacing 'yield' with another filtering of the x cube
yield <- x %>% filter(code == "E14_MKC_YLD")
yield <- x %>% filter(code == "NMS_MKC_YLD")


series <- ts(yield$value, frequency = 1, start = 2001)

# plot and check the diffs over time
plot(series)

# annual increase in yields; kg/cow
diff(series, lag=1)*1000

# plotting annual differences
plot(diff(series, lag=1)*1000)


# check the relative change y.o.y.
# note here the logarithmic formula to calculate relative change over time
plot((exp(diff(log(series)))-1)*100)

# check yield development in historic years
historic <- ts(series, start = 2005, end = 2025)
plot(diff(historic))

# calculate milk development in historic versus projected years

# mean difference
yield %>% filter(year<2026, year>2005) %>% group_by(code) %>% summarise(value = mean(diff(value)))
yield %>% filter(year>2025) %>% group_by(code) %>% summarise(value = mean(diff(value)))

# median difference
yield %>% filter(year<2026, year>2009) %>% group_by(code) %>% summarise(value = median(diff(value)))
yield %>% filter(year>2025) %>% group_by(code) %>% summarise(value = median(diff(value)))

# mean relative difference
yield %>% filter(year<2026, year>2009) %>% group_by(code) %>% summarise(value = mean((exp(diff(log(value)))-1)*100))
yield %>% filter(year>2025) %>% group_by(code) %>% summarise(value = mean((exp(diff(log(value)))-1)*100))


# calulate average annual growth rate
# with different base years (2015, 2025, ...), or base years that are three-year averages

bv <- filter(yield,year %in% c("2013","2014","2015")) %>% select(code, year, value)
bv <- bv %>% group_by(code) %>% summarise(value=mean(value))
yield <- yield %>% mutate(growth_rate_2015_avg = (100*exp(log(value/bv[['value']])/(as.numeric(year)-2015)))-100)

bv <- filter(yield,year %in% c("2023","2024","2025")) %>% select(code, year, value)
bv <- bv %>% group_by(code) %>% summarise(value=mean(value))
yield <- yield %>% mutate(growth_rate_2025_avg = (100*exp(log(value/bv[['value']])/(as.numeric(year)-2025)))-100)

yield <- yield %>% mutate(growth_rate_2015 = (100*exp(log(value/filter(yield,year=="2015")[['value']])/(as.numeric(year)-2015)))-100)
yield <- yield %>% mutate(growth_rate_2025 = (100*exp(log(value/filter(yield,year=="2025")[['value']])/(as.numeric(year)-2025)))-100)

# annual absolute increase in milk yields
yield <- yield %>% mutate(growth_rate_2025 = (100*exp(log(value/filter(yield,year=="2025")[['value']])/(as.numeric(year)-2025)))-100)



