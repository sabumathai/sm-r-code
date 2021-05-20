library(tibble)
library(tidyverse)
library(janitor)
library(dplyr)
library(formattable)
library(tidyr)
library(officer)
library(flextable)
library(writexl)
library(readxl)
library(customLayout)

options(digits=3)

setwd("/Users/sabumathai/3i Y5 - Revised Regional Decks/")

Y4 <- read_excel("3i Y5 - dataset.xlsx")
Y4 <- Y4[order(Y4[,1], Y4[,4]), ]

SlidesY4 <- Y4

SlidesY4 <- select(Y4, -c(7,10:13,17,21:89))

SlidesY4 <- select(SlidesY4, c(4,6,7,10, 9, 11:14,8,1:3,5))

colnames(SlidesY4) <- c("Country", "Internet users (% of households)", "Mobile subscribers (per 100 inhabitants; %)", "Average mobile download speed (Mbps)", "Average mobile upload speed (Mbps)", "Average mobile latency (ms)", "Network coverage (min. 2G) (% of population)", "Network coverage (min. 3G) (% of population)", "Network coverage (min. 4G) (% of population)", "Gender gap ratio in Internet access (% male – % female / % male)", "Region", "Core", "Income Group", "ISO")

SlideTableGlobal <- SlidesY4[,2:10]
SlideTableGlobalAvg <- colMeans(SlidesY4[,2:10])
SlideTableGlobalAvg <- data_frame(ID=names(SlidesY4[,2:10]), Global = SlideTableGlobalAvg)
SlideTableGlobalAvg

SlideTableLatAm <- SlidesY4[SlidesY4$Region=="Latam",2:10]
SlideTableLatamAvg <- colMeans(SlideTableLatAm[,1:9])
SlideTableLatamAvg <- data_frame(ID=names(SlidesY4[,2:10]), LatAm = SlideTableLatamAvg)
print(SlideTableLatamAvg)

SlideTableAsia <- SlidesY4[SlidesY4$Region=="Asia",2:10]
SlideTableAsiaAvg <- colMeans(SlideTableAsia[,1:9])
SlideTableAsiaAvg <- data_frame(ID=names(SlidesY4[,2:10]), Asia = SlideTableAsiaAvg)
print(SlideTableAsiaAvg)

SlideTableSSA <- SlidesY4[SlidesY4$Region=="SSA",2:10]
SlideTableSSAAvg <- colMeans(SlideTableSSA[,1:9])
SlideTableSSAAvg <- data_frame(ID=names(SlidesY4[,2:10]), SSA = SlideTableSSAAvg)
print(SlideTableSSAAvg)

SlideTableMENA <- SlidesY4[SlidesY4$Region=="MENA",2:10]
SlideTableMENAAvg <- colMeans(SlideTableMENA[,1:9])
SlideTableMENAAvg <- data_frame(ID=names(SlidesY4[,2:10]), MENA = SlideTableMENAAvg)
print(SlideTableMENAAvg)

SlideTableNA <- SlidesY4[SlidesY4$Region=="North America",2:10]
SlideTableNAAvg <- colMeans(SlideTableNA[,1:9])
SlideTableNAAvg <- data_frame(ID=names(SlidesY4[,2:10]), North_America = SlideTableNAAvg)
print(SlideTableNAAvg)

SlideTableEurope <- SlidesY4[SlidesY4$Region=="Europe",2:10]
SlideTableEuropeAvg <- colMeans(SlideTableEurope[,1:9])
SlideTableEuropeAvg <- data_frame(ID=names(SlidesY4[,2:10]), Europe = SlideTableEuropeAvg)
print(SlideTableEuropeAvg)

full_dplyr  <- full_join(SlideTableGlobalAvg, SlideTableNAAvg)
full_dplyr  <- full_join(full_dplyr, SlideTableEuropeAvg)
full_dplyr  <- full_join(full_dplyr, SlideTableMENAAvg)
full_dplyr  <- full_join(full_dplyr, SlideTableAsiaAvg)
full_dplyr  <- full_join(full_dplyr, SlideTableLatamAvg)
full_dplyr  <- full_join(full_dplyr, SlideTableSSAAvg)
print(full_dplyr)
print(colnames(full_dplyr))


#create country lists to use as counters in for loops

CountryList3iY4 <-SlidesY4[,1:11]
CountryList3iY4 <- CountryList3iY4[order(CountryList3iY4$Region),]
countries3iY4 <- unique(CountryList3iY4$Country)



# indicators as rows / regions and country as columns - cycle through countries in the last column
doc <- read_pptx(path = "3i_Y5_report_template.pptx")
doc <- on_slide( doc, index = 1)



for (i in seq_along(countries3iY4)) {

transformed_table_with_country <- full_dplyr
colnames(transformed_table_with_country) <- c("Indicator", "Global", "North America", "Europe", "MENA", "Asia", "Latin America", "SSA")

CountryList3iY4_transposed <- CountryList3iY4[CountryList3iY4$Country==countries3iY4[i], 1:10]

CountryList3iY4_transposed<- t(CountryList3iY4_transposed) %>% as_tibble()
CountryList3iY4_transposed <- CountryList3iY4_transposed %>%
  row_to_names(row_number = 1)
transformed_table_with_country <- add_column(transformed_table_with_country, type.convert(CountryList3iY4_transposed))

if (CountryList3iY4$Region[i]=="Asia"){
  transformed_table_with_country <- transformed_table_with_country[, c(1, 2, 6, 9, 3, 4, 5, 7, 8)]
} else if (CountryList3iY4$Region[i]=="Europe"){
  transformed_table_with_country <- transformed_table_with_country[, c(1, 2, 4, 9, 3, 5, 6, 7, 8)]
} else if (CountryList3iY4$Region[i]=="Latam"){
  transformed_table_with_country <- transformed_table_with_country[, c(1, 2, 7, 9, 3, 4, 5, 6, 8)]
} else if (CountryList3iY4$Region[i]=="MENA"){
  transformed_table_with_country <- transformed_table_with_country[, c(1, 2, 5, 9, 3, 4,  6, 7, 8)]
} else if (CountryList3iY4$Region[i]=="North America"){
  transformed_table_with_country <- transformed_table_with_country[, c(1, 2, 3, 9, 4, 5, 6, 7, 8)]
} else if (CountryList3iY4$Region[i]=="SSA"){
  transformed_table_with_country <- transformed_table_with_country[, c(1, 2, 8, 9, 3, 4, 5, 6, 7)]
  } 

if (CountryList3iY4$Region[i]=="Asia"){
  regionName <- "Asia"
} else if (CountryList3iY4$Region[i]=="Europe"){
  regionName <- "Europe"
} else if (CountryList3iY4$Region[i]=="Latam"){
  regionName <- "Latin America"
} else if (CountryList3iY4$Region[i]=="MENA"){
  regionName <- "MENA"
} else if (CountryList3iY4$Region[i]=="North America"){
  regionName <- "North America"
} else if (CountryList3iY4$Region[i]=="SSA"){
  regionName <- "SSA"
} 

transformed_table_with_country$Indicator <- c("Internet users                                \n(% of households)", "Mobile subscribers                     \n(per 100 inhabitants; %)", "Average mobile download speed\n(Mbps)                            ", "Average mobile upload speed\n(Mbps)                              ", "Average mobile latency\n(ms)                                     ", "Network coverage                    \n(min. 2G) (% of population)", "Network coverage                    \n(min. 3G) (% of population)", "Network coverage                    \n(min. 4G) (% of population)", "Gender gap ratio in Internet access \n(% male – % female /% male)")



flexty <- flextable(transformed_table_with_country)
flexty <- bold(flexty, part = "header")
flexty <- bold(flexty, j= ~ Indicator, part = "body")

flexty <- bold(flexty, j= regionName, part = "body")
flexty <- bold(flexty, j= countries3iY4[i], part = "body")

flexty <- fontsize(flexty, part = "header", size = 11)
flexty <- fontsize(flexty, part = "body", j = 2:9, size = 14)
# dark blue as background color for header
flexty <-  bg(flexty, bg = "#2B458F", part = "header")
# slate as background color for body
flexty <-  bg(flexty, bg = "#EFEDE5", part = "body")
# dark blue as background color for Indicator column
flexty <-  bg(flexty, j = ~ Indicator, bg = "#2B458F", part = "body")

# light blue color to highlight region and country
flexty <-  bg(flexty, j = regionName, bg = "#D2E0F2", part = "body")
flexty <-  bg(flexty, j = countries3iY4[i], bg = "#D2E0F2", part = "body")

# flexty <- autofit(flexty)
flexty <- width( flexty, width = 1.1 )
flexty <- width(flexty, j = ~ Indicator, width = 3)
flexty <- height_all( flexty, height = .5 )
flexty <- align(flexty, align = "left", part = "all")
flexty <- color(flexty, part = "header", color = "#FFFFFF")
flexty <- color(flexty, j = ~ Indicator, part = "body", color = "#FFFFFF")
flexty <- colformat_double(
  x = flexty,
  big.mark=",", digits = 1, na_str = "N/A")
flexty <- colformat_double(
  x = flexty,
  big.mark=",", i = 3, digits = 0, na_str = "N/A")
flexty <- colformat_double(
  x = flexty,
  big.mark=",", i = 4, digits = 0, na_str = "N/A")
flexty <- border_remove(x = flexty)
flexty <- border_outer(flexty, part="all", border = fp_border(color="white", width = 1) )
flexty <- border_inner_h(flexty, part="all", border = fp_border(color="white", width = 1) )
flexty <- border_inner_v(flexty, part="all", border = fp_border(color="white", width = 1) )


# print(transformed_table_with_country)
doc <- add_slide(doc)

doc <- ph_with(doc, value = paste(countries3iY4[i], "'s metrics", sep = ""), location = ph_location_type(type = "title"))
doc <- ph_with(doc, value = flexty,
               location = ph_location(left = .77, top = 1.31))


}


print(doc, target = "3i_Y5_Briefings.pptx")


# print(transformed_table_with_country)









