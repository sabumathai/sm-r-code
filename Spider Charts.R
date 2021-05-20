library(tibble)
library(tidyverse)
library(rworldmap)
library(ggthemes)
library(ggrepel)
library(janitor)
library(reactable)
library(data.table)
library(dplyr)
library(formattable)
library(tidyr)
library(officer)
library(formattable)
library(radarchart)
library(dplyr)
library(scales)
library(tibble)
library(fmsb)
library(readxl)

# Define fill colors
colors_fill <- c(scales::alpha("red", 0.1),
                 scales::alpha("blue", 0.1),
                 scales::alpha("orange", 0.1),
                 scales::alpha("yellow", 0.1),
                 scales::alpha("lightgreen", 0.1),
                 scales::alpha("purple", 0.1),
                 scales::alpha("gray", 0.1),
                 scales::alpha("cyan", 0.1))

# Define line colors
colors_line <-  c(scales::alpha("darkred", 0.9),
                  scales::alpha("darkblue", 0.9),
                  scales::alpha("darkorange", 0.9),
                  scales::alpha("gold", 0.9),
                  scales::alpha("darkgreen", 0.9),
                  scales::alpha("purple", 0.9),
                  scales::alpha("darkgray", 0.9),
                  scales::alpha("cyan", 0.9))

setwd("/Users/sabumathai/Y5_Radar_Charts/")
doc <- read_pptx(path = "3i_Y5_report_template.pptx")

all_Scores <- read_excel("III_Fifth_Edition_Scores_(DEFAULT_weight_profile).xlsx")
region_groups <- select(all_Scores, c(1, 2, 5, 11, 19, 28,32,37,42,46,56, 61,	68))
colnames(region_groups) <- c("Country", "Edition", "Usage", "Quality", "Infrastructure      ", "Electricity   ", "Price", "Competitive\nEnvironment", "Local\nContent", "  Relevant\nContent", "Literacy", "Trust &\nSafety", "Policy")

country_groups <- read_excel("Country Regions and Income Groups.xlsx")

region_groups <- left_join(country_groups, region_groups, by = "Country")



Edition_order <- c("E5", "E4", "E3", "E2", "E1")

Region_order <- c("North America", "Europe", "Middle East and North Africa", "Asia", "Latin America", "Sub-Saharan Africa")

region_groups$Region <- factor(region_groups$Region, ordered = TRUE, 
                                levels = Region_order)

region_groups$Edition <- factor(region_groups$Edition, ordered = TRUE, 
                                 levels = Edition_order)
region_groups<- region_groups[order(region_groups$Edition), ]



Income_Group_order <- c("High income", "Upper middle income", "Lower middle income", "Low income")

region_groups$Income_Group <- factor(region_groups$Income_Group, ordered = TRUE,
                                     levels = Income_Group_order)
region_groups_IG <- region_groups
region_groups_IG<- region_groups_IG[order(region_groups_IG$Income_Group), ]


# region_groups<- region_groups[region_groups$Edition=="E5" | region_groups$Edition=="E4", ]

region_sum <- region_groups[region_groups$Edition == "E5",] %>%
  arrange(Region) %>%
  group_by(Region) %>%
  summarise(across(where(is.numeric), mean, na.rm = TRUE))
rownames(region_sum) <- region_sum$Region
region_sum <- select(region_sum, -Region)

region_sum <- rbind(rep(100, 12), rep(0, 12), region_sum)
mypath <- paste("Y5_Radar_Charts/All Regions 3i Y5 Radar Chart.png")
png(file = mypath, width = 900, height = 465)
radarchart(region_sum, # title = Y4_Country_List[i],
           seg = 10,  # Number of axis segments
           pcol = colors_line,
           plwd = 2,
           plty = 1
           )
legend(1.5,1,
       legend=c("North America", "Europe", "Middle East and North Africa", "Asia", "Latin America", "Sub-Saharan Africa"),
       pch=c(15,16),
       col=colors_line,
       lty=1)
dev.off()


doc <- add_slide(doc)
doc <- ph_with(doc, value = "Average Scores by Region", location = ph_location_type(type = "title"))

doc <- ph_with(doc, value = external_img(mypath),
               location = ph_location(left = 1.39, top = 1.15, width = 10.22, height = 5.7))


income_sum <- region_groups_IG %>%
  group_by(Income_Group) %>%
  summarise(across(where(is.numeric), mean, na.rm = TRUE))
rownames(income_sum) <- income_sum$Income_Group
income_sum <- select(income_sum, -Income_Group)

income_sum <- rbind(rep(100, 12), rep(0, 12), income_sum)
mypath <- paste("Y5_Radar_Charts/Income Group 3i Y5 Radar Chart.png")
png(file = mypath, width = 900, height = 550)
radarchart(income_sum, # title = Y4_Country_List[i],
           seg = 10,  # Number of axis segments
           pcol = colors_line,
           plwd = 2,
           plty = 1
)
legend(1.4,1,
       legend=c("High income", "Upper middle income", "Lower middle income", "Low income"),
       pch=c(15,16),
       col=colors_line,
       lty=1)
dev.off()


doc <- add_slide(doc)
doc <- ph_with(doc, value = "Average Scores by Income Group", location = ph_location_type(type = "title"))

doc <- ph_with(doc, value = external_img(mypath),
               location = ph_location(left = 1.47, top = 1.07, width = 9.21, height =5.75))


LA_group <- region_groups[region_groups$Region=="Latin America" & region_groups$Edition != "E1" & region_groups$Edition != "E2" & region_groups$Edition != "E3",]
LA_sum <- LA_group %>%
  group_by(Edition) %>%
  summarise(across(where(is.numeric), mean, na.rm = TRUE))
rownames(LA_sum) <- LA_sum$Edition
LA_sum <- select(LA_sum, -Edition)
LA_sum <- rbind(rep(100, 12), rep(0, 12), LA_sum)
mypath <- paste("Y5_Radar_Charts/Latin America 3i Y5 Radar Chart.png")
png(file = mypath)
radarchart(LA_sum, # title = Y4_Country_List[i],
           seg = 10,  # Number of axis segments
           pcol = colors_line,
           pfcol = colors_fill,
           plwd = 2,
           plty = 1)
legend(1,1,
       legend=c("E5","E4"),
       pch=c(15,16),
       col=c("darkred","darkblue"),
       lty=1)
dev.off()

doc <- add_slide(doc)
doc <- ph_with(doc, value = "Latin America: E4 & E5 Average Scores", location = ph_location_type(type = "title"))

doc <- ph_with(doc, value = external_img(mypath),
               location = ph_location(left = 2.89, top = 1, width = 7.24, height =5.75))

Asia_group <- region_groups[region_groups$Region=="Asia" & region_groups$Edition != "E1" & region_groups$Edition != "E2" & region_groups$Edition != "E3",]
Asia_sum <- Asia_group %>%
  group_by(Edition) %>%
  summarise(across(where(is.numeric), mean, na.rm = TRUE))
rownames(Asia_sum) <- Asia_sum$Edition
Asia_sum <- select(Asia_sum, -Edition)
Asia_sum <- rbind(rep(100, 12), rep(0, 12), Asia_sum)
mypath <- paste("Y5_Radar_Charts/Asia 3i Y5 Radar Chart.png")
png(file = mypath)
radarchart(Asia_sum, # title = Y4_Country_List[i],
           seg = 10,  # Number of axis segments
           pcol = colors_line,
           pfcol = colors_fill,
           plwd = 2,
           plty = 1)
legend(1,1,
       legend=c("E5","E4"),
       pch=c(15,16),
       col=c("darkred","darkblue"),
       lty=1)
dev.off()

doc <- add_slide(doc)
doc <- ph_with(doc, value = "Asia: E4 & E5 Average Scores", location = ph_location_type(type = "title"))

doc <- ph_with(doc, value = external_img(mypath),
               location = ph_location(left = 2.89, top = 1, width = 7.24, height =5.75))

Europe_group <- region_groups[region_groups$Region=="Europe" & region_groups$Edition != "E1" & region_groups$Edition != "E2" & region_groups$Edition != "E3",]
Europe_sum <- Europe_group %>%
  group_by(Edition) %>%
  summarise(across(where(is.numeric), mean, na.rm = TRUE))
rownames(Europe_sum) <- Europe_sum$Edition
Europe_sum <- select(Europe_sum, -Edition)
Europe_sum <- rbind(rep(100, 12), rep(0, 12), Europe_sum)
mypath <- paste("Y5_Radar_Charts/Asia 3i Y5 Radar Chart.png")
png(file = mypath)
radarchart(Europe_sum, # title = Y4_Country_List[i],
           seg = 10,  # Number of axis segments
           pcol = colors_line,
           pfcol = colors_fill,
           plwd = 2,
           plty = 1)
legend(1,1,
       legend=c("E5","E4"),
       pch=c(15,16),
       col=c("darkred","darkblue"),
       lty=1)
dev.off()

doc <- add_slide(doc)
doc <- ph_with(doc, value = "Europe: E4 & E5 Average Scores", location = ph_location_type(type = "title"))

doc <- ph_with(doc, value = external_img(mypath),
               location = ph_location(left = 2.89, top = 1, width = 7.24, height =5.75))

MENA_group <- region_groups[region_groups$Region=="Middle East and North Africa" & region_groups$Edition != "E1" & region_groups$Edition != "E2" & region_groups$Edition != "E3",]
MENA_sum <- MENA_group %>%
  group_by(Edition) %>%
  summarise(across(where(is.numeric), mean, na.rm = TRUE))
rownames(MENA_sum) <- MENA_sum$Edition
MENA_sum <- select(MENA_sum, -Edition)
MENA_sum <- rbind(rep(100, 12), rep(0, 12), MENA_sum)
mypath <- paste("Y5_Radar_Charts/MENA 3i Y5 Radar Chart.png")
png(file = mypath)
radarchart(MENA_sum, # title = Y4_Country_List[i],
           seg = 10,  # Number of axis segments
           pcol = colors_line,
           pfcol = colors_fill,
           plwd = 2,
           plty = 1)
legend(1,1,
       legend=c("E5","E4"),
       pch=c(15,16),
       col=c("darkred","darkblue"),
       lty=1)
dev.off()

doc <- add_slide(doc)
doc <- ph_with(doc, value = "MENA: E4 & E5 Average Scores", location = ph_location_type(type = "title"))

doc <- ph_with(doc, value = external_img(mypath),
               location = ph_location(left = 2.89, top = 1, width = 7.24, height =5.75))

NA_group <- region_groups[region_groups$Region=="North America" & region_groups$Edition != "E1" & region_groups$Edition != "E2" & region_groups$Edition != "E3",]
NA_sum <- NA_group %>%
  group_by(Edition) %>%
  summarise(across(where(is.numeric), mean, na.rm = TRUE))
rownames(NA_sum) <- NA_sum$Edition
NA_sum <- select(NA_sum, -Edition)
NA_sum <- rbind(rep(100, 12), rep(0, 12), NA_sum)
mypath <- paste("Y5_Radar_Charts/NA 3i Y5 Radar Chart.png")
png(file = mypath)
radarchart(NA_sum, # title = Y4_Country_List[i],
           seg = 10,  # Number of axis segments
           pcol = colors_line,
           pfcol = colors_fill,
           plwd = 2,
           plty = 1)
legend(1,1,
       legend=c("E5","E4"),
       pch=c(15,16),
       col=c("darkred","darkblue"),
       lty=1)
dev.off()

doc <- add_slide(doc)
doc <- ph_with(doc, value = "North America: E4 & E5 Average Scores", location = ph_location_type(type = "title"))

doc <- ph_with(doc, value = external_img(mypath),
               location = ph_location(left = 2.89, top = 1, width = 7.24, height =5.75))

SSA_group <- region_groups[region_groups$Region=="Sub-Saharan Africa" & region_groups$Edition != "E1" & region_groups$Edition != "E2" & region_groups$Edition != "E3",]
SSA_sum <- SSA_group %>%
  group_by(Edition) %>%
  summarise(across(where(is.numeric), mean, SSA.rm = TRUE))
rownames(SSA_sum) <- SSA_sum$Edition
SSA_sum <- select(SSA_sum, -Edition)
SSA_sum <- rbind(rep(100, 12), rep(0, 12), SSA_sum)
mypath <- paste("Y5_Radar_Charts/SSA 3i Y5 Radar Chart.png")
png(file = mypath)
radarchart(SSA_sum, # title = Y4_Country_List[i],
           seg = 10,  # Number of axis segments
           pcol = colors_line,
           pfcol = colors_fill,
           plwd = 2,
           plty = 1)
legend(1,1,
       legend=c("E5","E4"),
       pch=c(15,16),
       col=c("darkred","darkblue"),
       lty=1)
dev.off()

doc <- add_slide(doc)
doc <- ph_with(doc, value = "SSA: E4 & E5 Average Scores", location = ph_location_type(type = "title"))

doc <- ph_with(doc, value = external_img(mypath),
               location = ph_location(left = 2.89, top = 1, width = 7.24, height =5.75))







country_region_map <- select(country_groups, -Income_Group)

Scores_Spider_Data <- select(all_Scores, c(1, 2, 5, 11, 19, 28,32,37,42,46,56, 61,	68))
Scores_Spider_Data <- left_join(Scores_Spider_Data, country_region_map, by = "Country")
Scores_Spider_Data <- Scores_Spider_Data[order(Scores_Spider_Data$Region, Scores_Spider_Data$Edition), ]
Scores_Spider_Data <- select(Scores_Spider_Data, -Region)


colnames(Scores_Spider_Data) <- c("Country", "Edition", "Usage", "Quality", "Infrastructure      ", "Electricity   ", "Price", "Competitive\nEnvironment", "Local\nContent", "  Relevant\nContent", "Literacy", "Trust &\nSafety", "Policy")


data_for_radar <- Scores_Spider_Data
data_for_radar_validate <- data_for_radar





Country_List <- unique(Scores_Spider_Data$Country)
data_for_radar_copy <- data_for_radar



Edition_order <- c("E5", "E4", "E3", "E2", "E1")

data_for_radar$Edition <- factor(data_for_radar$Edition, ordered = TRUE, 
                                levels = Edition_order)
data_for_radar<- data_for_radar[order(data_for_radar$Edition), ]

for (i in seq_along(Country_List)) {

data_for_radar_copy <- data_for_radar
data_for_radar_copy <- data_for_radar_copy %>%
  filter(Country == Country_List[i] & (Edition=="E5" | Edition=="E4"))
# rownames(data_for_radar_copy) <- data_for_radar_copy$Edition
data_for_radar_copy <- select(data_for_radar_copy, -Country, -Edition)

# data_for_radar_copy <- rbind(rep(max(data_for_radar_copy2), 12), rep(0, 12), data_for_radar_copy)
data_for_radar_copy <- rbind(rep(100, 12), rep(0, 12), data_for_radar_copy)
mypath <- paste("Y5_Radar_Charts/",Country_List[i], "3i Y5 Radar Chart", ".png")
png(file = mypath)
radarchart(data_for_radar_copy, # title = Y4_Country_List[i],
           seg = 10,  # Number of axis segments
           pcol = colors_line,
           pfcol = colors_fill,
           plwd = 2,
           plty = 1
           )
legend(1,-0.9,
       legend=c("E5","E4"),
       pch=c(15,16),
       col=c("darkred","darkblue"),
       lty=1)
dev.off()

# .rs.restartR() # run this from command line if error in png(file=mypath) : too many open devices.
doc <- add_slide(doc)
doc <- ph_with(doc, value = Country_List[i], location = ph_location_type(type = "title"))

doc <- ph_with(doc, value = external_img(mypath),
               location = ph_location(left = 2.89, top = 1, width = 7.24, height =5.75))

print(Country_List[i])
}

print(doc, target = "Y5_Radar_Charts/3i_Y5_radar_charts.pptx")

# print(countries_Asia)
