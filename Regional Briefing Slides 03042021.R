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
library(data.table)
library(ggplot2)
library(ggrepel)
library(plotly)
library(htmlwidgets)
library(webshot)
library(GGally)
library(orca)
library(ggradar)
suppressPackageStartupMessages(library(dplyr))
library(scales)
library(knitr)
library(plotly)
library(rworldmap)
library(maps)
library(ggmap)
library(mapdata)
library(writexl)
library(RColorBrewer)
library("viridis")
library(ggiraph)


lowcolor <- "yellow"
highcolor <- "darkred"

options(digits=1)

left  <- lay_new(matrix(1:2),1,c(1,11))
right <- lay_new(matrix(1:4), heights = c(1,5, 1, 5))
lay3   <- lay_bind_col(left,right, widths=c(15,15))
bottom <- lay_new(matrix(1:6,nc=2),widths = c(15,15), heights = c(1,5,1))
body <- lay_bind_row(lay3, bottom, heights = c(4,2))
titleLay <- lay_new(matrix(1:3,nc=3), widths = c(1,5,1), heights = 1)
cb_layout <- lay_bind_row(titleLay, body, heights = c(1,10))
# lay_show(cb_layout)

## create officer layout
cb_offLayout <- phl_layout(cb_layout, slideWidth = 13.33, slideHeight = 7.5, 
                           margins = c(0.25, 0.5, 0.25, 0.5),
                           innerMargins = rep(0.15,4))


left  <- lay_new(matrix(1:2),1,c(1,11))
right <- lay_new(matrix(1:2), heights = c(1,11))
lay3   <- lay_bind_col(left,right, widths=c(15,15))
bottom <- lay_new(matrix(1:6,nc=2),widths = c(15,15), heights = c(1,11,1))
body <- lay_bind_row(lay3, bottom, heights = c(4,2))
titleLay <- lay_new(matrix(1:3,nc=3), widths = c(1,5,1), heights = 1)
cb_layout <- lay_bind_row(titleLay, body, heights = c(1,10))
# lay_show(cb_layout)

## create officer layout
cb_offLayout2 <- phl_layout(cb_layout, slideWidth = 13.33, slideHeight = 7.5, 
                           margins = c(0.25, 0.5, 0.25, 0.5),
                           innerMargins = rep(0.15,4))


left  <- lay_new(matrix(1:2),1,c(1,11))
right <- lay_new(matrix(1:2), heights = c(1,11))
lay3   <- lay_bind_col(left,right, widths=c(15,15))
bottom <- lay_new(matrix(1:4,nc=2),widths = c(15,15), heights = c(1,11))
body <- lay_bind_row(lay3, bottom, heights = c(2,2.4))
titleLay <- lay_new(matrix(1:3,nc=3), widths = c(1,5,1), heights = 1)
cb_layout <- lay_bind_row(titleLay, body, heights = c(1,10))
# lay_show(cb_layout)

## create officer layout
cb_offLayout3 <- phl_layout(cb_layout, slideWidth = 13.33, slideHeight = 7.5, 
                            margins = c(0.25, 0.5, 0.25, 0.5),
                            innerMargins = rep(0.15,4))


setwd("/Users/sabumathai/3i Y5 - Revised Regional Decks/")
# getwd()

Model_exp <- read_excel("III_Fifth_Edition_Dataset.xlsx")

ggdata_wide <- select(Model_exp, 1:2, 6:7, 82:85)
ggdata_wide <- ggdata_wide %>%
  pivot_wider(names_from = Edition, values_from = c("1.1.4) Gender gap in internet access",	"1.1.5) Gender gap in mobile phone access", "BG20) Male internet users",	"BG21) Female internet users",	"BG22) Male mobile phone subscribers",	"BG23) Female mobile phone subscribers"))


ookla_quality_wide <- select(Model_exp, 1:2, 8:14)

ookla_quality_wide <- ookla_quality_wide %>%
  pivot_wider(names_from = Edition, values_from = c("1.2.1) Average fixed broadband upload speed", "1.2.2) Average fixed broadband download speed", "1.2.3) Average fixed broadband latency", "1.2.4) Average mobile upload speed",	"1.2.5) Average mobile download speed",	"1.2.6) Average mobile latency",	"1.2.7) Bandwidth capacity"))

ookla_quality_wide$FBUP_LTM <- (ookla_quality_wide$`1.2.1) Average fixed broadband upload speed_E5` - ookla_quality_wide$`1.2.1) Average fixed broadband upload speed_E4`)/ookla_quality_wide$`1.2.1) Average fixed broadband upload speed_E4`

ookla_quality_wide$FBUP_CAGR <- (ookla_quality_wide$`1.2.1) Average fixed broadband upload speed_E5`/ookla_quality_wide$`1.2.1) Average fixed broadband upload speed_E1`)^(1/4)-1

ookla_quality_wide$FBDL_LTM <- (ookla_quality_wide$`1.2.2) Average fixed broadband download speed_E5` - ookla_quality_wide$`1.2.2) Average fixed broadband download speed_E4`)/ookla_quality_wide$`1.2.2) Average fixed broadband download speed_E4`

ookla_quality_wide$FBDL_CAGR <- (ookla_quality_wide$`1.2.2) Average fixed broadband download speed_E5`/ookla_quality_wide$`1.2.2) Average fixed broadband download speed_E1`)^(1/4)-1


ookla_quality_wide$FBLAT_LTM <- (ookla_quality_wide$`1.2.3) Average fixed broadband latency_E5` - ookla_quality_wide$`1.2.3) Average fixed broadband latency_E4`)/ookla_quality_wide$`1.2.3) Average fixed broadband latency_E4`

ookla_quality_wide$FBLAT_CAGR <- (ookla_quality_wide$`1.2.3) Average fixed broadband latency_E5`/ookla_quality_wide$`1.2.3) Average fixed broadband latency_E1`)^(1/4)-1

ookla_quality_wide$MOBUP_LTM <- (ookla_quality_wide$`1.2.4) Average mobile upload speed_E5` - ookla_quality_wide$`1.2.4) Average mobile upload speed_E4`)/ookla_quality_wide$`1.2.4) Average mobile upload speed_E4`

ookla_quality_wide$MOBUP_CAGR <- (ookla_quality_wide$`1.2.4) Average mobile upload speed_E5`/ookla_quality_wide$`1.2.4) Average mobile upload speed_E1`)^(1/4)-1

ookla_quality_wide$MOBDL_LTM <- (ookla_quality_wide$`1.2.5) Average mobile download speed_E5` - ookla_quality_wide$`1.2.5) Average mobile download speed_E4`)/ookla_quality_wide$`1.2.5) Average mobile download speed_E4`

ookla_quality_wide$MOBDL_CAGR <- (ookla_quality_wide$`1.2.5) Average mobile download speed_E5`/ookla_quality_wide$`1.2.5) Average mobile download speed_E1`)^(1/4)-1

ookla_quality_wide$MOBLAT_LTM <- (ookla_quality_wide$`1.2.6) Average mobile latency_E5` - ookla_quality_wide$`1.2.6) Average mobile latency_E4`)/ookla_quality_wide$`1.2.6) Average mobile latency_E4`

ookla_quality_wide$MOBLAT_CAGR <- (ookla_quality_wide$`1.2.6) Average mobile latency_E5`/ookla_quality_wide$`1.2.6) Average mobile latency_E1`)^(1/4)-1

ookla_quality_wide <- select(ookla_quality_wide, 1, 37:48)

gg_wide <- select(Model_exp, Country, Edition, `1.1.4) Gender gap in internet access`)
gg_wide <- gg_wide %>%
  pivot_wider(names_from = Edition, values_from = c("1.1.4) Gender gap in internet access"))
gg_wide$PP_DIFF <- gg_wide$E5 - gg_wide$E4
gg_wide <- select(gg_wide, Country, PP_DIFF)


E4E5_QUALITY_WIDE <- select(Model_exp[Model_exp$Edition=="E5" | Model_exp$Edition=="E4", ], 1:2, 8:14)

E4E5_QUALITY_WIDE <- E4E5_QUALITY_WIDE %>%
  pivot_wider(names_from = Edition, values_from = c("1.2.1) Average fixed broadband upload speed", "1.2.2) Average fixed broadband download speed", "1.2.3) Average fixed broadband latency", "1.2.4) Average mobile upload speed",	"1.2.5) Average mobile download speed",	"1.2.6) Average mobile latency",	"1.2.7) Bandwidth capacity"))

CB_Y4 <- read_excel("Y5 - Country Briefing.xlsx")

Scores_Y4 <- read_excel("Y5_Scores.xlsx")

CB_Y4 <- left_join(CB_Y4, Scores_Y4, by = "Country")
CB_Y4 <- left_join(CB_Y4, E4E5_QUALITY_WIDE, by = "Country")
CB_Y4 <- left_join(CB_Y4, ggdata_wide, by = "Country")


CB_Y4$`Overall rank_global` <- CB_Y4$OVERALL_rank
CB_Y4$`Availability_global rank` <-CB_Y4$AVAILABILITY_rank
CB_Y4$`Affordability_overall rank` <- CB_Y4$AFFORDABILITY_rank
CB_Y4$`Relevance_overall rank`<- CB_Y4$RELEVANCE_rank
CB_Y4$`Readiness_overall rank`<- CB_Y4$READINESS_rank


CB_Y4$`Overall rank_regional` <- CB_Y4$OVERALL_regional_rank
CB_Y4$`Availabilty_regional rank` <- CB_Y4$AVAILABILITY_regional_rank
CB_Y4$`Affordability_regional rank` <- CB_Y4$AFFORDABILITY_regional_rank
CB_Y4$`Relevance_regional rank` <- CB_Y4$RELEVANCE_regional_rank
CB_Y4$`Readiness_regional rank` <- CB_Y4$READINESS_regional_rank


Y4_data_raw <- read_excel("3i Y5 - dataset.xlsx")
Y4_temp_dr <- select(Y4_data_raw, Country, fmai)
CB_Y4 <- left_join(CB_Y4, Y4_temp_dr, by = "Country")
CB_Y4 <- left_join(CB_Y4, gg_wide, by = "Country")
CB_Y4 <- left_join(CB_Y4, ookla_quality_wide, by = "Country")




doc <- read_pptx(path = "3i_Y5_report_template.pptx")





for(counter in c("Asia", "Europe (EUR)", "Latin America (LATAM)", "Middle East and North Africa (MENA)", "North America (NA)", "Sub-Saharan Africa (SSA)")){

Y4_data <- Y4_data_raw


if (counter=="Asia"){
  Y4_data <- Y4_data[Y4_data$Region=="Asia",]
} else if (counter=="Europe (EUR)"){
  Y4_data <- Y4_data[Y4_data$Region=="Europe",]
} else if (counter=="Latin America (LATAM)"){
  Y4_data <- Y4_data[Y4_data$Region=="Latam",]
} else if (counter=="Middle East and North Africa (MENA)"){
  Y4_data <- Y4_data[Y4_data$Region=="MENA",]
} else if (counter=="North America (NA)"){
  Y4_data <- Y4_data[Y4_data$Region=="North America",]
} else if (counter=="Sub-Saharan Africa (SSA)"){
  Y4_data <- Y4_data[Y4_data$Region=="SSA",]
} 

Y4_data$colorgg <- 0
Y4_data$sizegg <- 0

Y4_data$gglabelfont <- "bold"

Y4_data$colorgg <- Y4_data$fmai >= 10

for(ggcounter in seq_along(Y4_data$fmai)){
  if (Y4_data$fmai[ggcounter] >= 10){
    Y4_data$sizegg[ggcounter] <- 2
    Y4_data$colorgg[ggcounter] <- "#8B0000"
  }else if (Y4_data$fmai[ggcounter] > 0){
    Y4_data$sizegg[ggcounter] <- 1
    Y4_data$colorgg[ggcounter] <- "#696969"
  } else {
    Y4_data$sizegg[ggcounter] <- 1
    Y4_data$colorgg[ggcounter] <- "#CD0000"
  }
}


gg_slope_max <- max(Y4_data$`BG20) Male internet users / % of male population`)
if (max(Y4_data$`BG21) Female internet users / % of female population`)> max(Y4_data$`BG20) Male internet users / % of male population`)){
gg_slope_max <- max(Y4_data$`BG21) Female internet users / % of female population`)
}
gg_slope_min <- min(Y4_data$`BG20) Male internet users / % of male population`)
if (min(Y4_data$`BG21) Female internet users / % of female population`)<min(Y4_data$`BG20) Male internet users / % of male population`)){
  gg_slope_min <- min(Y4_data$`BG21) Female internet users / % of female population`)
}

#gender slope visualization
gg_slope <-   ggplot() +
  # set the colors
  scale_color_manual(values = c("#8B0000", "#CD0000", "#696969"), guide = "none")  +
  # add a line segment that goes from men to women for each country
  geom_segment(aes(x = 1, xend = 2, 
                   y = Y4_data$`BG21) Female internet users / % of female population`, 
                   yend = Y4_data$`BG20) Male internet users / % of male population`,
                   group = Y4_data$Country,
                   col = Y4_data$fmai), 
               size = Y4_data$sizegg) +
  # remove all axis stuff
  theme_classic() +
  theme(axis.line = element_blank(),
        axis.text = element_blank(),
        axis.title = element_blank(),
        axis.ticks = element_blank()) +
  # add vertical lines that act as axis for men
  geom_segment(mapping=aes(
               x = 1, xend = 1, 
               y = gg_slope_min-1*(gg_slope_max-gg_slope_min)/gg_slope_max,
               yend = gg_slope_max+1*(gg_slope_max-gg_slope_min)/gg_slope_max), 
               col = "grey70",
               size = 1) +
  # add vertical lines that act as axis for women
  geom_segment(mapping=aes(x = 2, xend = 2, 
              y = gg_slope_min-1*(gg_slope_max-gg_slope_min)/gg_slope_max,
              yend = gg_slope_max+1*(gg_slope_max-gg_slope_min)/gg_slope_max), 
              col = "grey70",
              size = 1) +
  geom_segment(mapping=aes(
    x = .98, xend = .98, 
    y = gg_slope_min-1*(gg_slope_max-gg_slope_min)/gg_slope_max,
    yend = gg_slope_max+1*(gg_slope_max-gg_slope_min)/gg_slope_max), 
    col = "grey70",
    size = .5) +
  # add vertical lines that act as axis for women
  geom_segment(mapping=aes(x = 2.02, xend = 2.02, 
                           y = gg_slope_min-1*(gg_slope_max-gg_slope_min)/gg_slope_max,
                           yend = gg_slope_max+1*(gg_slope_max-gg_slope_min)/gg_slope_max), 
               col = "grey70",
               size = .5) +
  # add the words "men" and "women" above their axes
  # geom_text(aes(x = x, y = y, label = label), fontface = "bold",
  #           data = data.frame(x = 1:2, 
  #                             y = gg_slope_max+8*(gg_slope_max-gg_slope_min)/gg_slope_max, label = c("Female\nInternet\nUsers", "Male\nInternet\nUsers")),
  #           col = "grey30",
  #           size = 5) +
  # set the limits of the x-axis so that the labels are not cut off
  scale_x_continuous(breaks = 1:2, limits = c(0.2, 2.8))+#, expand = expansion(mult = 0.3)) +
  geom_text_repel(aes(x = 0.98,
              y = Y4_data$`BG21) Female internet users / % of female population`,
              label = paste0(Y4_data$Country, " (",
                             round(Y4_data$`BG21) Female internet users / % of female population`, 1), "%)")),
              nudge_x = -0.15, direction = "y", col = "grey30", hjust = "right", fontface = "bold", size=5, force_pull = 5, segment.alpha = .5, segment.linetype=2, segment.curvature = 1e-20) +
  # add the success rate next to each point on the women axis
  geom_text_repel(aes(x = 2.02,
                y = Y4_data$`BG20) Male internet users / % of male population`,
                label = paste0(Y4_data$Country, " (", round(Y4_data$`BG20) Male internet users / % of male population`, 1), "%)")),
                nudge_x = 0.15, direction = "y",col = "grey30", hjust = "left", fontface = "bold",size=5, force_pull = 5, segment.alpha = .5, segment.linetype=2, segment.curvature = -1e-20)+scale_color_binned(limits = c(-14, 71), breaks = c(0,10, 20, 30, 40, 50, 60, 70), low = lowcolor, high = highcolor)+theme(legend.position = "none")#+
# add the white outline for the points at each rate for men
  # geom_point(aes(x = 1, 
  #              y = Y4_data$`BG21) Female internet users / % of female population`), size = 5,
  #          col = "white") +
  # # add the white outline for the points at each rate for women
  # geom_point(aes(x = 2, 
  #                y = Y4_data$`BG20) Male internet users / % of male population`), size = 5,
  #            col = "white") +
  # 
  # # add the actual points at each rate for men
  # geom_point(aes(x = 1, 
  #                y = Y4_data$`BG21) Female internet users / % of female population`), size = 4,
  #            col = "grey60") +
  # # add the actual points at each rate for men
  # geom_point(aes(x = 2, 
  #                y = Y4_data$`BG20) Male internet users / % of male population`), size = 4,
  #            col = "grey60") 

gg_slope


doc <- add_slide(doc)
doc <- ph_with(doc, value = gg_slope,
               location = ph_location(left = 0.5, top = .3, height=17, width=10))


print(counter)
}



Y4_data <- Y4_data_raw
affplot <- ggplot(Y4_data,
                  aes(x=inth, y = `2.1.2`)) +
  geom_point(position="jitter",
             mapping = aes(color=`Income Group`))+
  geom_segment(x=-20, xend=110, y=2, yend=2, linetype = "dotted")+
  xlab("Household Connectivity (% Population)")+ylab("Price of 1Gb Prepaid data (% of monthly GNI per capita)")+labs(color="Country Income Groups")+
  theme_classic() +
  geom_text_repel(aes(label=Country), size = 3, segment.color="grey", max.overlaps = Inf)+
  theme(legend.position="bottom", legend.title = element_text(size=10, face="bold"), legend.text=element_text(size=10, face="plain"))+
  scale_x_continuous(limits = c(-6, 101))+
  geom_text(x=-10, y=2.05, label="Greater than 2% share\nof GNI per capita is\nconsidered unaffordable", col = "grey40", hjust = "left", vjust = "bottom",  fontface = "italic", size = 3)+
  scale_y_continuous(breaks=c(0,2,5,10,15,20), limits = c(0,20))+
  theme(panel.background = element_rect(fill = "#EFEDE5", colour = "#EFEDE5"))
affplot

affplot_reg <- ggplot(Y4_data,
                  aes(x=inth, y = `2.1.2`)) +
  geom_point(position="jitter",
             mapping = aes(color=Region))+
  geom_segment(x=-20, xend=110, y=2, yend=2, linetype = "dotted")+
  xlab("Household Connectivity (% Population)")+ylab("Price of 1Gb Prepaid data (% of monthly GNI per capita)")+labs(color="Regions")+
  theme_classic() +
  geom_text_repel(aes(label=Country), size = 2, segment.color="grey", max.overlaps = Inf)+
  theme(legend.position="bottom", legend.title = element_text(size=10, face="bold"), legend.text=element_text(size=10, face="plain"))+
  scale_x_continuous(limits = c(-6, 101))+
  geom_text(x=-10, y=2.05, label="Greater than 2% share\nof GNI per capita is\nconsidered unaffordable", col = "grey40", hjust = "left", vjust = "bottom",  fontface = "italic", size = 3)+
  scale_y_continuous(breaks=c(0,2,5,10,15,20), limits = c(0,20))+
  theme(panel.background = element_rect(fill = "#EFEDE5", colour = "#EFEDE5"))

affplot_reg




Y4_data_raw$Country_label <- Y4_data_raw$Country
Y4_data_raw$fmai_color<- cut(Y4_data_raw$fmai, c(-Inf,-5,0,5, 10, 20, 40, Inf), c("#31A354", "#A1D99B", "#FFEDA0",  "#FEB24C", "#FD8D3C", "#FC4E2A", "#B10026"))



SSA_countries <- c("Côte d'Ivoire", "Ethiopia", "Ghana", "Kenya", "Liberia", "Madagascar", "Malawi", "Mozambique", "Namibia", "Nigeria", "Rwanda", "Senegal", "South Africa", "Sudan", "Tanzania", "Uganda", "Zambia", "Angola", "Benin", "Democratic Republic of the Congo", "Republic of Congo", "Guinea", "Mali", "Niger", "Sierra Leone", "Burundi", "Gabon", "Zimbabwe", "Botswana", "Burkina Faso", "Cameroon", "Central African Republic", "Chad", "Eritrea", "Gambia", "Guinea-Bissau", "Lesotho", "Mauritania", "Niger", "Somalia", "Togo", "Zaire", "Ivory Coast", "Swaziland", "South Sudan")
ssa.maps <- map_data("world", region = SSA_countries)
ssa.maps[ssa.maps == "Ivory Coast"] <- "Côte d'Ivoire"
ssa.maps[ssa.maps == "Democratic Republic of the Congo"] <- "Congo (DRC)"
ssa.maps$Country <- ssa.maps$region
ssa.maps <- left_join(ssa.maps, Y4_data_raw, by="Country")
for (i in seq_along(ssa.maps$Country_label)) {
  if(is.na(ssa.maps$Country_label[i])) {
    ssa.maps$Country_label[i] <- ""
    }
}
region.lab.data <- ssa.maps %>%
  group_by(Country_label) %>%
  summarise(long = mean(long), lat = mean(lat))
Y4_map <- select(Y4_data_raw, Country, fmai)
Y4_map$Country_label <- Y4_map$Country
Y4_map <- select(Y4_map, -Country)
region.lab.data <- left_join(region.lab.data, Y4_map, by = "Country_label")
region.lab.data <- na.omit(region.lab.data)
SSAplot <- ggplot(ssa.maps, aes(x = long, y = lat)) +
  geom_polygon(aes( group = group, fill = fmai), size = 0.25, color = "grey40")+
  geom_label_repel(aes(label = paste(Country_label,fmai)), data = region.lab.data, fontface = "bold", size = 3, force_pull = 5, max.overlaps = Inf)+
  theme_void()+
  theme(legend.position = "none")+scale_fill_binned(limits = c(-14, 71), breaks = c(0,10, 20, 30, 40, 50, 60, 70), low = lowcolor, high = highcolor)

Asia_countries <- c("Hong Kong", "Australia", "Bangladesh", "Cambodia", "China", "India", "Indonesia", "Iran", "Japan", "Kazakhstan", "Malaysia", "Mongolia", "Myanmar", "Pakistan", "Philippines", "Singapore", "South Korea", "Sri Lanka", "Taiwan", "Thailand", "Vietnam", "Nepal", "Azerbaijan", "Laos", "New Zealand", "Papua New Guinea", "Bhutan", "Afghanistan", "Tajikistan", "Kazakhstan", "Uzbekistan", "Turkmenistan", "Kyrgyzstan" )
a.maps <- map_data("world2", region = Asia_countries)
a.maps$Country <- a.maps$region
a.maps <- left_join(a.maps, Y4_data_raw, by="Country")
region.lab.data <- a.maps %>%
  group_by(Country_label) %>%
  summarise(long = mean(long), lat = mean(lat))
HK <- data.frame("Hong Kong", 118, 23)
names(HK) <- c("Country_label", "long", "lat")
region.lab.data<-rbind(region.lab.data, HK)
region.lab.data <- left_join(region.lab.data, Y4_map, by = "Country_label")
region.lab.data <- na.omit(region.lab.data)
Asiaplot <- ggplot(a.maps, aes(x = long, y = lat)) +
  geom_polygon(aes( group = group, fill = fmai), size=0.25, color="grey40")+
  # geom_label_repel(aes(label = paste(Country_label, fmai)), data = region.lab.data,  size = 3, fontface = "bold")+
  theme_void()+
  theme(legend.position = "none")+scale_fill_binned(limits = c(-14, 71), breaks = c(0,10, 20, 30, 40, 50, 60, 70), low = lowcolor, high = highcolor)+theme(
    plot.background = element_rect(fill = "transparent",colour = NA)
  )

LA_countries <- c("Dominican Republic", "Trinidad", "Argentina", "Brazil", "Chile", "Colombia", "El Salvador", "Guatemala", "Jamaica", "Mexico", "Peru", "Venezuela", "Costa Rica", "Ecuador", "Panama", "Uruguay", "Cuba", "Honduras", "Nicaragua", "Paraguay", "Bolivia", "Guyana", "Suriname", "French Guiana")
la.maps <- map_data("world2", region = LA_countries)
all.maps <- map_data("world2")
u_all.maps <- unique(all.maps[c("region")])
la.maps[la.maps == "Trinidad"] <- "Trinidad & Tobago"
la.maps[la.maps == "Dominican Republic"] <- "Dominican Republic "
la.maps$Country <- la.maps$region
la.maps <- left_join(la.maps, Y4_data_raw, by="Country")
region.lab.data <- la.maps %>%
  group_by(Country_label) %>%
  summarise(long = mean(long), lat = mean(lat))
region.lab.data <- left_join(region.lab.data, Y4_map, by = "Country_label")
region.lab.data <- na.omit(region.lab.data)
LAplot <- ggplot(la.maps, aes(x = long, y = lat)) +
  geom_polygon(aes(group = group, fill = fmai), size=0.25, color="grey40")+
  geom_label_repel(aes(label = paste(Country_label, fmai)), data = region.lab.data,  size = 3, hjust = 0.5)+
  theme_void()+
  theme(legend.position = "none")+scale_fill_binned(limits = c(-14, 71), breaks = c(0,10, 20, 30, 40, 50, 60, 70), low = lowcolor, high = highcolor)

EU_countries <- c("Ireland", "UK", "Austria", "Belgium", "Bulgaria", "Denmark", "Estonia", "France", "Germany", "Greece", "Hungary", "Italy", "Netherlands", "Poland", "Portugal", "Romania", "Russia", "Spain", "Sweden", "Turkey", "Czech Republic", "Finland", "Switzerland", "Ukraine", "Croatia", "Latvia", "Lithuania", "Slovakia","Belarus", "Moldova", "Slovenia")

eu.maps <- map_data("world", region = EU_countries)
eu.maps[eu.maps == "UK"] <- "United Kingdom"
eu.maps$Country <- eu.maps$region
eu.maps <- left_join(eu.maps, Y4_data_raw, by="Country")
region.lab.data <- eu.maps %>%
  group_by(Country_label) %>%
  summarise(long = mean(long), lat = mean(lat))
region.lab.data <- left_join(region.lab.data, Y4_map, by = "Country_label")
region.lab.data <- na.omit(region.lab.data)
EUplot <- ggplot(eu.maps, aes(x = long, y = lat)) +
  geom_polygon(aes(group = group, fill = fmai), size=.25, color="grey40")+
  geom_label_repel(aes(label = paste(Country_label, fmai)), data = region.lab.data,  size = 3, hjust = 0.5)+
  theme_void()+
  theme(legend.position = "none")+scale_fill_binned(limits = c(-14, 71), breaks = c(0,10, 20, 30, 40, 50, 60, 70), low = lowcolor, high = highcolor)

  


NA_countries <- c("Canada", "USA")

na.maps <- map_data("world2", region = NA_countries)
na.maps[na.maps == "USA"] <- "United States"
na.maps$Country <- na.maps$region
na.maps <- left_join(na.maps, Y4_data_raw, by="Country")
region.lab.data <- na.maps %>%
  group_by(Country_label) %>%
  summarise(long = mean(long), lat = mean(lat))
region.lab.data <- left_join(region.lab.data, Y4_map, by = "Country_label")
region.lab.data <- na.omit(region.lab.data)
NAplot <- ggplot(na.maps, aes(x = long, y = lat)) +
  geom_polygon(aes(group = group, fill = fmai), size=.75, color="grey40")+
  geom_label_repel(aes(label = paste(Country_label, fmai)), data = region.lab.data,  size = 3, hjust = 0.5)+
  theme_void()+
  theme(legend.position = "none")+scale_fill_binned(limits = c(-14, 71), breaks = c(0,10, 20, 30, 40, 50, 60, 70), low = lowcolor, high = highcolor)
NAplot



MENA_countries <- c("Algeria", "United Arab Emirates", "Egypt", "Kuwait", "Morocco", "Oman", "Qatar", "Saudi Arabia", "Israel", "Jordan", "Tunisia", "Bahrain", "Lebanon", "Libya")

mena.maps <- map_data("world", region = MENA_countries)
# mena.maps[mena.maps == "USA"] <- "United States"
mena.maps[mena.maps == "United Arab Emirates"] <- "UAE"
mena.maps$Country <- mena.maps$region
mena.maps <- left_join(mena.maps, Y4_data_raw, by="Country")
region.lab.data <- mena.maps %>%
  group_by(Country_label) %>%
  summarise(long = mean(long), lat = mean(lat))
region.lab.data <- left_join(region.lab.data, Y4_map, by = "Country_label")
region.lab.data <- na.omit(region.lab.data)
MENAplot <- ggplot(mena.maps, aes(x = long, y = lat)) +
  geom_polygon_interactive(aes(group = group, fill = fmai), size=.75, color="grey40")+
  geom_label_repel(aes(label = paste(Country_label, fmai)), data = region.lab.data,  size = 3, hjust = 0.5)+
  theme_void()+
  theme(legend.position = "right", legend.key.height = unit(2.0, "cm"), legend.title = element_text(size=10, face="bold"))+
  labs(fill="Gender\nGap")+scale_fill_binned(limits = c(-14, 71), breaks = c(0,10, 20, 30, 40, 50, 60, 70), low = lowcolor, high = highcolor)

MENAplot



doc <- add_slide(doc)
doc <- ph_with(doc, value = affplot,
               location = ph_location(left = 5, top = .5, height=7, width=12))
doc <- add_slide(doc)
doc <- ph_with(doc, value = SSAplot,
               location = ph_location(left = .5, top = .5, height=7.5))
doc <- add_slide(doc)
doc <- ph_with(doc, value = Asiaplot,
               location = ph_location(left = 0, top = .0, height=7.5))
doc <- add_slide(doc)
doc <- ph_with(doc, value = EUplot,
               location = ph_location(left = 0, top = .0, height=7.5))
doc <- add_slide(doc)
doc <- ph_with(doc, value = LAplot,
               location = ph_location(left = 0, top = .0, height=7.5))
doc <- add_slide(doc)
doc <- ph_with(doc, value = NAplot,
               location = ph_location(left = 0, top = .0, height=7.5))
doc <- add_slide(doc)
doc <- ph_with(doc, value = MENAplot,
               location = ph_location(left = 0, top = .0, height=7.5))


print(doc, target = "3i_Y5_Regional Front Matter.pptx")

# SSAplot
# Asiaplot
# EUplot
# LAplot
# NAplot
# MENAplot


# ###################
# MENA_countries <- c("Algeria", "United Arab Emirates", "Egypt", "Kuwait", "Morocco", "Oman", "Qatar", "Saudi Arabia", "Israel", "Jordan", "Tunisia", "Bahrain", "Lebanon", "Libya")
# 
# mena.maps <- map_data("world", region = MENA_countries)
# # mena.maps[mena.maps == "USA"] <- "United States"
# mena.maps[mena.maps == "United Arab Emirates"] <- "UAE"
# mena.maps$Country <- mena.maps$region
# mena.maps <- left_join(mena.maps, Y4_data_raw, by="Country")
# region.lab.data <- mena.maps %>%
#   group_by(Country_label) %>%
#   summarise(long = mean(long), lat = mean(lat))
# region.lab.data <- left_join(region.lab.data, Y4_map, by = "Country_label")
# region.lab.data <- na.omit(region.lab.data)
# MENAplot <- ggplot(mena.maps, aes(x = long, y = lat)) +
#   geom_polygon_interactive(aes(group = group, fill = fmai, tooltip = paste0(Country_label, "\n",fmai)), size=.75, color="grey40")+
#   # geom_label_repel(aes(label = paste(Country_label, fmai)), data = region.lab.data,  size = 3, hjust = 0.5)+
#   theme_void()+
#   # theme(legend.position = "right", legend.key.height = unit(2.0, "cm"), legend.title = element_text(size=10, face="bold"))+
#   labs(fill="Gender\nGap")+scale_fill_binned(limits = c(-14, 71), breaks = c(0,10, 20, 30, 40, 50, 60, 70), low = lowcolor, high = highcolor)
# 
# MENAplot
# 
# 
# # display ------
# x <- girafe(ggobj = MENAplot)
# if( interactive() ) print(x)

