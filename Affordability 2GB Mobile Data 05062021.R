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
  library(ggthemes)
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
  library(stringr)
  library(directlabels)
  library(ggbeeswarm)
  
  soptions(digits=1)
  

  
  
  setwd("/Users/sabumathai/3i Y5 - Revised Regional Decks/")
  # getwd()
  GNI2020 <- read_excel("2020 GNI.xlsx")
  prepaid2GB_prices <- read_excel("Additional Data EIU 2GB.xlsx")
  CB1 <- read_excel("Category 0309.xlsx")
  CB2 <- read_excel("Subcategory 0309.xlsx")
  CB3 <- read_excel("Indicator_S2.xlsx")
  CB4 <- read_excel("CBN.xlsx")
  CB5 <- read_excel("3i Y5 Reporter - Strengths and Weaknesses.xlsx")
  
  CB1 <- left_join(CB1, CB2, by = "Country")
  CB1 <- left_join(CB1, CB3, by = "Country")
  CB1 <- left_join(CB1, CB4, by = "Country")
  CB1 <- left_join(CB1, CB5, by = "Country")
  
  
  # write_xlsx(x = CB1, path = "3i Y5 CBN.xlsx", col_names = TRUE)
  
  Model_exp <- read_excel("III_Fifth_Edition_Dataset.xlsx")
  
  # CB_Y4 <- read_excel("Y5 - Country Briefing.xlsx")
  CB_Y4 <- CB1
  
  Scores_Y4 <- read_excel("Y5_Scores.xlsx")
  
  CB_Y4 <- left_join(CB_Y4, Scores_Y4, by = "Country")
  prepaid2GB_aff <- left_join(GNI2020, prepaid2GB_prices, by = "Country")
  prepaid2GB_aff <- select(prepaid2GB_aff, 1, 3, 5)
  prepaid2GB_aff$monthlyGNIpc <- (prepaid2GB_aff$`GNI per capita, 2020`/365)*30
  prepaid2GB_aff$aff2G <-prepaid2GB_aff$prepaid_price_2GB_USD_2020_Q4 / prepaid2GB_aff$monthlyGNIpc*100
  

  
  
  Y4_data_raw <- read_excel("3i Y5 - dataset.xlsx")
  Y4_temp_dr <- select(Y4_data_raw, Country, fmai)
  Y4_data_raw <- left_join(Y4_data_raw, prepaid2GB_aff, by = "Country")

  
  CB_Y4 <- left_join(CB_Y4, Y4_temp_dr, by = "Country")
  
  
  doc <- read_pptx(path = "3i_Y5_report_template.pptx")
  
  
  
  
  
  for(counter in seq_along(CB_Y4$Country)){
  
  Y4_data <- Y4_data_raw
  
  
  CB_Y4$Region_short <- CB_Y4$Region
  CB_Y4$Region_short[CB_Y4$Region_short=="Asia"] <- "ASIA"
  CB_Y4$Region_short[CB_Y4$Region_short=="Europe (EUR)"] <- "EUR"
  CB_Y4$Region_short[CB_Y4$Region_short=="Latin America (LATAM)"] <- "LATAM"
  CB_Y4$Region_short[CB_Y4$Region_short=="Middle East and North Africa (MENA)"] <- "MENA"
  CB_Y4$Region_short[CB_Y4$Region_short=="North America (NA)"] <- "NA"
  CB_Y4$Region_short[CB_Y4$Region_short=="Sub-Saharan Africa (SSA)"] <- "SSA"
  
  
  
  if (CB_Y4$Region[counter]=="Asia"){
    Y4_data <- Y4_data[Y4_data$Region=="Asia",]
  } else if (CB_Y4$Region[counter]=="Europe (EUR)"){
    Y4_data <- Y4_data[Y4_data$Region=="Europe",]
  } else if (CB_Y4$Region[counter]=="Latin America (LATAM)"){
    Y4_data <- Y4_data[Y4_data$Region=="Latam",]
  } else if (CB_Y4$Region[counter]=="Middle East and North Africa (MENA)"){
    Y4_data <- Y4_data[Y4_data$Region=="MENA",]
  } else if (CB_Y4$Region[counter]=="North America (NA)"){
    Y4_data <- Y4_data[Y4_data$Region=="North America",]
  } else if (CB_Y4$Region[counter]=="Sub-Saharan Africa (SSA)"){
    Y4_data <- Y4_data[Y4_data$Region=="SSA",]
  } 
  
  
  affplot <- ggplot(Y4_data, 
                    aes(x=inth, y = aff2G)) + 
    geom_point(position="jitter", 
               mapping = aes(color=`Income Group`))+
    geom_hline(yintercept=2, linetype = "dotted")+
    xlab("\nHousehold Connectivity (% Population)")+ylab("Price of 2Gb prepaid data (% monthly GNI per capita)\n")+labs(color="Country Income Groups", size="Gender Gap Ratio")+
    geom_text_repel(aes(label=Country), size = 3, segment.color="grey", max.overlaps = Inf)+ 
    scale_x_continuous(limits = c(-6, 101))+
    scale_y_continuous(breaks=c(0,2,5,10,15,20, 25, 30, 35))+theme_classic()+theme(legend.position="bottom", legend.title = element_text(size=9, face="bold"), legend.text=element_text(size=9, face="plain"))+
    theme(panel.background = element_rect(fill = "#D8E3EA", colour = "#D8E3EA"))
    
    # theme(axis.line.x = element_blank(), axis.ticks.x = element_blank(), axis.title.x = element_text(face="bold"), axis.title.y = element_text(face="bold"))
  
  
  
  
  
  
  Y4_data$hsetplotfill <- Y4_data$Country==CB_Y4$Country[counter]
  
  hsetplot <- ggplot() +
    scale_fill_manual(values = c("#9a9597", "#ffb237"))+
    geom_col(data=Y4_data, aes(fill=Y4_data$hsetplotfill, x=reorder(Y4_data$Country, -Y4_data$`2.1.1`), y=Y4_data$`2.1.1`))+coord_flip()+theme_classic() +theme(axis.title.x = element_text(face = "bold", size=7), axis.text.x = element_text(face="plain",size=6),axis.title.y = element_blank(), axis.text.y = element_text(size=6), panel.background = element_rect(fill = "#D8E3EA", colour = "#D8E3EA"))+
    theme(legend.position="none")+labs(y= "Handset Affordability")
  hsetplot
  

  
  Y4_data_counter <- 0
  for(tempcounter in seq_along(Y4_data$Country)){
    if(Y4_data$Country[tempcounter]==CB_Y4$Country[counter]){
      Y4_data_counter <- tempcounter
    }
  }
  
  
  Y4_data$aff2G <- round(Y4_data$aff2G, 1)
  fbafftb <- data.frame(fbaffdf = c(paste0(Y4_data$aff2G[Y4_data_counter],"%"), paste0("Cost of 2GB Mobile Data (Prepaid) as a share of GNI per capita in ", Y4_data$Country[Y4_data_counter])))
  
  fbaffft <-flextable(fbafftb)
  fbaffft <- align(fbaffft, align = "center", part = "all")
  fbaffft <- color(fbaffft, part = "header", color = "#FFFFFF")
  fbaffft <- delete_part(fbaffft, part = "header")
  fbaffft <- fontsize(fbaffft, i = 1, part = "body", size = 44)
  fbaffft <- fontsize(fbaffft, i = 2, part = "body", size = 12)
  fbaffft <- bold(fbaffft, j = 1, part = "body")
  fbaffft <- border_remove(x = fbaffft)
  fbaffft <- width(fbaffft, j=1, width = 2.5)
  
  
  Y4_data$mobs <- round(Y4_data$mobs, 1)
  flbsubs <- data.frame(flbsubs = c(paste0("2 for 2"), paste0("2GB of mobile broadband data for 2% or less of monthly GNI per capita")))
  
  flbsubsft <-flextable(flbsubs)
  flbsubsft <- align(flbsubsft, align = "center", part = "all")
  flbsubsft <- color(flbsubsft, part = "header", color = "#FFFFFF")
  flbsubsft <- delete_part(flbsubsft, part = "header")
  flbsubsft <- fontsize(flbsubsft, i = 1, part = "body", size = 44)
  flbsubsft <- fontsize(flbsubsft, i = 2, part = "body", size = 12)
  flbsubsft <- bold(flbsubsft, j = 1, part = "body")
  flbsubsft <- border_remove(x = flbsubsft)
  flbsubsft <- width(flbsubsft, j=1, width = 2.5)
  
  #Titles
  
  
  mappath <- paste0("Silhouettes/",CB_Y4$Country[counter],".png")
  flagpath <- paste0("Flags-Icon-Set/",CB_Y4$ISO2[counter],".png")
  
  if(CB_Y4$Country[counter]=="Côte d'Ivoire"){
    mappath <- "Silhouettes/Cote d_Ivoire.png"
  }

  
  doc <- add_slide(doc)
  doc <- ph_with(doc, value = external_img(flagpath), location = ph_location(left = .5, top = .17, width = 1, height = .87))
  doc <- ph_with(doc, value = external_img(mappath), location = ph_location(left = 11.5, top = .17, width = 1, height = .84))
  
  doc <- ph_with(doc, value = paste("Affordability:", CB_Y4$Country[counter]), location = ph_location(left = 1.5, top = .3, width = 10, alignment = "l", height = .6))
  doc <- ph_with(doc, value = affplot,
                 location = ph_location(left = .45, top = 1, height=6.05, width=10))
  doc <- ph_with(doc, value = hsetplot,
                 location = ph_location(left = 10.68, top = 1, height=3, width=2.5))
  doc <- ph_with(doc, fbaffft,
                 location = ph_location(left = 10.68, top = 4.04, height=2, width=2.5))
  doc <- ph_with(doc, flbsubsft,
                 location = ph_location(left = 10.68, top = 5.56, height=2, width=2.5))
  
  print(CB_Y4$Country[counter])
  }
  doc <- add_slide(doc)
  
  
  
  print(doc, target = "Affordability_2GB.pptx")
  
  
  
  
  