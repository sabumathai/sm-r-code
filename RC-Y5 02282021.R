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

options(digits=3)


setwd("/Users/sabumathai/3i Y5 - Revised Regional Decks/")
# getwd()


RC_Y5 <- read_excel("Y5 - Country Briefing.xlsx")
RC_Y5 <- select(RC_Y5, 2,3,7, 15, 23, 31, 39)
RC_Y5$`Overall rank_global` <- RC_Y5$OVERALL_rank
RC_Y5$`Availability_global rank` <-RC_Y5$AVAILABILITY_rank
RC_Y5$`Affordability_overall rank` <- RC_Y5$AFFORDABILITY_rank
RC_Y5$`Relevance_overall rank`<- RC_Y5$RELEVANCE_rank
RC_Y5$`Readiness_overall rank`<- RC_Y5$READINESS_rank
RC_Y5 <- select(RC_Y5, 1, 2, 8:12)

doc <- read_pptx(path = "3i_Y5_report_template.pptx")
# Scores_Y5 <- read_excel("Y4_Scores.xlsx")

# CB_Y4 <- full_join(CB_Y4, Scores_Y4, by = "Country")
# 

RCY5_ASIA <- RC_Y5[RC_Y5$Region=="Asia",]
RCY5_NA <- RC_Y5[RC_Y5$Region=="North America (NA)",]
RCY5_EUR <- RC_Y5[RC_Y5$Region=="Europe (EUR)",]
RCY5_LA <- RC_Y5[RC_Y5$Region=="Latin America (LATAM)",]
RCY5_MENA <- RC_Y5[RC_Y5$Region=="Middle East and North Africa (MENA)",]
RCY5_SSA <- RC_Y5[RC_Y5$Region=="Sub-Saharan Africa (SSA)",]

#Asia
Overall_Asia <- select(RCY5_ASIA, 1,3)
Overall_Asia$removesign <-Overall_Asia$`Overall rank_global`
for(i in seq_along(Overall_Asia$removesign)){
  if(str_sub(Overall_Asia$removesign[i], 1, 1)=='=')
    Overall_Asia$removesign[i] <- str_sub(Overall_Asia$removesign[i], 2, nchar(Overall_Asia$removesign[i]))
}
Overall_Asia$numeric <-as.numeric(Overall_Asia$removesign)
Overall_Asia <- Overall_Asia[order(Overall_Asia$numeric),]
Overall_Asia <-select(Overall_Asia,1,2)

Av_Asia <- select(RCY5_ASIA, 1,4)
Av_Asia$removesign <-Av_Asia$`Availability_global rank`
for(i in seq_along(Av_Asia$removesign)){
  if(str_sub(Av_Asia$removesign[i], 1, 1)=='=')
    Av_Asia$removesign[i] <- str_sub(Av_Asia$removesign[i], 2, nchar(Av_Asia$removesign[i]))
}
Av_Asia$numeric <-as.numeric(Av_Asia$removesign)
Av_Asia <- Av_Asia[order(Av_Asia$numeric),]
Av_Asia <-select(Av_Asia,1,2)

Aff_Asia <- select(RCY5_ASIA, 1,5)
Aff_Asia$removesign <-Aff_Asia$`Affordability_overall rank`
for(i in seq_along(Aff_Asia$removesign)){
  if(str_sub(Aff_Asia$removesign[i], 1, 1)=='=')
    Aff_Asia$removesign[i] <- str_sub(Aff_Asia$removesign[i], 2, nchar(Aff_Asia$removesign[i]))
}
Aff_Asia$numeric <-as.numeric(Aff_Asia$removesign)
Aff_Asia <- Aff_Asia[order(Aff_Asia$numeric),]
Aff_Asia <-select(Aff_Asia,1,2)

Rel_Asia <- select(RCY5_ASIA, 1,6)
Rel_Asia$removesign <-Rel_Asia$`Relevance_overall rank`
for(i in seq_along(Rel_Asia$removesign)){
  if(str_sub(Rel_Asia$removesign[i], 1, 1)=='=')
    Rel_Asia$removesign[i] <- str_sub(Rel_Asia$removesign[i], 2, nchar(Rel_Asia$removesign[i]))
}
Rel_Asia$numeric <-as.numeric(Rel_Asia$removesign)
Rel_Asia <- Rel_Asia[order(Rel_Asia$numeric),]
Rel_Asia <-select(Rel_Asia,1,2)

Read_Asia <- select(RCY5_ASIA, 1,7)
Read_Asia$removesign <-Read_Asia$`Readiness_overall rank`
for(i in seq_along(Read_Asia$removesign)){
  if(str_sub(Read_Asia$removesign[i], 1, 1)=='=')
    Read_Asia$removesign[i] <- str_sub(Read_Asia$removesign[i], 2, nchar(Read_Asia$removesign[i]))
}
Read_Asia$numeric <-as.numeric(Read_Asia$removesign)
Read_Asia <- Read_Asia[order(Read_Asia$numeric),]
Read_Asia <-select(Read_Asia,1,2)

Asia <- cbind(Overall_Asia, Av_Asia, Aff_Asia, Rel_Asia, Read_Asia)
names(Asia)[1] <- paste("Country1")
names(Asia)[3] <- paste("Country2")
names(Asia)[5] <- paste("Country3")
names(Asia)[7] <- paste("Country4")
names(Asia)[9] <- paste("Country5")

Asia <-clean_names(Asia)
Asiaft <- flextable(Asia)
Asiaft

Asiaft <- width(Asiaft, j=1, width = 4.3)
Asiaft <- width(Asiaft, j=2, width = 1.3)
Asiaft <- height_all(Asiaft, height = 0.38, part = "all")
Asiaft <- hrule(Asiaft, rule = "exact")

Asiaft <- border_remove(x = Asiaft)
# Asiaft <- border_outer(Asiaft, part="all", border = fp_border(color="white", width = 1) )
Asiaft <- border_inner_h(Asiaft, part="all", border = fp_border(color="white", width = 1) )
Asiaft <- border_inner_v(Asiaft, part="all", border = fp_border(color="white", width = 1) )
Asiaft <- align(Asiaft, align = "center", part = "header")
# dark blue as background color for header
Asiaft <-  bg(Asiaft, bg = "#2B458F", part = "header")
# slate as background color for body
Asiaft <-  bg(Asiaft, bg = "#EFEDE5", part = "body")
Asiaft <- color(Asiaft, part = "header", color = "#FFFFFF")
Asiaft <- bold(Asiaft, part = "all")
# Asiaft <- bold(Asiaft, i= 1, part = "body")
Asiaft  <- height_all(Asiaft , height = 0.18, part = "body")
Asiaft <- height_all(Asiaft , height = 0.4, part = "header")
Asiaft <- fontsize(Asiaft, part = "all", size = 13)
Asiaft


for(counter in seq_along(RCY5_ASIA$Country)){
  temptable <- Asiaft
  doc <- add_slide(doc)
  doc <- ph_with(doc, value = paste("Inclusive Internet Index:", RCY5_ASIA$Country[counter]), location = ph_location(left = .5, top = .15, width = 12, height = .6))
  
  temptable <- bg(temptable, i = ~ country1 == RCY5_ASIA$Country[counter], j = ~ country1 + overall_rank_global, bg = "#CFE1F4")
  temptable <- bg(temptable, i = ~ country2 == RCY5_ASIA$Country[counter], j = ~ country2 + availability_global_rank, bg = "#CFE1F4")
  temptable <- bg(temptable, i = ~ country3 == RCY5_ASIA$Country[counter], j = ~ country3 + affordability_overall_rank, bg = "#CFE1F4")
  temptable <- bg(temptable, i = ~ country4 == RCY5_ASIA$Country[counter], j = ~ country4 + relevance_overall_rank, bg = "#CFE1F4")
  temptable <- bg(temptable, i = ~ country5 == RCY5_ASIA$Country[counter], j = ~ country5 + readiness_overall_rank, bg = "#CFE1F4")
  temptable
  
  temptable <- merge_at(temptable, i = 1, j = 1:2, part = "header")
  temptable <- merge_at(temptable, i = 1, j = 3:4, part = "header")
  temptable <- merge_at(temptable, i = 1, j = 5:6, part = "header")
  temptable <- merge_at(temptable, i = 1, j = 7:8, part = "header")
  temptable <- merge_at(temptable, i = 1, j = 9:10, part = "header")
  temptable <- set_header_labels(temptable, country1 = "Overall Rank", overall_rank_global ="Overall Rank", country2 = "Availability Rank", availability_global_rank = "Availability Rank",  country3 = "Affordability Rank", affordability_overall_rank = "Affordability Rank", country4 = "Relevance Rank", relevance_overall_rank = "Relevance Rank", country5 = "Readiness Rank", readiness_overall_rank = "Readiness Rank")  
  temptable <- width(temptable, width = .6)   
  temptable <- width(temptable, j=1, width = 1.85)
  temptable <- width(temptable, j=3, width = 1.85)
  temptable <- width(temptable, j=5, width = 1.85)
  temptable <- width(temptable, j=7, width = 1.85)
  temptable <- width(temptable, j=9, width = 1.85)
  
  
  doc <- ph_with(doc, value = temptable, location = ph_location(left = 0.5, top = 0.75, height = 5.95, width = 12.11))
}


Asiaft <- merge_at(Asiaft, i = 1, j = 1:2, part = "header")
Asiaft <- merge_at(Asiaft, i = 1, j = 3:4, part = "header")
Asiaft <- merge_at(Asiaft, i = 1, j = 5:6, part = "header")
Asiaft <- merge_at(Asiaft, i = 1, j = 7:8, part = "header")
Asiaft <- merge_at(Asiaft, i = 1, j = 9:10, part = "header")
Asiaft <- set_header_labels(Asiaft, country1 = "Overall Rank", overall_rank_global ="Overall Rank", country2 = "Availability Rank", availability_global_rank = "Availability Rank",  country3 = "Affordability Rank", affordability_overall_rank = "Affordability Rank", country4 = "Relevance Rank", relevance_overall_rank = "Relevance Rank", country5 = "Readiness Rank", readiness_overall_rank = "Readiness Rank")  
Asiaft <- width(Asiaft, width = .6)  
Asiaft <- width(Asiaft, j=1, width = 1.85)
Asiaft <- width(Asiaft, j=3, width = 1.85)
Asiaft <- width(Asiaft, j=5, width = 1.85)
Asiaft <- width(Asiaft, j=7, width = 1.85)
Asiaft <- width(Asiaft, j=9, width = 1.85)

doc <- add_slide(doc)
doc <- ph_with(doc, value = "Inclusive Internet Index: Asia", location = ph_location(left = .5, top = .15, width = 12, height = .6))
doc <- ph_with(doc, value = Asiaft, location = ph_location(left = 0.5, top = 0.75, height = 5.95, width = 12.11))





#Latam
Overall_LA <- select(RCY5_LA, 1,3)
Overall_LA$removesign <-Overall_LA$`Overall rank_global`
for(i in seq_along(Overall_LA$removesign)){
  if(str_sub(Overall_LA$removesign[i], 1, 1)=='=')
    Overall_LA$removesign[i] <- str_sub(Overall_LA$removesign[i], 2, nchar(Overall_LA$removesign[i]))
}
Overall_LA$numeric <-as.numeric(Overall_LA$removesign)
Overall_LA <- Overall_LA[order(Overall_LA$numeric),]
Overall_LA <-select(Overall_LA,1,2)

Av_LA <- select(RCY5_LA, 1,4)
Av_LA$removesign <-Av_LA$`Availability_global rank`
for(i in seq_along(Av_LA$removesign)){
  if(str_sub(Av_LA$removesign[i], 1, 1)=='=')
    Av_LA$removesign[i] <- str_sub(Av_LA$removesign[i], 2, nchar(Av_LA$removesign[i]))
}
Av_LA$numeric <-as.numeric(Av_LA$removesign)
Av_LA <- Av_LA[order(Av_LA$numeric),]
Av_LA <-select(Av_LA,1,2)

Aff_LA <- select(RCY5_LA, 1,5)
Aff_LA$removesign <-Aff_LA$`Affordability_overall rank`
for(i in seq_along(Aff_LA$removesign)){
  if(str_sub(Aff_LA$removesign[i], 1, 1)=='=')
    Aff_LA$removesign[i] <- str_sub(Aff_LA$removesign[i], 2, nchar(Aff_LA$removesign[i]))
}
Aff_LA$numeric <-as.numeric(Aff_LA$removesign)
Aff_LA <- Aff_LA[order(Aff_LA$numeric),]
Aff_LA <-select(Aff_LA,1,2)

Rel_LA <- select(RCY5_LA, 1,6)
Rel_LA$removesign <-Rel_LA$`Relevance_overall rank`
for(i in seq_along(Rel_LA$removesign)){
  if(str_sub(Rel_LA$removesign[i], 1, 1)=='=')
    Rel_LA$removesign[i] <- str_sub(Rel_LA$removesign[i], 2, nchar(Rel_LA$removesign[i]))
}
Rel_LA$numeric <-as.numeric(Rel_LA$removesign)
Rel_LA <- Rel_LA[order(Rel_LA$numeric),]
Rel_LA <-select(Rel_LA,1,2)

Read_LA <- select(RCY5_LA, 1,7)
Read_LA$removesign <-Read_LA$`Readiness_overall rank`
for(i in seq_along(Read_LA$removesign)){
  if(str_sub(Read_LA$removesign[i], 1, 1)=='=')
    Read_LA$removesign[i] <- str_sub(Read_LA$removesign[i], 2, nchar(Read_LA$removesign[i]))
}
Read_LA$numeric <-as.numeric(Read_LA$removesign)
Read_LA <- Read_LA[order(Read_LA$numeric),]
Read_LA <-select(Read_LA,1,2)

LA <- cbind(Overall_LA, Av_LA, Aff_LA, Rel_LA, Read_LA)
names(LA)[1] <- paste("Country1")
names(LA)[3] <- paste("Country2")
names(LA)[5] <- paste("Country3")
names(LA)[7] <- paste("Country4")
names(LA)[9] <- paste("Country5")

LA <- clean_names(LA)
LAft <- flextable(LA)
LAft

LAft <- width(LAft, j=1, width = 4.3)
LAft <- width(LAft, j=2, width = 1.3)
LAft <- height_all(LAft, height = 0.38, part = "all")
LAft <- hrule(LAft, rule = "exact")

LAft <- border_remove(x = LAft)
# LAft <- border_outer(LAft, part="all", border = fp_border(color="white", width = 1) )
LAft <- border_inner_h(LAft, part="all", border = fp_border(color="white", width = 1) )
LAft <- border_inner_v(LAft, part="all", border = fp_border(color="white", width = 1) )
LAft <- align(LAft, align = "center", part = "header")
# dark blue as background color for header
LAft <-  bg(LAft, bg = "#2B458F", part = "header")
# slate as background color for body
LAft <-  bg(LAft, bg = "#EFEDE5", part = "body")
LAft <- color(LAft, part = "header", color = "#FFFFFF")
LAft <- bold(LAft, part = "all")
# LAft <- bold(LAft, i= 1, part = "body")
LAft  <- height_all(LAft , height = 0.24, part = "body")
LAft <- height_all(LAft , height = 0.4, part = "header")
LAft <- fontsize(LAft, part = "all", size = 13)

LAft


for(counter in seq_along(RCY5_LA$Country)){
  temptable <- LAft
  doc <- add_slide(doc)
  doc <- ph_with(doc, value = paste("Inclusive Internet Index:", RCY5_LA$Country[counter]), location = ph_location(left = .5, top = .15, width = 12, height = .6))
  
  temptable <- bg(temptable, i = ~ country1 == RCY5_LA$Country[counter], j = ~ country1 + overall_rank_global, bg = "#CFE1F4")
  temptable <- bg(temptable, i = ~ country2 == RCY5_LA$Country[counter], j = ~ country2 + availability_global_rank, bg = "#CFE1F4")
  temptable <- bg(temptable, i = ~ country3 == RCY5_LA$Country[counter], j = ~ country3 + affordability_overall_rank, bg = "#CFE1F4")
  temptable <- bg(temptable, i = ~ country4 == RCY5_LA$Country[counter], j = ~ country4 + relevance_overall_rank, bg = "#CFE1F4")
  temptable <- bg(temptable, i = ~ country5 == RCY5_LA$Country[counter], j = ~ country5 + readiness_overall_rank, bg = "#CFE1F4")
  temptable
  
  temptable <- merge_at(temptable, i = 1, j = 1:2, part = "header")
  temptable <- merge_at(temptable, i = 1, j = 3:4, part = "header")
  temptable <- merge_at(temptable, i = 1, j = 5:6, part = "header")
  temptable <- merge_at(temptable, i = 1, j = 7:8, part = "header")
  temptable <- merge_at(temptable, i = 1, j = 9:10, part = "header")
  temptable <- set_header_labels(temptable, country1 = "Overall Rank", overall_rank_global ="Overall Rank", country2 = "Availability Rank", availability_global_rank = "Availability Rank",  country3 = "Affordability Rank", affordability_overall_rank = "Affordability Rank", country4 = "Relevance Rank", relevance_overall_rank = "Relevance Rank", country5 = "Readiness Rank", readiness_overall_rank = "Readiness Rank")  
  temptable <- width(temptable, width = .6)   
  temptable <- width(temptable, j=1, width = 1.85)
  temptable <- width(temptable, j=3, width = 1.85)
  temptable <- width(temptable, j=5, width = 1.85)
  temptable <- width(temptable, j=7, width = 1.85)
  temptable <- width(temptable, j=9, width = 1.85)
  
  
  doc <- ph_with(doc, value = temptable, location = ph_location(left = 0.5, top = 1, height = 5.95, width = 12.35))
}


LAft <- merge_at(LAft, i = 1, j = 1:2, part = "header")
LAft <- merge_at(LAft, i = 1, j = 3:4, part = "header")
LAft <- merge_at(LAft, i = 1, j = 5:6, part = "header")
LAft <- merge_at(LAft, i = 1, j = 7:8, part = "header")
LAft <- merge_at(LAft, i = 1, j = 9:10, part = "header")
LAft <- set_header_labels(LAft, country1 = "Overall Rank", overall_rank_global ="Overall Rank", country2 = "Availability Rank", availability_global_rank = "Availability Rank",  country3 = "Affordability Rank", affordability_overall_rank = "Affordability Rank", country4 = "Relevance Rank", relevance_overall_rank = "Relevance Rank", country5 = "Readiness Rank", readiness_overall_rank = "Readiness Rank")  
LAft <- width(LAft, width = .6)  
LAft <- width(LAft, j=1, width = 1.85)
LAft <- width(LAft, j=3, width = 1.85)
LAft <- width(LAft, j=5, width = 1.85)
LAft <- width(LAft, j=7, width = 1.85)
LAft <- width(LAft, j=9, width = 1.85)

doc <- add_slide(doc)
doc <- ph_with(doc, value = "Inclusive Internet Index: Latin America", location = ph_location(left = .5, top = .15, width = 12, height = .6))
doc <- ph_with(doc, value = LAft, location = ph_location(left = 0.5, top = 1, height = 5.95, width = 12.11))



#NA
Overall_NA <- select(RCY5_NA, 1,3)
Overall_NA$removesign <-Overall_NA$`Overall rank_global`
for(i in seq_along(Overall_NA$removesign)){
  if(str_sub(Overall_NA$removesign[i], 1, 1)=='=')
    Overall_NA$removesign[i] <- str_sub(Overall_NA$removesign[i], 2, nchar(Overall_NA$removesign[i]))
}
Overall_NA$numeric <-as.numeric(Overall_NA$removesign)
Overall_NA <- Overall_NA[order(Overall_NA$numeric),]
Overall_NA <-select(Overall_NA,1,2)

Av_NA <- select(RCY5_NA, 1,4)
Av_NA$removesign <-Av_NA$`Availability_global rank`
for(i in seq_along(Av_NA$removesign)){
  if(str_sub(Av_NA$removesign[i], 1, 1)=='=')
    Av_NA$removesign[i] <- str_sub(Av_NA$removesign[i], 2, nchar(Av_NA$removesign[i]))
}
Av_NA$numeric <-as.numeric(Av_NA$removesign)
Av_NA <- Av_NA[order(Av_NA$numeric),]
Av_NA <-select(Av_NA,1,2)

Aff_NA <- select(RCY5_NA, 1,5)
Aff_NA$removesign <-Aff_NA$`Affordability_overall rank`
for(i in seq_along(Aff_NA$removesign)){
  if(str_sub(Aff_NA$removesign[i], 1, 1)=='=')
    Aff_NA$removesign[i] <- str_sub(Aff_NA$removesign[i], 2, nchar(Aff_NA$removesign[i]))
}
Aff_NA$numeric <-as.numeric(Aff_NA$removesign)
Aff_NA <- Aff_NA[order(Aff_NA$numeric),]
Aff_NA <-select(Aff_NA,1,2)

Rel_NA <- select(RCY5_NA, 1,6)
Rel_NA$removesign <-Rel_NA$`Relevance_overall rank`
for(i in seq_along(Rel_NA$removesign)){
  if(str_sub(Rel_NA$removesign[i], 1, 1)=='=')
    Rel_NA$removesign[i] <- str_sub(Rel_NA$removesign[i], 2, nchar(Rel_NA$removesign[i]))
}
Rel_NA$numeric <-as.numeric(Rel_NA$removesign)
Rel_NA <- Rel_NA[order(Rel_NA$numeric),]
Rel_NA <-select(Rel_NA,1,2)

Read_NA <- select(RCY5_NA, 1,7)
Read_NA$removesign <-Read_NA$`Readiness_overall rank`
for(i in seq_along(Read_NA$removesign)){
  if(str_sub(Read_NA$removesign[i], 1, 1)=='=')
    Read_NA$removesign[i] <- str_sub(Read_NA$removesign[i], 2, nchar(Read_NA$removesign[i]))
}
Read_NA$numeric <-as.numeric(Read_NA$removesign)
Read_NA <- Read_NA[order(Read_NA$numeric),]
Read_NA <-select(Read_NA,1,2)

AllNA <- cbind(Overall_NA, Av_NA, Aff_NA, Rel_NA, Read_NA)
names(AllNA)[1] <- paste("Country1")
names(AllNA)[3] <- paste("Country2")
names(AllNA)[5] <- paste("Country3")
names(AllNA)[7] <- paste("Country4")
names(AllNA)[9] <- paste("Country5")

AllNA <- clean_names(AllNA)
NAft <- flextable(AllNA)
NAft

NAft <- width(NAft, j=1, width = 4.3)
NAft <- width(NAft, j=2, width = 1.3)
NAft <- height_all(NAft, height = 0.38, part = "all")
NAft <- hrule(NAft, rule = "exact")

NAft <- border_remove(x = NAft)
# NAft <- border_outer(NAft, part="all", border = fp_border(color="white", width = 1) )
NAft <- border_inner_h(NAft, part="all", border = fp_border(color="white", width = 1) )
NAft <- border_inner_v(NAft, part="all", border = fp_border(color="white", width = 1) )
NAft <- align(NAft, align = "center", part = "header")
# dark blue as background color for header
NAft <-  bg(NAft, bg = "#2B458F", part = "header")
# slate as background color for body
NAft <-  bg(NAft, bg = "#EFEDE5", part = "body")
NAft <- color(NAft, part = "header", color = "#FFFFFF")
NAft <- bold(NAft, part = "all")
# NAft <- bold(NAft, i= 1, part = "body")
NAft  <- height_all(NAft , height = 0.38, part = "body")
NAft <- height_all(NAft , height = 0.4, part = "header")
NAft <- fontsize(NAft, part = "all", size = 13)
NAft


for(counter in seq_along(RCY5_NA$Country)){
  temptable <- NAft
  doc <- add_slide(doc)
  doc <- ph_with(doc, value = paste("Inclusive Internet Index:", RCY5_NA$Country[counter]), location = ph_location(left = .5, top = .15, width = 12, height = .6))
  
  temptable <- bg(temptable, i = ~ country1 == RCY5_NA$Country[counter], j = ~ country1 + overall_rank_global, bg = "#CFE1F4")
  temptable <- bg(temptable, i = ~ country2 == RCY5_NA$Country[counter], j = ~ country2 + availability_global_rank, bg = "#CFE1F4")
  temptable <- bg(temptable, i = ~ country3 == RCY5_NA$Country[counter], j = ~ country3 + affordability_overall_rank, bg = "#CFE1F4")
  temptable <- bg(temptable, i = ~ country4 == RCY5_NA$Country[counter], j = ~ country4 + relevance_overall_rank, bg = "#CFE1F4")
  temptable <- bg(temptable, i = ~ country5 == RCY5_NA$Country[counter], j = ~ country5 + readiness_overall_rank, bg = "#CFE1F4")
  temptable
  
  temptable <- merge_at(temptable, i = 1, j = 1:2, part = "header")
  temptable <- merge_at(temptable, i = 1, j = 3:4, part = "header")
  temptable <- merge_at(temptable, i = 1, j = 5:6, part = "header")
  temptable <- merge_at(temptable, i = 1, j = 7:8, part = "header")
  temptable <- merge_at(temptable, i = 1, j = 9:10, part = "header")
  temptable <- set_header_labels(temptable, country1 = "Overall Rank", overall_rank_global ="Overall Rank", country2 = "Availability Rank", availability_global_rank = "Availability Rank",  country3 = "Affordability Rank", affordability_overall_rank = "Affordability Rank", country4 = "Relevance Rank", relevance_overall_rank = "Relevance Rank", country5 = "Readiness Rank", readiness_overall_rank = "Readiness Rank")  
  temptable <- width(temptable, width = .6)   
  temptable <- width(temptable, j=1, width = 1.85)
  temptable <- width(temptable, j=3, width = 1.85)
  temptable <- width(temptable, j=5, width = 1.85)
  temptable <- width(temptable, j=7, width = 1.85)
  temptable <- width(temptable, j=9, width = 1.85)
  
  
  doc <- ph_with(doc, value = temptable, location = ph_location(left = 0.5, top = 1, height = 5.95, width = 12.11))
}


NAft <- merge_at(NAft, i = 1, j = 1:2, part = "header")
NAft <- merge_at(NAft, i = 1, j = 3:4, part = "header")
NAft <- merge_at(NAft, i = 1, j = 5:6, part = "header")
NAft <- merge_at(NAft, i = 1, j = 7:8, part = "header")
NAft <- merge_at(NAft, i = 1, j = 9:10, part = "header")
NAft <- set_header_labels(NAft, country1 = "Overall Rank", overall_rank_global ="Overall Rank", country2 = "Availability Rank", availability_global_rank = "Availability Rank",  country3 = "Affordability Rank", affordability_overall_rank = "Affordability Rank", country4 = "Relevance Rank", relevance_overall_rank = "Relevance Rank", country5 = "Readiness Rank", readiness_overall_rank = "Readiness Rank")  
NAft <- width(NAft, width = .6)  
NAft <- width(NAft, j=1, width = 1.85)
NAft <- width(NAft, j=3, width = 1.85)
NAft <- width(NAft, j=5, width = 1.85)
NAft <- width(NAft, j=7, width = 1.85)
NAft <- width(NAft, j=9, width = 1.85)

doc <- add_slide(doc)
doc <- ph_with(doc, value = "Inclusive Internet Index: North America", location = ph_location(left = .5, top = .15, width = 12, height = .6))
doc <- ph_with(doc, value = NAft, location = ph_location(left = 0.5, top = 1, height = 5.95, width = 12.11))


#Eur
Overall_EUR <- select(RCY5_EUR, 1,3)
Overall_EUR$removesign <-Overall_EUR$`Overall rank_global`
for(i in seq_along(Overall_EUR$removesign)){
  if(str_sub(Overall_EUR$removesign[i], 1, 1)=='=')
    Overall_EUR$removesign[i] <- str_sub(Overall_EUR$removesign[i], 2, nchar(Overall_EUR$removesign[i]))
}
Overall_EUR$numeric <-as.numeric(Overall_EUR$removesign)
Overall_EUR <- Overall_EUR[order(Overall_EUR$numeric),]
Overall_EUR <-select(Overall_EUR,1,2)

Av_EUR <- select(RCY5_EUR, 1,4)
Av_EUR$removesign <-Av_EUR$`Availability_global rank`
for(i in seq_along(Av_EUR$removesign)){
  if(str_sub(Av_EUR$removesign[i], 1, 1)=='=')
    Av_EUR$removesign[i] <- str_sub(Av_EUR$removesign[i], 2, nchar(Av_EUR$removesign[i]))
}
Av_EUR$numeric <-as.numeric(Av_EUR$removesign)
Av_EUR <- Av_EUR[order(Av_EUR$numeric),]
Av_EUR <-select(Av_EUR,1,2)

Aff_EUR <- select(RCY5_EUR, 1,5)
Aff_EUR$removesign <-Aff_EUR$`Affordability_overall rank`
for(i in seq_along(Aff_EUR$removesign)){
  if(str_sub(Aff_EUR$removesign[i], 1, 1)=='=')
    Aff_EUR$removesign[i] <- str_sub(Aff_EUR$removesign[i], 2, nchar(Aff_EUR$removesign[i]))
}
Aff_EUR$numeric <-as.numeric(Aff_EUR$removesign)
Aff_EUR <- Aff_EUR[order(Aff_EUR$numeric),]
Aff_EUR <-select(Aff_EUR,1,2)

Rel_EUR <- select(RCY5_EUR, 1,6)
Rel_EUR$removesign <-Rel_EUR$`Relevance_overall rank`
for(i in seq_along(Rel_EUR$removesign)){
  if(str_sub(Rel_EUR$removesign[i], 1, 1)=='=')
    Rel_EUR$removesign[i] <- str_sub(Rel_EUR$removesign[i], 2, nchar(Rel_EUR$removesign[i]))
}
Rel_EUR$numeric <-as.numeric(Rel_EUR$removesign)
Rel_EUR <- Rel_EUR[order(Rel_EUR$numeric),]
Rel_EUR <-select(Rel_EUR,1,2)

Read_EUR <- select(RCY5_EUR, 1,7)
Read_EUR$removesign <-Read_EUR$`Readiness_overall rank`
for(i in seq_along(Read_EUR$removesign)){
  if(str_sub(Read_EUR$removesign[i], 1, 1)=='=')
    Read_EUR$removesign[i] <- str_sub(Read_EUR$removesign[i], 2, nchar(Read_EUR$removesign[i]))
}
Read_EUR$numeric <-as.numeric(Read_EUR$removesign)
Read_EUR <- Read_EUR[order(Read_EUR$numeric),]
Read_EUR <-select(Read_EUR,1,2)

AllEUR <- cbind(Overall_EUR, Av_EUR, Aff_EUR, Rel_EUR, Read_EUR)
names(AllEUR)[1] <- paste("Country1")
names(AllEUR)[3] <- paste("Country2")
names(AllEUR)[5] <- paste("Country3")
names(AllEUR)[7] <- paste("Country4")
names(AllEUR)[9] <- paste("Country5")

AllEUR <- clean_names(AllEUR)
EURft <- flextable(AllEUR)
EURft

EURft <- width(EURft, j=1, width = 4.3)
EURft <- width(EURft, j=2, width = 1.3)
EURft <- height_all(EURft, height = 0.38, part = "all")
EURft <- hrule(EURft, rule = "exact")

EURft <- border_remove(x = EURft)
# EURft <- border_outer(EURft, part="all", border = fp_border(color="white", width = 1) )
EURft <- border_inner_h(EURft, part="all", border = fp_border(color="white", width = 1) )
EURft <- border_inner_v(EURft, part="all", border = fp_border(color="white", width = 1) )
EURft <- align(EURft, align = "center", part = "header")
# dark blue as background color for header
EURft <-  bg(EURft, bg = "#2B458F", part = "header")
# slate as background color for body
EURft <-  bg(EURft, bg = "#EFEDE5", part = "body")
EURft <- color(EURft, part = "header", color = "#FFFFFF")
EURft <- bold(EURft, part = "all")
# EURft <- bold(EURft, i= 1, part = "body")
EURft  <- height_all(EURft , height = 0.18, part = "body")
EURft <- height_all(EURft , height = 0.4, part = "header")
EURft <- fontsize(EURft, part = "all", size = 12.5)
EURft


for(counter in seq_along(RCY5_EUR$Country)){
  temptable <- EURft
  doc <- add_slide(doc)
  doc <- ph_with(doc, value = paste("Inclusive Internet Index:", RCY5_EUR$Country[counter]), location = ph_location(left = .5, top = .15, width = 12, height = .6))
  
  temptable <- bg(temptable, i = ~ country1 == RCY5_EUR$Country[counter], j = ~ country1 + overall_rank_global, bg = "#CFE1F4")
  temptable <- bg(temptable, i = ~ country2 == RCY5_EUR$Country[counter], j = ~ country2 + availability_global_rank, bg = "#CFE1F4")
  temptable <- bg(temptable, i = ~ country3 == RCY5_EUR$Country[counter], j = ~ country3 + affordability_overall_rank, bg = "#CFE1F4")
  temptable <- bg(temptable, i = ~ country4 == RCY5_EUR$Country[counter], j = ~ country4 + relevance_overall_rank, bg = "#CFE1F4")
  temptable <- bg(temptable, i = ~ country5 == RCY5_EUR$Country[counter], j = ~ country5 + readiness_overall_rank, bg = "#CFE1F4")
  temptable
  
  temptable <- merge_at(temptable, i = 1, j = 1:2, part = "header")
  temptable <- merge_at(temptable, i = 1, j = 3:4, part = "header")
  temptable <- merge_at(temptable, i = 1, j = 5:6, part = "header")
  temptable <- merge_at(temptable, i = 1, j = 7:8, part = "header")
  temptable <- merge_at(temptable, i = 1, j = 9:10, part = "header")
  temptable <- set_header_labels(temptable, country1 = "Overall Rank", overall_rank_global ="Overall Rank", country2 = "Availability Rank", availability_global_rank = "Availability Rank",  country3 = "Affordability Rank", affordability_overall_rank = "Affordability Rank", country4 = "Relevance Rank", relevance_overall_rank = "Relevance Rank", country5 = "Readiness Rank", readiness_overall_rank = "Readiness Rank")  
  temptable <- width(temptable, width = .75)   
  temptable <- width(temptable, j=1, width = 1.7)
  temptable <- width(temptable, j=3, width = 1.7)
  temptable <- width(temptable, j=5, width = 1.7)
  temptable <- width(temptable, j=7, width = 1.7)
  temptable <- width(temptable, j=9, width = 1.7)
  
  
  doc <- ph_with(doc, value = temptable, location = ph_location(left = 0.5, top = .75, height = 5.95, width = 12.11))
}


EURft <- merge_at(EURft, i = 1, j = 1:2, part = "header")
EURft <- merge_at(EURft, i = 1, j = 3:4, part = "header")
EURft <- merge_at(EURft, i = 1, j = 5:6, part = "header")
EURft <- merge_at(EURft, i = 1, j = 7:8, part = "header")
EURft <- merge_at(EURft, i = 1, j = 9:10, part = "header")
EURft <- set_header_labels(EURft, country1 = "Overall Rank", overall_rank_global ="Overall Rank", country2 = "Availability Rank", availability_global_rank = "Availability Rank",  country3 = "Affordability Rank", affordability_overall_rank = "Affordability Rank", country4 = "Relevance Rank", relevance_overall_rank = "Relevance Rank", country5 = "Readiness Rank", readiness_overall_rank = "Readiness Rank")  
EURft <- width(EURft, width = .75)  
EURft <- width(EURft, j=1, width = 1.7)
EURft <- width(EURft, j=3, width = 1.7)
EURft <- width(EURft, j=5, width = 1.7)
EURft <- width(EURft, j=7, width = 1.7)
EURft <- width(EURft, j=9, width = 1.7)

doc <- add_slide(doc)
doc <- ph_with(doc, value = "Inclusive Internet Index: Europe", location = ph_location(left = .5, top = .15, width = 12, height = .6))
doc <- ph_with(doc, value = EURft, location = ph_location(left = 0.5, top = 0.75, height = 5.95, width = 12.11))


#SSA
Overall_SSA <- select(RCY5_SSA, 1,3)
Overall_SSA$removesign <-Overall_SSA$`Overall rank_global`
for(i in seq_along(Overall_SSA$removesign)){
  if(str_sub(Overall_SSA$removesign[i], 1, 1)=='=')
    Overall_SSA$removesign[i] <- str_sub(Overall_SSA$removesign[i], 2, nchar(Overall_SSA$removesign[i]))
}
Overall_SSA$numeric <-as.numeric(Overall_SSA$removesign)
Overall_SSA <- Overall_SSA[order(Overall_SSA$numeric),]
Overall_SSA <-select(Overall_SSA,1,2)

Av_SSA <- select(RCY5_SSA, 1,4)
Av_SSA$removesign <-Av_SSA$`Availability_global rank`
for(i in seq_along(Av_SSA$removesign)){
  if(str_sub(Av_SSA$removesign[i], 1, 1)=='=')
    Av_SSA$removesign[i] <- str_sub(Av_SSA$removesign[i], 2, nchar(Av_SSA$removesign[i]))
}
Av_SSA$numeric <-as.numeric(Av_SSA$removesign)
Av_SSA <- Av_SSA[order(Av_SSA$numeric),]
Av_SSA <-select(Av_SSA,1,2)

Aff_SSA <- select(RCY5_SSA, 1,5)
Aff_SSA$removesign <-Aff_SSA$`Affordability_overall rank`
for(i in seq_along(Aff_SSA$removesign)){
  if(str_sub(Aff_SSA$removesign[i], 1, 1)=='=')
    Aff_SSA$removesign[i] <- str_sub(Aff_SSA$removesign[i], 2, nchar(Aff_SSA$removesign[i]))
}
Aff_SSA$numeric <-as.numeric(Aff_SSA$removesign)
Aff_SSA <- Aff_SSA[order(Aff_SSA$numeric),]
Aff_SSA <-select(Aff_SSA,1,2)

Rel_SSA <- select(RCY5_SSA, 1,6)
Rel_SSA$removesign <-Rel_SSA$`Relevance_overall rank`
for(i in seq_along(Rel_SSA$removesign)){
  if(str_sub(Rel_SSA$removesign[i], 1, 1)=='=')
    Rel_SSA$removesign[i] <- str_sub(Rel_SSA$removesign[i], 2, nchar(Rel_SSA$removesign[i]))
}
Rel_SSA$numeric <-as.numeric(Rel_SSA$removesign)
Rel_SSA <- Rel_SSA[order(Rel_SSA$numeric),]
Rel_SSA <-select(Rel_SSA,1,2)

Read_SSA <- select(RCY5_SSA, 1,7)
Read_SSA$removesign <-Read_SSA$`Readiness_overall rank`
for(i in seq_along(Read_SSA$removesign)){
  if(str_sub(Read_SSA$removesign[i], 1, 1)=='=')
    Read_SSA$removesign[i] <- str_sub(Read_SSA$removesign[i], 2, nchar(Read_SSA$removesign[i]))
}
Read_SSA$numeric <-as.numeric(Read_SSA$removesign)
Read_SSA <- Read_SSA[order(Read_SSA$numeric),]
Read_SSA <-select(Read_SSA,1,2)

AllSSA <- cbind(Overall_SSA, Av_SSA, Aff_SSA, Rel_SSA, Read_SSA)
names(AllSSA)[1] <- paste("Country1")
names(AllSSA)[3] <- paste("Country2")
names(AllSSA)[5] <- paste("Country3")
names(AllSSA)[7] <- paste("Country4")
names(AllSSA)[9] <- paste("Country5")

AllSSA <- clean_names(AllSSA)
SSAft <- flextable(AllSSA)
SSAft

SSAft <- width(SSAft, j=1, width = 4.3)
SSAft <- width(SSAft, j=2, width = 1.3)
SSAft <- height_all(SSAft, height = 0.38, part = "all")
SSAft <- hrule(SSAft, rule = "exact")

SSAft <- border_remove(x = SSAft)
# SSAft <- border_outer(SSAft, part="all", border = fp_border(color="white", width = 1) )
SSAft <- border_inner_h(SSAft, part="all", border = fp_border(color="white", width = 1) )
SSAft <- border_inner_v(SSAft, part="all", border = fp_border(color="white", width = 1) )
SSAft <- align(SSAft, align = "center", part = "header")
# dark blue as background color for header
SSAft <-  bg(SSAft, bg = "#2B458F", part = "header")
# slate as background color for body
SSAft <-  bg(SSAft, bg = "#EFEDE5", part = "body")
SSAft <- color(SSAft, part = "header", color = "#FFFFFF")
SSAft <- bold(SSAft, part = "all")
# SSAft <- bold(SSAft, i= 1, part = "body")
SSAft  <- height_all(SSAft , height = 0.18, part = "body")
SSAft <- height_all(SSAft , height = 0.4, part = "header")
SSAft <- fontsize(SSAft, part = "all", size = 12)
SSAft


for(counter in seq_along(RCY5_SSA$Country)){
  temptable <- SSAft
  doc <- add_slide(doc)
  doc <- ph_with(doc, value = paste("Inclusive Internet Index:", RCY5_SSA$Country[counter]), location = ph_location(left = .5, top = .15, width = 12, height = .6))
  
  temptable <- bg(temptable, i = ~ country1 == RCY5_SSA$Country[counter], j = ~ country1 + overall_rank_global, bg = "#CFE1F4")
  temptable <- bg(temptable, i = ~ country2 == RCY5_SSA$Country[counter], j = ~ country2 + availability_global_rank, bg = "#CFE1F4")
  temptable <- bg(temptable, i = ~ country3 == RCY5_SSA$Country[counter], j = ~ country3 + affordability_overall_rank, bg = "#CFE1F4")
  temptable <- bg(temptable, i = ~ country4 == RCY5_SSA$Country[counter], j = ~ country4 + relevance_overall_rank, bg = "#CFE1F4")
  temptable <- bg(temptable, i = ~ country5 == RCY5_SSA$Country[counter], j = ~ country5 + readiness_overall_rank, bg = "#CFE1F4")
  temptable
  
  temptable <- merge_at(temptable, i = 1, j = 1:2, part = "header")
  temptable <- merge_at(temptable, i = 1, j = 3:4, part = "header")
  temptable <- merge_at(temptable, i = 1, j = 5:6, part = "header")
  temptable <- merge_at(temptable, i = 1, j = 7:8, part = "header")
  temptable <- merge_at(temptable, i = 1, j = 9:10, part = "header")
  temptable <- set_header_labels(temptable, country1 = "Overall Rank", overall_rank_global ="Overall Rank", country2 = "Availability Rank", availability_global_rank = "Availability Rank",  country3 = "Affordability Rank", affordability_overall_rank = "Affordability Rank", country4 = "Relevance Rank", relevance_overall_rank = "Relevance Rank", country5 = "Readiness Rank", readiness_overall_rank = "Readiness Rank")  
  temptable <- width(temptable, width = .6)   
  temptable <- width(temptable, j=1, width = 1.85)
  temptable <- width(temptable, j=3, width = 1.85)
  temptable <- width(temptable, j=5, width = 1.85)
  temptable <- width(temptable, j=7, width = 1.85)
  temptable <- width(temptable, j=9, width = 1.85)
  
  
  doc <- ph_with(doc, value = temptable, location = ph_location(left = 0.5, top = .75, height = 5.95, width = 12.11))
}


SSAft <- merge_at(SSAft, i = 1, j = 1:2, part = "header")
SSAft <- merge_at(SSAft, i = 1, j = 3:4, part = "header")
SSAft <- merge_at(SSAft, i = 1, j = 5:6, part = "header")
SSAft <- merge_at(SSAft, i = 1, j = 7:8, part = "header")
SSAft <- merge_at(SSAft, i = 1, j = 9:10, part = "header")
SSAft <- set_header_labels(SSAft, country1 = "Overall Rank", overall_rank_global ="Overall Rank", country2 = "Availability Rank", availability_global_rank = "Availability Rank",  country3 = "Affordability Rank", affordability_overall_rank = "Affordability Rank", country4 = "Relevance Rank", relevance_overall_rank = "Relevance Rank", country5 = "Readiness Rank", readiness_overall_rank = "Readiness Rank")  
SSAft <- width(SSAft, width = .6)  
SSAft <- width(SSAft, j=1, width = 1.85)
SSAft <- width(SSAft, j=3, width = 1.85)
SSAft <- width(SSAft, j=5, width = 1.85)
SSAft <- width(SSAft, j=7, width = 1.85)
SSAft <- width(SSAft, j=9, width = 1.85)

doc <- add_slide(doc)
doc <- ph_with(doc, value = "Inclusive Internet Index: Sub-Saharan Africa", location = ph_location(left = .5, top = .15, width = 12, height = .6))
doc <- ph_with(doc, value = SSAft, location = ph_location(left = 0.5, top = .75, height = 5.95, width = 12.11))





#MENA
Overall_MENA <- select(RCY5_MENA, 1,3)
Overall_MENA$removesign <-Overall_MENA$`Overall rank_global`
for(i in seq_along(Overall_MENA$removesign)){
  if(str_sub(Overall_MENA$removesign[i], 1, 1)=='=')
    Overall_MENA$removesign[i] <- str_sub(Overall_MENA$removesign[i], 2, nchar(Overall_MENA$removesign[i]))
}
Overall_MENA$numeric <-as.numeric(Overall_MENA$removesign)
Overall_MENA <- Overall_MENA[order(Overall_MENA$numeric),]
Overall_MENA <-select(Overall_MENA,1,2)

Av_MENA <- select(RCY5_MENA, 1,4)
Av_MENA$removesign <-Av_MENA$`Availability_global rank`
for(i in seq_along(Av_MENA$removesign)){
  if(str_sub(Av_MENA$removesign[i], 1, 1)=='=')
    Av_MENA$removesign[i] <- str_sub(Av_MENA$removesign[i], 2, nchar(Av_MENA$removesign[i]))
}
Av_MENA$numeric <-as.numeric(Av_MENA$removesign)
Av_MENA <- Av_MENA[order(Av_MENA$numeric),]
Av_MENA <-select(Av_MENA,1,2)

Aff_MENA <- select(RCY5_MENA, 1,5)
Aff_MENA$removesign <-Aff_MENA$`Affordability_overall rank`
for(i in seq_along(Aff_MENA$removesign)){
  if(str_sub(Aff_MENA$removesign[i], 1, 1)=='=')
    Aff_MENA$removesign[i] <- str_sub(Aff_MENA$removesign[i], 2, nchar(Aff_MENA$removesign[i]))
}
Aff_MENA$numeric <-as.numeric(Aff_MENA$removesign)
Aff_MENA <- Aff_MENA[order(Aff_MENA$numeric),]
Aff_MENA <-select(Aff_MENA,1,2)

Rel_MENA <- select(RCY5_MENA, 1,6)
Rel_MENA$removesign <-Rel_MENA$`Relevance_overall rank`
for(i in seq_along(Rel_MENA$removesign)){
  if(str_sub(Rel_MENA$removesign[i], 1, 1)=='=')
    Rel_MENA$removesign[i] <- str_sub(Rel_MENA$removesign[i], 2, nchar(Rel_MENA$removesign[i]))
}
Rel_MENA$numeric <-as.numeric(Rel_MENA$removesign)
Rel_MENA <- Rel_MENA[order(Rel_MENA$numeric),]
Rel_MENA <-select(Rel_MENA,1,2)

Read_MENA <- select(RCY5_MENA, 1,7)
Read_MENA$removesign <-Read_MENA$`Readiness_overall rank`
for(i in seq_along(Read_MENA$removesign)){
  if(str_sub(Read_MENA$removesign[i], 1, 1)=='=')
    Read_MENA$removesign[i] <- str_sub(Read_MENA$removesign[i], 2, nchar(Read_MENA$removesign[i]))
}
Read_MENA$numeric <-as.numeric(Read_MENA$removesign)
Read_MENA <- Read_MENA[order(Read_MENA$numeric),]
Read_MENA <-select(Read_MENA,1,2)

AllMENA <- cbind(Overall_MENA, Av_MENA, Aff_MENA, Rel_MENA, Read_MENA)
names(AllMENA)[1] <- paste("Country1")
names(AllMENA)[3] <- paste("Country2")
names(AllMENA)[5] <- paste("Country3")
names(AllMENA)[7] <- paste("Country4")
names(AllMENA)[9] <- paste("Country5")

AllMENA <- clean_names(AllMENA)
MENAft <- flextable(AllMENA)
MENAft


MENAft <- width(MENAft, j=1, width = 4.3)
MENAft <- width(MENAft, j=2, width = 1.3)
MENAft <- height_all(MENAft, height = 0.38, part = "all")
MENAft <- hrule(MENAft, rule = "exact")

MENAft <- border_remove(x = MENAft)
# MENAft <- border_outer(MENAft, part="all", border = fp_border(color="white", width = 1) )
MENAft <- border_inner_h(MENAft, part="all", border = fp_border(color="white", width = 1) )
MENAft <- border_inner_v(MENAft, part="all", border = fp_border(color="white", width = 1) )
MENAft <- align(MENAft, align = "center", part = "header")
# dark blue as background color for header
MENAft <-  bg(MENAft, bg = "#2B458F", part = "header")
# slate as background color for body
MENAft <-  bg(MENAft, bg = "#EFEDE5", part = "body")
MENAft <- color(MENAft, part = "header", color = "#FFFFFF")
MENAft <- bold(MENAft, part = "all")
# MENAft <- bold(MENAft, i= 1, part = "body")
MENAft  <- height_all(MENAft , height = 0.38, part = "body")
MENAft <- height_all(MENAft , height = 0.4, part = "header")
MENAft <- fontsize(MENAft, part = "all", size = 14)
MENAft


for(counter in seq_along(RCY5_MENA$Country)){
  temptable <- MENAft
  doc <- add_slide(doc)
  doc <- ph_with(doc, value = paste("Inclusive Internet Index:", RCY5_MENA$Country[counter]), location = ph_location(left = .5, top = .15, width = 12, height = .6))
  
  temptable <- bg(temptable, i = ~ country1 == RCY5_MENA$Country[counter], j = ~ country1 + overall_rank_global, bg = "#CFE1F4")
  temptable <- bg(temptable, i = ~ country2 == RCY5_MENA$Country[counter], j = ~ country2 + availability_global_rank, bg = "#CFE1F4")
  temptable <- bg(temptable, i = ~ country3 == RCY5_MENA$Country[counter], j = ~ country3 + affordability_overall_rank, bg = "#CFE1F4")
  temptable <- bg(temptable, i = ~ country4 == RCY5_MENA$Country[counter], j = ~ country4 + relevance_overall_rank, bg = "#CFE1F4")
  temptable <- bg(temptable, i = ~ country5 == RCY5_MENA$Country[counter], j = ~ country5 + readiness_overall_rank, bg = "#CFE1F4")
  temptable
  
  temptable <- merge_at(temptable, i = 1, j = 1:2, part = "header")
  temptable <- merge_at(temptable, i = 1, j = 3:4, part = "header")
  temptable <- merge_at(temptable, i = 1, j = 5:6, part = "header")
  temptable <- merge_at(temptable, i = 1, j = 7:8, part = "header")
  temptable <- merge_at(temptable, i = 1, j = 9:10, part = "header")
  temptable <- set_header_labels(temptable, country1 = "Overall Rank", overall_rank_global ="Overall Rank", country2 = "Availability Rank", availability_global_rank = "Availability Rank",  country3 = "Affordability Rank", affordability_overall_rank = "Affordability Rank", country4 = "Relevance Rank", relevance_overall_rank = "Relevance Rank", country5 = "Readiness Rank", readiness_overall_rank = "Readiness Rank")  
  temptable <- width(temptable, width = .6)
  temptable <- width(temptable, j=1, width = 1.85)
  temptable <- width(temptable, j=3, width = 1.85)
  temptable <- width(temptable, j=5, width = 1.85)
  temptable <- width(temptable, j=7, width = 1.85)
  temptable <- width(temptable, j=9, width = 1.85)
  
  
  doc <- ph_with(doc, value = temptable, location = ph_location(left = 0.5, top = 1, height = 5.95, width = 12.11))
}


MENAft <- merge_at(MENAft, i = 1, j = 1:2, part = "header")
MENAft <- merge_at(MENAft, i = 1, j = 3:4, part = "header")
MENAft <- merge_at(MENAft, i = 1, j = 5:6, part = "header")
MENAft <- merge_at(MENAft, i = 1, j = 7:8, part = "header")
MENAft <- merge_at(MENAft, i = 1, j = 9:10, part = "header")
MENAft <- set_header_labels(MENAft, country1 = "Overall Rank", overall_rank_global ="Overall Rank", country2 = "Availability Rank", availability_global_rank = "Availability Rank",  country3 = "Affordability Rank", affordability_overall_rank = "Affordability Rank", country4 = "Relevance Rank", relevance_overall_rank = "Relevance Rank", country5 = "Readiness Rank", readiness_overall_rank = "Readiness Rank")  
MENAft <- width(MENAft, width = .6)  
MENAft <- width(MENAft, j=1, width = 1.85)
MENAft <- width(MENAft, j=3, width = 1.85)
MENAft <- width(MENAft, j=5, width = 1.85)
MENAft <- width(MENAft, j=7, width = 1.85)
MENAft <- width(MENAft, j=9, width = 1.85)

doc <- add_slide(doc)
doc <- ph_with(doc, value = "Inclusive Internet Index: MENA", location = ph_location(left = .5, top = .15, width = 12, height = .6))
doc <- ph_with(doc, value = MENAft, location = ph_location(left = 0.5, top = 1, height = 5.95, width = 12.11))






print(doc, target = "3i_Y5_RC.pptx")









