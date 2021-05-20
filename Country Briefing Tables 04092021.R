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
  
  ggdata_wide <- select(Model_exp, 1:2, 6:7, 82:85)
  ggdata_wide <- ggdata_wide %>%
    pivot_wider(names_from = Edition, values_from = c("1.1.4) Gender gap in internet access",	"1.1.5) Gender gap in mobile phone access", "BG20) Male internet users",	"BG21) Female internet users",	"BG22) Male mobile phone subscribers",	"BG23) Female mobile phone subscribers"))
  
  
  ookla_quality_wide <- select(Model_exp, 1:2, 8:14)
  
  ookla_quality_wide <- ookla_quality_wide %>%
    pivot_wider(names_from = Edition, values_from = c("1.2.1) Average fixed broadband upload speed", "1.2.2) Average fixed broadband download speed", "1.2.3) Average fixed broadband latency", "1.2.4) Average mobile upload speed",	"1.2.5) Average mobile download speed",	"1.2.6) Average mobile latency",	"1.2.7) Bandwidth capacity"))
  
  # ookla_quality_wide_save <- ookla_quality_wide
  
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
  
  # ookla_quality_wide <- select(ookla_quality_wide, 1, 37:48)
  
  gg_wide <- select(Model_exp, Country, Edition, `1.1.4) Gender gap in internet access`)
  gg_wide <- gg_wide %>%
    pivot_wider(names_from = Edition, values_from = c("1.1.4) Gender gap in internet access"))
  gg_wide$PP_DIFF <- gg_wide$E5 - gg_wide$E4
  gg_wide <- select(gg_wide, Country, PP_DIFF)
  
  
  E4E5_QUALITY_WIDE <- select(Model_exp[Model_exp$Edition=="E5" | Model_exp$Edition=="E4", ], 1:2, 8:14)
  
  E4E5_QUALITY_WIDE <- E4E5_QUALITY_WIDE %>%
    pivot_wider(names_from = Edition, values_from = c("1.2.1) Average fixed broadband upload speed", "1.2.2) Average fixed broadband download speed", "1.2.3) Average fixed broadband latency", "1.2.4) Average mobile upload speed",	"1.2.5) Average mobile download speed",	"1.2.6) Average mobile latency",	"1.2.7) Bandwidth capacity"))
  
  # CB_Y4 <- read_excel("Y5 - Country Briefing.xlsx")
  CB_Y4 <- CB1
  
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
  
  CB_Y4$`Largest YoY Change1_change_exclude_survey_and_qual` <- CB_Y4$`subset5_Largest YoY_Data_Change_1_change`
  CB_Y4$`Largest YoY Change2_change_exclude_survey_and_qual` <- CB_Y4$`subset5_Largest YoY_Data_Change_2_change`
  CB_Y4$`Largest YoY Change1_name_exclude_survey_and_qual` <- CB_Y4$`subset5_Largest YoY_Data_Change_1_name`
  CB_Y4$`Largest YoY Change2_name_exclude_survey_and_qual` <- CB_Y4$`subset5_Largest YoY_Data_Change_2_name`
  
  
  
  
  Y4_data_raw <- read_excel("3i Y5 - dataset.xlsx")
  Y4_temp_dr <- select(Y4_data_raw, Country, fmai)
  CB_Y4 <- left_join(CB_Y4, Y4_temp_dr, by = "Country")
  CB_Y4 <- left_join(CB_Y4, gg_wide, by = "Country")
  CB_Y4 <- left_join(CB_Y4, ookla_quality_wide, by = "Country")
  
  # Readinessscores <- select(Scores_Y4, 1, 11,14)
  # Y4_data_raw <- full_join(Y4_data_raw, Readinessscores, by = "Country")
  
  # Y4_data_ro <- select(Y4_data_raw, 1:5, 9:20, 67:90) 
  
  
  
  # Y4_Scores_Spider_Data <- select(Scores_Y4, c(1, 4:7, 9:10, 12:13, 15:17))
  # colnames(Y4_Scores_Spider_Data) <- c("Country", "Usage", "Quality", "Infrastructure      ", "Electricity   ", "Price", "Competitive\nEnvironment", "Local\nContent", "  Relevant\nContent", "Literacy", "Trust &\nSafety", "Policy")
  # rownames(Y4_Scores_Spider_Data) <- Y4_Scores_Spider_Data$Country
  # data_for_radar <- Y4_Scores_Spider_Data
  # data_for_radar_validate <- data_for_radar
  # rownames(data_for_radar) <- Y4_Scores_Spider_Data$Country
  # Y4_Country_List <- Y4_Scores_Spider_Data$Country
  # data_for_radar_copy <- data_for_radar 
  
  regional_quality_data <- CB_Y4 %>%
    group_by(Region) %>%
    summarise(across(where(is.numeric), mean, na.rm = TRUE))
  
  regional_quality_data <- select(regional_quality_data, 1, 93:139)
  
  overall_quality_data <- CB_Y4 %>%
    summarise(across(where(is.numeric), mean, na.rm = TRUE))
  overall_quality_data$Region <- "Overall"
  overall_quality_data <- select(overall_quality_data, 139, 92:138)
  
  averages_quality_data <- rbind(regional_quality_data, overall_quality_data)
  
  averages_quality_data <- select(averages_quality_data, 1:31)
  
  temp_colnames <- colnames(averages_quality_data)
  
  colnames(averages_quality_data) <- c("Region", "1.2.1) FB UL_AE1", "1.2.1) FB UL_AE2", "1.2.1) FB UL_AE3", "1.2.1) FB UL_AE4","1.2.1) FB UL_AE5", "1.2.2) FB DL_AE1", "1.2.2) FB DL_AE2", "1.2.2) FB DL_AE3",  "1.2.2) FB DL_AE4", "1.2.2) FB DL_AE5", "1.2.3) FB L_AE1", "1.2.3) FB L_AE2", "1.2.3) FB L_AE3", "1.2.3) FB L_AE4", "1.2.3) FB L_AE5","1.2.4) M UL_AE1", "1.2.4) M UL_AE2", "1.2.4) M UL_AE3", "1.2.4) M UL_AE4","1.2.4) M UL_AE5", "1.2.5) M DL_AE1","1.2.5) M DL_AE2", "1.2.5) M DL_AE3", "1.2.5) M DL_AE4", "1.2.5) M DL_AE5", "1.2.6) M L_AE1", "1.2.6) M L_AE2", "1.2.6) M L_AE3", "1.2.6) M L_AE4", "1.2.6) M L_AE5") 
  CB_Y4 <- left_join(CB_Y4, averages_quality_data, by = "Region")
  
  
  CB_Y4 <- CB_Y4[order(CB_Y4$Region),]
  
  
  doc <- read_pptx(path = "3i_Y5_report_template.pptx")
  
  
  
  
  
  for(counter in seq_along(CB_Y4$Country)){
  
  Y4_data <- Y4_data_raw
  
  
  # data_for_radar_copy <- data_for_radar
  # data_for_radar_copy <- data_for_radar_copy %>%
  #   filter(Country == CB_Y4$Country[counter])
  # data_for_radar_copy <- select(data_for_radar_copy, -Country)
  # ggradar(data_for_radar_copy)
  
  
    
  scorerankdata <- data.table(va=c(paste("Score\n",CB_Y4$`1) AVAILABILITY`[counter]), "Usage", CB_Y4$`1.1) USAGE`[counter], "Quality", CB_Y4$`1.2) QUALITY†††`[counter], "Infrastructure", CB_Y4$`1.3) INFRASTRUCTURE`[counter], "Electricity",CB_Y4$`1.4) ELECTRICITY`[counter]),
                      vb=c(paste("Rank\n", CB_Y4$`Availability_global rank`[counter]), "Usage", CB_Y4$Usage_rank[counter], "Quality", CB_Y4$Quality_rank[counter], "Infrastructure", CB_Y4$Infrastructure_rank[counter], "Electricity",CB_Y4$Electricity_rank[counter]), 
                      vc=c(paste("Score\n",CB_Y4$`2) AFFORDABILITY`[counter]), "Price", CB_Y4$`2.1) PRICE`[counter], "Competitive Environment", CB_Y4$`2.2) COMPETITIVE ENVIRONMENT`[counter], "", "", "","" ),
                      vd=c(paste("Rank\n",CB_Y4$`Affordability_overall rank`[counter]), "Price", CB_Y4$Price_rank[counter], "Competitive Environment", CB_Y4$Competitive_environment_rank[counter], "", "", "","" ),
                      ve=c(paste("Score\n", CB_Y4$`3) RELEVANCE`[counter]), "Local Content", CB_Y4$`3.1) LOCAL CONTENT†††`[counter], "Relevant Content", CB_Y4$`3.2) RELEVANT CONTENT`[counter], "", "", "","" ),
                      vf=c(paste("Rank\n",CB_Y4$`Relevance_overall rank`[counter]), "Local Content", CB_Y4$Local_content_rank[counter], "Relevant Content", CB_Y4$Relevant_content_rank[counter], "", "", "","" ),
                      vg=c(paste("Score\n",CB_Y4$`4) READINESS`[counter]), "Literacy", CB_Y4$`4.1) LITERACY`[counter], "Trust & Safety", CB_Y4$`4.2) TRUST & SAFETY`[counter], "Policy", CB_Y4$`4.3) POLICY`[counter], "","" ),
                      vh=c(paste("Rank\n",CB_Y4$`Readiness_overall rank`[counter]), "Literacy", CB_Y4$Literacy_rank[counter], "Trust & Safety", CB_Y4$`Trust_&_Safety_rank`[counter], "Policy", CB_Y4$Policy_rank[counter], "","" ))
  # scorerankdata
  
  
  # scorerankdata <- phl_adjust_table(scorerankdata, olay = cb_offLayout, id = 5)
  scorerank <- flextable(scorerankdata)
  
  scorerank <- set_header_labels(scorerank, va= "Availability", vb="Availability", vc="Affordability", vd="Affordability", ve="Relevance", vf="Relevance", vg="Readiness", vh="Readiness" )                     
  scorerank <- merge_at(scorerank, i = 1, j = 1:2, part = "header")
  scorerank <- merge_at(scorerank, i = 1, j = 3:4, part = "header")
  scorerank <- merge_at(scorerank, i = 1, j = 5:6, part = "header")
  scorerank <- merge_at(scorerank, i = 1, j = 7:8, part = "header")
  scorerank <- merge_at(scorerank, i = 2, j = 1:2, part = "body")
  scorerank <- merge_at(scorerank, i = 2, j = 3:4, part = "body")
  scorerank <- merge_at(scorerank, i = 2, j = 5:6, part = "body")
  scorerank <- merge_at(scorerank, i = 2, j = 7:8, part = "body")
  scorerank <- merge_at(scorerank, i = 4, j = 1:2, part = "body")
  scorerank <- merge_at(scorerank, i = 4, j = 3:4, part = "body")
  scorerank <- merge_at(scorerank, i = 4, j = 5:6, part = "body")
  scorerank <- merge_at(scorerank, i = 4, j = 7:8, part = "body")
  scorerank <- merge_at(scorerank, i = 6, j = 1:2, part = "body")
  scorerank <- merge_at(scorerank, i = 6, j = 3:4, part = "body")
  scorerank <- merge_at(scorerank, i = 6, j = 5:6, part = "body")
  scorerank <- merge_at(scorerank, i = 6, j = 7:8, part = "body")
  scorerank <- merge_at(scorerank, i = 8, j = 1:2, part = "body")
  scorerank <- merge_at(scorerank, i = 8, j = 3:4, part = "body")
  scorerank <- merge_at(scorerank, i = 8, j = 5:6, part = "body")
  scorerank <- merge_at(scorerank, i = 8, j = 7:8, part = "body")
  scorerank <- merge_at(scorerank, i = 7, j = 3:4, part = "body")
  scorerank <- merge_at(scorerank, i = 7, j = 5:6, part = "body")
  scorerank <- merge_at(scorerank, i = 9, j = 3:4, part = "body")
  scorerank <- merge_at(scorerank, i = 9, j = 5:6, part = "body")
  scorerank <- merge_at(scorerank, i = 6:9, j = 3:4, part = "body")
  scorerank <- merge_at(scorerank, i = 6:9, j = 5:6, part = "body")
  
  
  scorerank <- bold(scorerank, part = "all")
  # scorerank <- bold(scorerank, i= 1, part = "body")
  scorerank <- fontsize(scorerank, part = "all", size = 10)
  # dark blue as background color for header
  scorerank <-  bg(scorerank, bg = "#2B458F", part = "header")
  # slate as background color for body
  scorerank <-  bg(scorerank, bg = "#EFEDE5", part = "body")
  
  # dark blue as background color for Indicator column
  scorerank <-  bg(scorerank, i = 1, bg = "#2B458F", part = "body")
  
  # light blue color to highlight region and country
  # scorerank <-  bg(scorerank, j = regionName, bg = "#D2E0F2", part = "body")
  # scorerank <-  bg(scorerank, j = countries3iY4[i], bg = "#D2E0F2", part = "body")
  
  # scorerank <- autofit(scorerank)
  scorerank <- align(scorerank, align = "center", part = "all")
  scorerank <- color(scorerank, part = "header", color = "#FFFFFF")
  scorerank <- color(scorerank, i = 1, part = "body", color = "#FFFFFF")
  
  # scorerank <- colformat_num(
  #   x = scorerank,
  #   big.mark=",", digits = 1, na_str = "N/A")
  # scorerank <- colformat_num(
  #   x = scorerank,
  #   big.mark=",", i = 3, digits = 0, na_str = "N/A")
  # scorerank <- colformat_num(
  #   x = scorerank,
  #   big.mark=",", i = 4, digits = 0, na_str = "N/A")
  scorerank <- border_remove(x = scorerank)
  scorerank <- border_outer(scorerank, part="all", border = fp_border(color="white", width = 1) )
  scorerank <- border_inner_h(scorerank, part="all", border = fp_border(color="white", width = 1) )
  scorerank <- border_inner_v(scorerank, part="all", border = fp_border(color="white", width = 1) )
  
  
  scorerank <- padding(scorerank, padding=0, part = "all")
  
  scorerank <- height_all(scorerank, height = 0.38, part = "all")
  scorerank <- hrule(scorerank, rule = "exact")
  scorerank <- width(scorerank, width = .7)
  
  
  
  # scoreranks <- add_header_row(scorerank, values = "Index Scores and Ranks")
  # scorerank
  
  
  
  # Largest YoY change, need to exclude gender data; add caption via layout
  
  sign1 <- str_sub(CB_Y4$`Largest YoY Change1_change_exclude_survey_and_qual`[counter], 1,1)
  sign2 <- str_sub(CB_Y4$`Largest YoY Change2_change_exclude_survey_and_qual`[counter], 1,1)
  
  bigYOYdata <- data.table(Indicator = c(paste(sign1, " ", CB_Y4$`Largest YoY Change1_name_exclude_survey_and_qual`[counter]), paste(sign2, " ", CB_Y4$`Largest YoY Change2_name_exclude_survey_and_qual`[counter])),
                           Change = c(CB_Y4$`Largest YoY Change1_change_exclude_survey_and_qual`[counter], CB_Y4$`Largest YoY Change2_change_exclude_survey_and_qual`[counter]))
  
  if(sign1=="+"){
    sign1 <- "#00B050"
  } else if(sign1=="-"){
    sign1 <- "#FF0000"
  } else {
     sign1 <- "black"
  }
  
  if(sign2=="+"){
    sign2 <- "#00B050"
  } else if(sign2=="-"){
    sign2 <- "#FF0000"
  } else {
    sign2 <- "black"
  }
  
  # bigYOYdata
  
  bigYOY <- flextable(bigYOYdata)
  bigYOY <- width(bigYOY, j=1, width = 4.3)
  bigYOY <- width(bigYOY, j=2, width = 1.3)
  bigYOY <- height_all(bigYOY, height = 0.38, part = "all")
  bigYOY <- hrule(bigYOY, rule = "exact")
  
    
  
  bigYOY <- set_header_labels(bigYOY, Indicator="Indicator", Change="% Change")
  # bigYOY <- merge_at(bigYOY, i = 1, j = 1:2, part = "header")
  bigYOY <- align(bigYOY, align = "center", part = "header")
  bigYOY <- align(bigYOY, j=2, align = "center", part = "body")
  bigYOY <- color(bigYOY, part = "header", color = "#FFFFFF")
  bigYOY <- color(bigYOY, i=1, j=2, part = "body", color = sign1)
  bigYOY <- color(bigYOY, i=2, j=2, part = "body", color = sign2)
  
  bigYOY <- bold(bigYOY, part = "all")
  # bigYOY <- bold(bigYOY, i= 1, part = "body")
  bigYOY <- fontsize(bigYOY, part = "all", size = 10)
  # dark blue as background color for header
  bigYOY <-  bg(bigYOY, bg = "#2B458F", part = "header")
  # slate as background color for body
  bigYOY <-  bg(bigYOY, bg = "#EFEDE5", part = "body")
  
  bigYOY <- border_remove(x = bigYOY)
  bigYOY <- border_outer(bigYOY, part="all", border = fp_border(color="white", width = 1) )
  bigYOY <- border_inner_h(bigYOY, part="all", border = fp_border(color="white", width = 1) )
  bigYOY <- border_inner_v(bigYOY, part="all", border = fp_border(color="white", width = 1) )
  
  # bigYOY
  
  
  strengthsdata <- data.table(subcat = c(CB_Y4$highest_ranked_subcategory_name[counter], 
                                         CB_Y4$second_highest_ranked_subcategory_name[counter], 
                                         CB_Y4$third_highest_ranked_subcategory_name[counter]),
                              grank = c(CB_Y4$highest_ranked_subcategory_rank[counter], 
                                        CB_Y4$second_highest_ranked_subcategory_rank[counter], 
                                        CB_Y4$third_highest_ranked_subcategory_rank[counter]),
                              rrank = c(CB_Y4$highest_ranked_subcategory_regional_rank[counter], 
                                        CB_Y4$second_highest_ranked_subcategory_regional_rank[counter], 
                                        CB_Y4$third_highest_ranked_subcategory_regional_rank[counter]))
  
  # strengthsdata
  
  strengths <- flextable(strengthsdata)
  strengths <- width(strengths, j=1, width = 3.6)
  strengths <- width(strengths, j=2, width = 1)
  strengths <- width(strengths, j=3, width = 1)
  strengths <- height_all(strengths, height = 0.42, part = "all")
  strengths <- hrule(strengths, rule = "exact")
  # strengths
  
    
  if (CB_Y4$Region[counter]=="Asia"){
    shortregion <- "Asia"
  } else if (CB_Y4$Region[counter]=="Europe (EUR)"){
    shortregion <- "EUR"
  } else if (CB_Y4$Region[counter]=="Latin America (LATAM)"){
    shortregion <- "LATAM"
  } else if (CB_Y4$Region[counter]=="Middle East and North Africa (MENA)"){
    shortregion <- "MENA"
  } else if (CB_Y4$Region[counter]=="North America (NA)"){
    shortregion <- "NA"
  } else if (CB_Y4$Region[counter]=="Sub-Saharan Africa (SSA)"){
    shortregion <- "SSA"
  } 
  
  strengths <- set_header_labels(strengths, subcat = "Sub-Category", grank="Global Rank / 120", rrank= paste(shortregion, " Rank / ", CB_Y4$`Region_number of countries`[counter]))
  # strengths <- merge_at(strengths, i = 1, j = 1:2, part = "header")
  strengths <- align(strengths, align = "center", part = "header")
  strengths <- align(strengths, j=2, align = "center", part = "body")
  strengths <- align(strengths, j=3, align = "center", part = "body")
  strengths <- color(strengths, part = "header", color = "#FFFFFF")
  strengths <- bold(strengths, part = "all")
  # strengths <- bold(strengths, i= 1, part = "body")
  strengths <- fontsize(strengths, part = "all", size = 10)
  # dark blue as background color for header
  strengths <-  bg(strengths, bg = "#2B458F", part = "header")
  # slate as background color for body
  strengths <-  bg(strengths, bg = "#EFEDE5", part = "body")
  
  strengths <- border_remove(x = strengths)
  strengths <- border_outer(strengths, part="all", border = fp_border(color="white", width = 1) )
  strengths <- border_inner_h(strengths, part="all", border = fp_border(color="white", width = 1) )
  strengths <- border_inner_v(strengths, part="all", border = fp_border(color="white", width = 1) )
  
  # strengths
  
  
  
  
  weaknessesdata <- data.table(subcat = c(CB_Y4$`Overall_worst subcategory1`[counter], 
                                          CB_Y4$`Overall_worst subcategory2`[counter], 
                                          CB_Y4$`Overall_worst subcategory3`[counter]),
                              grank = c(CB_Y4$`Overall_worst subcategory1_rank`[counter], 
                                        CB_Y4$`Overall_worst subcategory2_rank`[counter], 
                                        CB_Y4$`Overall_worst subcategory3_rank`[counter]),
                              rrank = c(CB_Y4$`Overall_worst subcategory1_regional rank`[counter], 
                                        CB_Y4$`Overall_worst subcategory2_regional rank`[counter],
                                        CB_Y4$`Overall_worst subcategory3_regional rank`[counter]))
  
  # weaknessesdata
  
  weaknesses <- flextable(weaknessesdata)
  weaknesses <- width(weaknesses, j=1, width = 3.6)
  weaknesses <- width(weaknesses, j=2, width = 1)
  weaknesses <- width(weaknesses, j=3, width = 1)
  weaknesses <- height_all(weaknesses, height = 0.42, part = "all")
  weaknesses <- hrule(weaknesses, rule = "exact")
  # weaknesses
  
  weaknesses <- set_header_labels(weaknesses, subcat = "Sub-Category", grank="Global Rank / 120", rrank= paste(shortregion, " Rank / ", CB_Y4$`Region_number of countries`[counter]))
  # weaknesses <- merge_at(weaknesses, i = 1, j = 1:2, part = "header")
  weaknesses <- align(weaknesses, align = "center", part = "header")
  weaknesses <- align(weaknesses, j=2, align = "center", part = "body")
  weaknesses <- align(weaknesses, j=3, align = "center", part = "body")
  weaknesses <- color(weaknesses, part = "header", color = "#FFFFFF")
  weaknesses <- bold(weaknesses, part = "all")
  # weaknesses <- bold(weaknesses, i= 1, part = "body")
  weaknesses <- fontsize(weaknesses, part = "all", size = 10)
  # dark blue as background color for header
  weaknesses <-  bg(weaknesses, bg = "#2B458F", part = "header")
  # slate as background color for body
  weaknesses <-  bg(weaknesses, bg = "#EFEDE5", part = "body")
  
  weaknesses <- border_remove(x = weaknesses)
  weaknesses <- border_outer(weaknesses, part="all", border = fp_border(color="white", width = 1) )
  weaknesses <- border_inner_h(weaknesses, part="all", border = fp_border(color="white", width = 1) )
  weaknesses <- border_inner_v(weaknesses, part="all", border = fp_border(color="white", width = 1) )
  
  # weaknesses
  
  
  ggsign1 <- ""
  ggsign2 <- "-"
  
  if(CB_Y4$fmai[counter]>10){
    ggsign1 <- "+"
  } else if(CB_Y4$fmai[counter]< 0){
    ggsign1 <- "-"
  } else {
    ggsign1 <- ""
  }
  
  if(CB_Y4$PP_DIFF[counter]>0) {
      ggsign2 <- "+"
    } else if (CB_Y4$PP_DIFF[counter]<0){
      ggsign2 <- "-"
    } else {
      ggsign2 <- ""
    }
  
  
  
  genderdata <- data.table(Indicator = c("Gender Gap Ratio", paste(ggsign2, " ", "Percentage Point Change in Gender Gap Ratio")),
                           Value = c(CB_Y4$fmai[counter], CB_Y4$PP_DIFF[counter])) 
  
  if(CB_Y4$fmai[counter]>10){
    ggsign1 <- "#FF0000"
  } else if(CB_Y4$fmai[counter]<= 0){
    ggsign1 <- "#00B050"
  } else {
    ggsign1 <- "black"
  }
  
  if(ggsign2=="+"){
    ggsign2 <- "#FF0000"
  } else if(ggsign2=="-"){
    ggsign2 <- "#00B050"
  } else {
    ggsign2 <- "black"
  }
  
  # genderdata
  
  gender <- flextable(genderdata)
  gender <- width(gender, j=1, width = 4.3)
  gender <- width(gender, j=2, width = 1.3)
  gender <- height_all(gender, height = 0.38, part = "all")
  gender <- hrule(gender, rule = "exact")
  
  
  
  # gender <- set_header_labels(gender, Indicator="Indicator", Value="% Change")
  # gender <- merge_at(gender, i = 1, j = 1:2, part = "header")
  gender <- align(gender, align = "center", part = "header")
  gender <- align(gender, j=2, align = "center", part = "body")
  gender <- color(gender, part = "header", color = "#FFFFFF")
  gender <- color(gender, i=1, j=2, part = "body", color = ggsign1)
  gender <- color(gender, i=2, j=2, part = "body", color = ggsign2)
  gender <- bold(gender, part = "all")
  # gender <- bold(gender, i= 1, part = "body")
  gender <- fontsize(gender, part = "all", size = 10)
  # dark blue as background color for header
  gender <-  bg(gender, bg = "#2B458F", part = "header")
  # slate as background color for body
  gender <-  bg(gender, bg = "#EFEDE5", part = "body")
  
  gender <- border_remove(x = gender)
  gender <- border_outer(gender, part="all", border = fp_border(color="white", width = 1) )
  gender <- border_inner_h(gender, part="all", border = fp_border(color="white", width = 1) )
  gender <- border_inner_v(gender, part="all", border = fp_border(color="white", width = 1) )
  
  # gender
  
  
  
  yoyrankdata <- data.table(Category = c("", "Overall", "Availability", "Affordability", "Relevance", "Readiness"),
                        va = c("Rank", CB_Y4$`Overall rank_global`[counter],
                               CB_Y4$`Availability_global rank`[counter], 
                               CB_Y4$`Affordability_overall rank`[counter],
                               CB_Y4$`Relevance_overall rank`[counter],
                               CB_Y4$`Readiness_overall rank`[counter]),
                        vb = c("Change",CB_Y4$OVERALL_rank_change[counter],CB_Y4$AVAILABILITY_rank_change[counter],CB_Y4$AFFORDABILITY_rank_change[counter],CB_Y4$RELEVANCE_rank_change[counter],CB_Y4$READINESS_rank_change[counter]),
                        vc = c("Rank", 
                              CB_Y4$`Overall rank_regional`[counter], CB_Y4$`Availabilty_regional rank`[counter], CB_Y4$`Affordability_regional rank`[counter], CB_Y4$`Relevance_regional rank`[counter], CB_Y4$`Readiness_regional rank`[counter]),
                        vd = c("Change",CB_Y4$OVERALL_regional_rank_change[counter],CB_Y4$AVAILABILITY_regional_rank_change[counter],CB_Y4$AFFORDABILITY_regional_rank_change[counter],CB_Y4$RELEVANCE_regional_rank_change[counter],CB_Y4$READINESS_regional_rank_change[counter]))
  # yoyrankdata
  
  
  yoyrank <- flextable(yoyrankdata)
  
  yoyglobalsign <- str_sub(yoyrankdata$vb, 1,1)
  yoyglobalsign
  yoyglobalcolor <- c("black", "black", "black", "black", "black", "black")
  
  for(i in seq_along(yoyglobalsign)){
  if(yoyglobalsign[i]=="+"){
    yoyglobalcolor[i] <- "#00B050"
    }
    else if(yoyglobalsign[i]=="-"){
      yoyglobalcolor[i] <- "#FF0000"
    }
    else {
      yoyglobalcolor[i] <- "black"
    }
  }
  
  # yoyglobalcolor
  yoyglobalcolor[1] <- "white"
  
  
  yoyregionalsign <- str_sub(yoyrankdata$vd, 1,1)
  # yoyregionalsign
  yoyregionalcolor <- c("black", "black", "black", "black", "black", "black")
  
  for(i in seq_along(yoyregionalsign)){
    if(yoyregionalsign[i]=="+"){
      yoyregionalcolor[i] <- "#00B050"
    }
    else if(yoyregionalsign[i]=="-"){
      yoyregionalcolor[i] <- "#FF0000"
    }
    else {
      yoyregionalcolor[i] <- "black"
    }
  }
  
  # yoyregionalcolor
  yoyregionalcolor[1] <- "white"
  
  
  
  
  yoyrank <- set_header_labels(yoyrank, Category="Category ", va = "Global (120) ", vb = "Global (120) ", vc = paste0("", shortregion, " (", CB_Y4$`Region_number of countries`[counter],") "), vd = paste0("", shortregion, " (", CB_Y4$`Region_number of countries`[counter], ") "))
  
  yoyrank <- bold(yoyrank, part = "all")
  # yoyrank <- bold(yoyrank, i= 1, part = "body")
  yoyrank <- fontsize(yoyrank, part = "all", size = 10)
  # dark blue as background color for header
  yoyrank <-  bg(yoyrank, bg = "#2B458F", part = "header")
  # slate as background color for body
  yoyrank <-  bg(yoyrank, bg = "#EFEDE5", part = "body")
  # dark blue as background color for Indicator column
  yoyrank <-  bg(yoyrank, i = 1, bg = "#2B458F", part = "body")
  # yoyrank <- autofit(yoyrank)
  yoyrank <- align(yoyrank, align = "center", part = "all")
  yoyrank <- color(yoyrank, part = "header", color = "#FFFFFF")
  yoyrank <- color(yoyrank, i = 1, part = "body", color = "#FFFFFF")
  
  for(colorcode in seq_along(yoyglobalcolor)){
  yoyrank <- color(yoyrank, i = colorcode, j = 3, part = "body", color = yoyglobalcolor[colorcode])
  }
  for(colorcode in seq_along(yoyregionalcolor)){
    yoyrank <- color(yoyrank, i = colorcode, j = 5, part = "body", color = yoyregionalcolor[colorcode])
  }
  
  yoyrank <- border_remove(x = yoyrank)
  yoyrank <- border_outer(yoyrank, part="all", border = fp_border(color="white", width = 1) )
  yoyrank <- border_inner_h(yoyrank, part="all", border = fp_border(color="white", width = 1) )
  yoyrank <- border_inner_v(yoyrank, part="all", border = fp_border(color="white", width = 1) )
  
  
  yoyrank <- merge_at(yoyrank, i = 1, j = 2:3, part = "header")
  yoyrank <- merge_at(yoyrank, i = 1, j = 4:5, part = "header")
  
  yoyrank <- padding(yoyrank, padding=0, part = "all")
  
  yoyrank <- height_all(yoyrank, height = 0.62, part = "all")
  yoyrank <- height_all(yoyrank, height = 0.62, part = "header")
  yoyrank <- hrule(yoyrank, rule = "exact")
  yoyrank <- width(yoyrank, width = 1.12)
  yoyrank <- width(yoyrank, j=1, width = 2)
  yoyrank <- width(yoyrank, j=2, width = .9)
  yoyrank <- width(yoyrank, j=3, width = .9)
  yoyrank <- width(yoyrank, j=4, width = .9)
  yoyrank <- width(yoyrank, j=5, width = .9)
  # yoyrank
  
  # ggplot(CB_Y4, aes(x=CB_Y4$`1) AVAILABILITY`, fill="blue")) +
  #   geom_density(alpha=.5) +
  #   theme(
  #     legend.position="bottom",
  #     panel.spacing = unit(0.1, "lines"),
  #     axis.ticks.x=element_blank()
  #   )
  # 
  # ggplot(CB_Y4, aes(x=CB_Y4$`2) AFFORDABILITY`, fill="blue")) +
  #   geom_density(alpha=.5) +
  #   theme(
  #     legend.position="bottom",
  #     panel.spacing = unit(0.1, "lines"),
  #     axis.ticks.x=element_blank()
  #   )
  # 
  # ggplot(CB_Y4, aes(x=CB_Y4$`3) RELEVANCE`, fill="blue")) +
  #   geom_density(alpha=.5) +
  #   theme(
  #     legend.position="bottom",
  #     panel.spacing = unit(0.1, "lines"),
  #     axis.ticks.x=element_blank()
  #   )
  # 
  # ggplot(CB_Y4, aes(x=CB_Y4$`4) READINESS`, fill="blue")) +
  #   geom_density(alpha=.5) +
  #   theme(
  #     legend.position="bottom",
  #     panel.spacing = unit(0.1, "lines"),
  #     axis.ticks.x=element_blank()
  #   )
  # 
  # av_plot <- ggplot() +
  #   scale_fill_manual(values = c("#9a9597", "#ffb237"))+
  #   geom_col(data=CB_Y4[CB_Y4$Region==CB_Y4$Region[1],], aes(x=reorder(Country, `1) AVAILABILITY`), y=`1) AVAILABILITY`))+theme_classic() + theme(axis.title.x = element_text(face = "bold", size=7), axis.text.x = element_text(face="plain",size=6),axis.title.y = element_blank(), axis.text.y = element_text(size=6), panel.background = element_rect(fill = "#EFEDE5", colour = "#EFEDE5"))+
  #   theme(legend.position="none")+labs(y= "Availability")
  # av_plot
  
  
  GGRBP <- data.frame(AGGR = c("Female Users", "Male Users"), BGGR = c(CB_Y4$`BG21) Female internet users_E5`[counter], CB_Y4$`BG20) Male internet users_E5`[counter]))
  
  d=data.frame(x=c(1.2,1.8,1.8), 
               y=c(CB_Y4$`BG21) Female internet users_E5`[counter],CB_Y4$`BG21) Female internet users_E5`[counter],CB_Y4$`BG20) Male internet users_E5`[counter]))
  
  sq=data.frame(x1=2.2, x2=3, y1=CB_Y4$`BG21) Female internet users_E5`[counter], y2=CB_Y4$`BG20) Male internet users_E5`[counter])
    
    
  
  
  
  
               
  GGRbarplot <- ggplot(GGRBP, aes(x=AGGR, y=BGGR)) +
    geom_bar(stat="identity", colour = "black", width = .4, position = "dodge") +
    geom_text(aes(label = paste0(format(BGGR, nsmall=0), "%")), 
              position = position_dodge(width = .4), vjust=1.5,
              colour = "white", fontface="bold", size = 3) +
    theme(axis.title.y = element_text(face = "italic"), axis.text.x = element_text(face = "italic"), panel.background = element_rect(fill = "#EFEDE5", colour = "#EFEDE5", size = 0.5), panel.grid.major = element_blank(), panel.grid.minor = element_blank(), legend.position="none", legend.title = element_text(size=10, face="bold"), legend.text=element_text(size=10, face="plain"))+labs(y= "% of population with internet access", x = "")+scale_x_discrete(labels = c('Female\nInternet Users','Male\nInternet Users'))+ ylim(0, 100)
    
  
  if(CB_Y4$`1.1.4) Gender gap in internet access_E5`[counter]>5){
    GGRbarplot <- GGRbarplot+
      geom_segment(aes(x=1.2, xend=1.8, y=CB_Y4$`BG21) Female internet users_E5`[counter], yend=CB_Y4$`BG20) Male internet users_E5`[counter]))+
      geom_segment(aes(x=1.2, xend=1.8, y=CB_Y4$`BG21) Female internet users_E5`[counter], yend=CB_Y4$`BG21) Female internet users_E5`[counter]))+
      geom_segment(aes(x=2.2, xend=3, y=CB_Y4$`BG20) Male internet users_E5`[counter]), yend=CB_Y4$`BG20) Male internet users_E5`[counter])+
      geom_segment(aes(x=1.8, xend=2.2, y=CB_Y4$`BG21) Female internet users_E5`[counter], yend=CB_Y4$`BG21) Female internet users_E5`[counter]), linetype=2)+geom_segment(aes(x=2.2, xend=3, y=CB_Y4$`BG21) Female internet users_E5`[counter], yend=CB_Y4$`BG21) Female internet users_E5`[counter]))+
      geom_polygon(data=d, inherit.aes = F,mapping=aes(x=x, y=y, fill=CB_Y4$fmai[counter]))+
    geom_rect(data=sq, inherit.aes = F, mapping=aes(xmin=x1, xmax=x2, ymin=y1, ymax=y2, fill=CB_Y4$fmai[counter]))+scale_fill_manual(values = c("#9a9597", "#ffffff"))+
      geom_text(x=2.6, y=CB_Y4$`BG21) Female internet users_E5`[counter]-1, label=paste0("Internet gender gap:\n",CB_Y4$fmai[counter]), col = "black", hjust = "center", vjust = "top",  fontface = "bold", size = 3)+scale_fill_binned(limits = c(-14, 71), breaks = c(0,10, 20, 30, 40, 50, 60, 70), low = "yellow", high = "darkred")
  }
  GGRbarplot
  
  fbwidth <- 5
  
  mbwidth <- 5
  
  mobGGRBP <- data.frame(AGGR = c("Female Users", "Male Users"), BGGR = c(CB_Y4$`BG23) Female mobile phone subscribers_E5`[counter], CB_Y4$`BG22) Male mobile phone subscribers_E5`[counter]))
  
  mobd=data.frame(x=c(1.2,1.8,1.8), 
               y=c(CB_Y4$`BG23) Female mobile phone subscribers_E5`[counter],CB_Y4$`BG23) Female mobile phone subscribers_E5`[counter],CB_Y4$`BG22) Male mobile phone subscribers_E5`[counter]))
  
  mobtraffic <- CB_Y4$`BG23) Female mobile phone subscribers_E5`[counter]>10
  
  mobGGRbarplot <- ggplot(mobGGRBP, aes(x=AGGR, y=BGGR)) +
    geom_bar(stat="identity", colour = "black", width = .4, position = "dodge") +
    geom_text(aes(label = paste0(format(BGGR, nsmall=0), "%")), 
              position = position_dodge(width = .4), vjust=1.5,
              colour = "white", fontface="bold", size = 3) +
    theme(axis.title.y = element_text(face = "italic"), axis.text.x = element_text(face = "italic"), panel.background = element_rect(fill = "#EFEDE5", colour = "#EFEDE5", size = 0.5), panel.grid.major = element_blank(), panel.grid.minor = element_blank(), legend.position="none", legend.title = element_text(size=10, face="bold"), legend.text=element_text(size=10, face="plain"))+labs(y= "% of population with mobile access", x = "")+ 
    scale_x_discrete(labels = c('Female\nMobile Users','Male\nMobile Users'))+ ylim(0, 100)
  
  if(CB_Y4$`1.1.5) Gender gap in mobile phone access_E5`[counter]>5){
    mobGGRbarplot <- mobGGRbarplot +
      geom_segment(aes(x=1.2, xend=1.8, y=CB_Y4$`BG23) Female mobile phone subscribers_E5`[counter], yend=CB_Y4$`BG22) Male mobile phone subscribers_E5`[counter]))+
      geom_segment(aes(x=1.2, xend=1.8, y=CB_Y4$`BG23) Female mobile phone subscribers_E5`[counter], yend=CB_Y4$`BG23) Female mobile phone subscribers_E5`[counter]))+geom_segment(aes(x=2.2, xend=3, y=CB_Y4$`BG22) Male mobile phone subscribers_E5`[counter]), yend=CB_Y4$`BG22) Male mobile phone subscribers_E5`[counter])+
      geom_segment(aes(x=1.8, xend=2.2, y=CB_Y4$`BG23) Female mobile phone subscribers_E5`[counter], yend=CB_Y4$`BG23) Female mobile phone subscribers_E5`[counter]), linetype=2)+
      geom_segment(aes(x=2.2, xend=3, y=CB_Y4$`BG23) Female mobile phone subscribers_E5`[counter], yend=CB_Y4$`BG23) Female mobile phone subscribers_E5`[counter]))+
      geom_polygon(data=mobd, mapping=aes(x=x, y=y, fill=CB_Y4$`1.1.5) Gender gap in mobile phone access_E5`[counter]))+
      geom_rect(mapping=aes(xmin=2.2, xmax=3, ymin=CB_Y4$`BG23) Female mobile phone subscribers_E5`[counter], ymax=CB_Y4$`BG22) Male mobile phone subscribers_E5`[counter], fill=CB_Y4$`1.1.5) Gender gap in mobile phone access_E5`[counter]))+
      geom_text(x=2.6, y=CB_Y4$`BG23) Female mobile phone subscribers_E5`[counter]-1, label=paste0("Mobile gender gap:\n",CB_Y4$`1.1.5) Gender gap in mobile phone access_E5`[counter]), col = "black", hjust = "center", vjust = "top",  fontface = "bold", size = 3)+scale_fill_binned(limits = c(-14, 71), breaks = c(0,10, 20, 30, 40, 50, 60, 70), low = "yellow", high = "darkred") + theme(legend.position = "right", legend.key.height = unit(1.0, "cm"), legend.title = element_text(size=10, face="bold"))+labs(fill="Gender\nGap")
    mbwidth <- 5.8
  }
  
  if(CB_Y4$`1.1.5) Gender gap in mobile phone access_E5`[counter]<=5 & CB_Y4$`1.1.4) Gender gap in internet access_E5`[counter]>5){
    GGRbarplot  <- GGRbarplot  + theme(legend.position = "right", legend.key.height = unit(1.0, "cm"), legend.title = element_text(size=10, face="bold"))+labs(fill="Gender\nGap")
    fbwidth <- 5.8
  }
  
  
  
  
  mobGGRbarplot
  
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
  
  
  CB_Y4$Region_short <- CB_Y4$Region
  CB_Y4$Region_short[CB_Y4$Region_short=="Asia"] <- "ASIA"
  CB_Y4$Region_short[CB_Y4$Region_short=="Europe (EUR)"] <- "EUR"
  CB_Y4$Region_short[CB_Y4$Region_short=="Latin America (LATAM)"] <- "LATAM"
  CB_Y4$Region_short[CB_Y4$Region_short=="Middle East and North Africa (MENA)"] <- "MENA"
  CB_Y4$Region_short[CB_Y4$Region_short=="North America (NA)"] <- "NA"
  CB_Y4$Region_short[CB_Y4$Region_short=="Sub-Saharan Africa (SSA)"] <- "SSA"
  
  
  # completed 041021
  #Fixed Broadband Charts
  #Upload Speed
  UPLyears <- c("2016", "2017", "2018", "2019", "2020")
  UPLvalue <- c(CB_Y4$`1.2.1) Average fixed broadband upload speed_E1`[counter], CB_Y4$`1.2.1) Average fixed broadband upload speed_E2`[counter], CB_Y4$`1.2.1) Average fixed broadband upload speed_E3`[counter], CB_Y4$`1.2.1) Average fixed broadband upload speed_E4.y`[counter], CB_Y4$`1.2.1) Average fixed broadband upload speed_E5.y`[counter])
  reg_UPLvalue <- c(CB_Y4$`1.2.1) FB UL_AE1`[counter], CB_Y4$`1.2.1) FB UL_AE2`[counter], CB_Y4$`1.2.1) FB UL_AE3`[counter], CB_Y4$`1.2.1) FB UL_AE4`[counter], CB_Y4$`1.2.1) FB UL_AE5`[counter])
  UPLyearsdata <- data.frame(UPLyears,UPLvalue, reg_UPLvalue)
  UPLyearsdata$Region <- CB_Y4$Region[counter]
  
  #FB Download Speed
  DLyears <- c("2016", "2017", "2018", "2019", "2020")
  DLvalue <- c(CB_Y4$`1.2.2) Average fixed broadband download speed_E1`[counter], CB_Y4$`1.2.2) Average fixed broadband download speed_E2`[counter], CB_Y4$`1.2.2) Average fixed broadband download speed_E3`[counter], CB_Y4$`1.2.2) Average fixed broadband download speed_E4.y`[counter], CB_Y4$`1.2.2) Average fixed broadband download speed_E5.y`[counter])
  reg_DLvalue <- c(CB_Y4$`1.2.2) FB DL_AE1`[counter], CB_Y4$`1.2.2) FB DL_AE2`[counter], CB_Y4$`1.2.2) FB DL_AE3`[counter], CB_Y4$`1.2.2) FB DL_AE4`[counter], CB_Y4$`1.2.2) FB DL_AE5`[counter])
  DLyearsdata <- data.frame(DLyears,DLvalue, reg_DLvalue)
  DLyearsdata$Region <- CB_Y4$Region_short[counter]
  
  
  DLmax <- max(select(DLyearsdata, 2:3))+5
  UPLmax <- max(select(UPLyearsdata, 2:3))+5
  DLmax <- max(UPLmax, DLmax)
  
  # Grouped
  ULbar <- ggplot(UPLyearsdata) + 
    geom_bar(aes(fill=UPLyears, y=UPLvalue, x=UPLyears), position="dodge", stat="identity")+ylab("Mbps")+geom_text(x=UPLyears, y=UPLvalue- 0.5, label = paste(round(UPLvalue,1)), vjust = "top", fontface = "bold", col = "white")+labs(title="Fixed Upload Speeds")+theme_economist()+theme(legend.position = "none")+theme(axis.text = element_text(size = 10, face = "bold.italic"))+theme(title = element_text(size = 8, color = "black", face = "bold"), axis.title.y = element_text(size = 10, face = "bold.italic"))+geom_line(aes(y=reg_UPLvalue, x=UPLyears, group =1, color = Region), alpha = 0.8, size = 1.5)+geom_point(aes(y=reg_UPLvalue, x=UPLyears, group =1), alpha = 0.5,size = 3, color = "red")+scale_fill_economist()+theme(axis.line.x = element_blank(), axis.title.x = element_blank(), axis.ticks.x = element_blank(), axis.text.x = element_blank(), legend.text=element_text(size=10))+labs(fill = "Year", color = "Regional\nAverage")+ylim(0, DLmax)
  
  ULbar
  
    
  
  
  # Grouped
  DLbar <- ggplot(DLyearsdata) + 
    geom_bar(aes(fill=DLyears, y=DLvalue, x=DLyears), position="dodge", stat="identity")+ylab("Mbps")+geom_text(x=DLyears, y=DLvalue- 0.5, label = paste(round(DLvalue,1)), vjust = "top", fontface = "bold", col = "white")+labs(title="Fixed Download Speeds")+theme_economist()+theme(legend.position = "right")+theme(axis.text = element_text(size = 10, face = "bold.italic"))+theme(title = element_text(size = 8, color = "black", face = "bold"), axis.title.y = element_blank())+geom_line(aes(y=reg_DLvalue, x=DLyears, group =1, color = Region), alpha = 0.8,size = 1.5)+geom_point(aes(y=reg_DLvalue, x=DLyears, group =1), alpha = 0.5,size = 3, color = "red")+scale_fill_economist()+theme(axis.line.x = element_blank(), axis.title.x = element_blank(), axis.ticks.x = element_blank(), axis.text.x = element_blank(), axis.ticks.y = element_blank(), axis.text.y = element_blank(), legend.text=element_text(size=10))+labs(fill = "Year", color = "Regional\nAverage")+ylim(0, DLmax)
  
  DLbar
  
  
  #Mobile Charts
  MUPLyears <- c("2016", "2017", "2018", "2019", "2020")
  MUPLvalue <- c(CB_Y4$`1.2.4) Average mobile upload speed_E1`[counter], CB_Y4$`1.2.4) Average mobile upload speed_E2`[counter], CB_Y4$`1.2.4) Average mobile upload speed_E3`[counter], CB_Y4$`1.2.4) Average mobile upload speed_E4.y`[counter], CB_Y4$`1.2.4) Average mobile upload speed_E5.y`[counter])
  reg_MUPLvalue <- c(CB_Y4$`1.2.4) M UL_AE1`[counter], CB_Y4$`1.2.4) M UL_AE2`[counter], CB_Y4$`1.2.4) M UL_AE3`[counter], CB_Y4$`1.2.4) M UL_AE4`[counter], CB_Y4$`1.2.4) M UL_AE5`[counter])
  MUPLyearsdata <- data.frame(MUPLyears,MUPLvalue, reg_MUPLvalue)
  MUPLyearsdata$Region <- CB_Y4$Region[counter]
  
  # create a dataset
  MDLyears <- c("2016", "2017", "2018", "2019", "2020")
  MDLvalue <- c(CB_Y4$`1.2.5) Average mobile download speed_E1`[counter], CB_Y4$`1.2.5) Average mobile download speed_E2`[counter], CB_Y4$`1.2.5) Average mobile download speed_E3`[counter], CB_Y4$`1.2.5) Average mobile download speed_E4.y`[counter], CB_Y4$`1.2.5) Average mobile download speed_E5.y`[counter])
  reg_MDLvalue <- c(CB_Y4$`1.2.5) M DL_AE1`[counter], CB_Y4$`1.2.5) M DL_AE2`[counter], CB_Y4$`1.2.5) M DL_AE3`[counter], CB_Y4$`1.2.5) M DL_AE4`[counter], CB_Y4$`1.2.5) M DL_AE5`[counter])
  MDLyearsdata <- data.frame(MDLyears,MDLvalue, reg_MDLvalue)
  MDLyearsdata$Region <- CB_Y4$Region_short[counter]
  
  MDLmax <- max(select(MDLyearsdata, 2:3))+5
  MUPLmax <- max(select(MUPLyearsdata, 2:3))+5
  MDLmax <- max(MUPLmax, MDLmax)
  
  
  # Grouped
  MULbar <- ggplot(MUPLyearsdata) + 
    geom_bar(aes(fill=MUPLyears, y=MUPLvalue, x=MUPLyears), position="dodge", stat="identity")+ylab("Mbps")+geom_text(x=MUPLyears, y=MUPLvalue- 0.5, label = paste(round(MUPLvalue,1)), vjust = "top", fontface = "bold", col = "white")+labs(title="Mobile Upload Speeds")+theme_economist()+theme(legend.position = "none")+theme(title = element_text(size = 8, color = "black", face = "bold"), axis.title.x = element_text(face = "bold.italic"), axis.title.y = element_text(size = 10, face = "bold.italic"))+geom_line(aes(y=reg_MUPLvalue, x=MUPLyears, group =1, color = Region), alpha = 0.8, size = 1.5)+geom_point(aes(y=reg_MUPLvalue, x=MUPLyears, group =1), alpha = 0.5,size = 3, color = "red")+scale_fill_economist()+theme(axis.line.x = element_blank(), axis.title.x = element_blank(), axis.ticks.x = element_blank(), axis.text.x = element_blank(), legend.text=element_text(size=10))+labs(fill = "Year", color = "Regional\nAverage")+ylim(0, MDLmax)
  
  MULbar
  
  
  # Grouped
  MDLbar <- ggplot(MDLyearsdata) + 
    geom_bar(aes(fill=MDLyears, y=MDLvalue, x=MDLyears), position="dodge", stat="identity")+ylab("Mbps")+geom_text(x=MDLyears, y=MDLvalue- 0.5, label = paste(round(MDLvalue,1)), vjust = "top", fontface = "bold", col = "white")+labs(title="Mobile Download Speeds")+theme_economist()+theme(legend.position = "right")+theme(axis.text = element_text(size = 8, face = "bold.italic"))+theme(title = element_text(size = 8, color = "black", face = "bold"), axis.title.y = element_blank())+geom_line(aes(y=reg_MDLvalue, x=MDLyears, group =1, color = Region), alpha = 0.8, size = 1.5)+geom_point(aes(y=reg_MDLvalue, x=MDLyears, group =1), alpha = 0.5,size = 3, color = "red")+scale_fill_economist()+theme(axis.line.x = element_blank(), axis.title.x = element_blank(), axis.ticks.x = element_blank(), axis.text.x = element_blank(), axis.ticks.y = element_blank(), axis.text.y = element_blank(), legend.text=element_text(size=10))+labs(fill = "Year", color = "Regional\nAverage")+ylim(0, MDLmax)
  
  MDLbar
  
  
  
  # dl_bp <- ggplot(CB_Y4[CB_Y4$Region==CB_Y4$Region[counter],], aes(x=1, y=`1.2.2) Average fixed broadband download speed_E5.y`)) + 
  #   geom_boxplot()+geom_jitter(position = position_jitter(seed = 1), color="black", size=0.4)+theme_economist()+
  #   theme(legend.position="right", legend.title = element_text(size=10, face="bold"), legend.text=element_text(size=10, face="plain"))+
  #   xlab("")+ylab("Mbps\n")+labs(title="Fixed Download Speeds")+
  #   geom_text_repel(position = position_jitter(seed = 1), aes(x=1, y=`1.2.2) Average fixed broadband download speed_E5.y`, label=Country))+scale_color_economist()+
  #   theme(axis.text.x = element_blank(), axis.line.x = element_blank(), axis.ticks.x = element_blank(), axis.title.x = element_text(face="bold"), axis.title.y = element_text(face="bold"))
  # 
  # dl_bp
  
  obc_latyears <- c("2016", "2017", "2018", "2019", "2020")
  obc_latvalue <- c(CB_Y4$`1.2.3) Average fixed broadband latency_E1`[counter], CB_Y4$`1.2.3) Average fixed broadband latency_E2`[counter], CB_Y4$`1.2.3) Average fixed broadband latency_E3`[counter], CB_Y4$`1.2.3) Average fixed broadband latency_E4.y`[counter], CB_Y4$`1.2.3) Average fixed broadband latency_E5.y`[counter])
  obc_reg_latvalue <- c(CB_Y4$`1.2.3) FB L_AE1`[counter], CB_Y4$`1.2.3) FB L_AE2`[counter], CB_Y4$`1.2.3) FB L_AE3`[counter], CB_Y4$`1.2.3) FB L_AE4`[counter], CB_Y4$`1.2.3) FB L_AE5`[counter])
  
  obc_latyearsdata <- data.frame(obc_latyears,obc_latvalue, obc_reg_latvalue)
  obc_latyearsdata$Region <- CB_Y4$Region[counter]
  
  obc_latency <- ggplot(obc_latyearsdata)+
    geom_bar(aes(fill=obc_latyears, y=obc_latvalue, x=obc_latyears), position="dodge", stat="identity")+
    coord_flip()+
    theme_economist()+
    geom_text(aes(x=obc_latyears, y=obc_latvalue- 1, label = paste(round(obc_latvalue,1))), vjust = "center", hjust = "left", fontface = "bold", col = "white")+
    labs(title="Fixed Latency (ms)")+
    guides(colour = guide_legend(reverse = TRUE))+
    theme(legend.position = "none")+
    theme(axis.text = element_text(size = 8, face = "bold.italic"))+
    theme(title = element_text(size = 8, color = "black", face = "bold"))+
    # geom_line(aes(y=obc_reg_latvalue, x=obc_latyears, group =1, color = Region), size = 1.5)+
    # geom_point(aes(y=obc_reg_latvalue, x=obc_latyears, group =1), size = 3, color = "red")+
    scale_fill_economist()+
    theme(panel.grid.major = element_blank(), panel.grid.minor = element_blank(), axis.title.x = element_blank(), axis.title.y = element_blank(), axis.line.x = element_blank(), axis.ticks.x = element_blank(), axis.ticks.y = element_blank(), axis.text.y = element_blank(), legend.text=element_text(size=10), plot.margin = unit(c(0,0,0.5,1.3), "cm"))+
    labs(fill = "Year", color = "Regional\nAverage")+
    geom_text(aes(x=obc_latyears, y=1, label = paste(obc_latyears)), vjust = "center", hjust = "right", size = 4, fontface = "bold.italic", col = "white")+
    scale_y_reverse()+
    geom_segment(aes(x = 0.5, xend = 1.5, y = CB_Y4$`1.2.3) FB L_AE1`[counter], yend = CB_Y4$`1.2.3) FB L_AE1`[counter], alpha = 0.8,color = "red")) +
    geom_segment(aes(x = 1.5, xend = 2.5, y = CB_Y4$`1.2.3) FB L_AE2`[counter], yend = CB_Y4$`1.2.3) FB L_AE2`[counter], alpha = 0.8,color = "red")) +
    geom_segment(aes(x = 2.5, xend = 3.5, y = CB_Y4$`1.2.3) FB L_AE3`[counter], yend = CB_Y4$`1.2.3) FB L_AE3`[counter], alpha = 0.8,color = "red")) +
    geom_segment(aes(x = 3.5, xend = 4.5, y = CB_Y4$`1.2.3) FB L_AE4`[counter], yend = CB_Y4$`1.2.3) FB L_AE4`[counter], alpha = 0.8,color = "red")) +
    geom_segment(aes(x = 4.5, xend = 5.5, y = CB_Y4$`1.2.3) FB L_AE5`[counter], yend = CB_Y4$`1.2.3) FB L_AE5`[counter], alpha = 0.8,color = "red")) 
  
  obc_latency
  
  mobc_latyears <- c("2016", "2017", "2018", "2019", "2020")
  mobc_latvalue <- c(CB_Y4$`1.2.6) Average mobile latency_E1`[counter], CB_Y4$`1.2.6) Average mobile latency_E2`[counter], CB_Y4$`1.2.6) Average mobile latency_E3`[counter], CB_Y4$`1.2.6) Average mobile latency_E4.y`[counter], CB_Y4$`1.2.6) Average mobile latency_E5.y`[counter])
  mobc_reg_latvalue <- c(CB_Y4$`1.2.6) M L_AE1`[counter], CB_Y4$`1.2.6) M L_AE2`[counter], CB_Y4$`1.2.6) M L_AE3`[counter], CB_Y4$`1.2.6) M L_AE4`[counter], CB_Y4$`1.2.6) M L_AE5`[counter])
  
  mobc_latyearsdata <- data.frame(obc_latyears,obc_latvalue, obc_reg_latvalue)
  mobc_latyearsdata$Region <- CB_Y4$Region[counter]
  
  
  
  mobc_latency <- ggplot(mobc_latyearsdata)+
    geom_bar(aes(fill=mobc_latyears, y=mobc_latvalue, x=mobc_latyears), position="dodge", stat="identity")+
    coord_flip()+
    theme_economist()+
    geom_text(aes(x=mobc_latyears, y=mobc_latvalue- 2, label = paste(round(mobc_latvalue,1))), vjust = "center", hjust = "left", fontface = "bold", col = "white")+
    labs(title="Mobile Latency (ms)")+
    guides(colour = guide_legend(reverse = TRUE))+
    theme(legend.position = "none")+
    theme(axis.text = element_text(size = 8, face = "bold.italic"))+
    theme(title = element_text(size = 8, color = "black", face = "bold"))+
    # geom_line(aes(y=mobc_reg_latvalue, x=mobc_latyears, group =1, color = Region), size = 1.5)+
    # geom_point(aes(y=mobc_reg_latvalue, x=mobc_latyears, group =1), size = 3, color = "red")+
    scale_fill_economist()+
    theme(panel.grid.major = element_blank(), panel.grid.minor = element_blank(), axis.title.x = element_blank(), axis.title.y = element_blank(), axis.line.x = element_blank(), axis.ticks.x = element_blank(),axis.ticks.y = element_blank(), axis.text.y = element_blank(), legend.text=element_text(size=10), plot.margin = unit(c(0,0,0.5,1.3), "cm"))+
    labs(fill = "Year", color = "Regional\nAverage")+
    geom_text(aes(x=mobc_latyears, y=1, label = paste(mobc_latyears)), vjust = "center", hjust = "right", size = 4, fontface = "bold.italic", col = "white")+
    scale_y_reverse()+
    geom_segment(aes(x = 0.5, xend = 1.5, y = CB_Y4$`1.2.6) M L_AE1`[counter], yend = CB_Y4$`1.2.6) M L_AE1`[counter], alpha = 0.8,color = "red")) +
    geom_segment(aes(x = 1.5, xend = 2.5, y = CB_Y4$`1.2.6) M L_AE2`[counter], yend = CB_Y4$`1.2.6) M L_AE2`[counter], alpha = 0.8,color = "red")) +
    geom_segment(aes(x = 2.5, xend = 3.5, y = CB_Y4$`1.2.6) M L_AE3`[counter], yend = CB_Y4$`1.2.6) M L_AE3`[counter], alpha = 0.8,color = "red")) +
    geom_segment(aes(x = 3.5, xend = 4.5, y = CB_Y4$`1.2.6) M L_AE4`[counter], yend = CB_Y4$`1.2.6) M L_AE4`[counter], alpha = 0.8,color = "red")) +
    geom_segment(aes(x = 4.5, xend = 5.5, y = CB_Y4$`1.2.6) M L_AE5`[counter], yend = CB_Y4$`1.2.6) M L_AE5`[counter], alpha = 0.8,color = "red"))
    
  
  mobc_latency
  
  
  temp_CB_Y4 <- CB_Y4
  temp_CB_Y4[temp_CB_Y4=="Asia"] <- "ASIA"
  temp_CB_Y4[temp_CB_Y4=="Europe (EUR)"] <- "EUR"
  temp_CB_Y4[temp_CB_Y4=="Latin America (LATAM)"] <- "LATAM"
  temp_CB_Y4[temp_CB_Y4=="Middle East and North Africa (MENA)"] <- "MENA"
  temp_CB_Y4[temp_CB_Y4=="North America (NA)"] <- "NA"
  temp_CB_Y4[temp_CB_Y4=="Sub-Saharan Africa (SSA)"] <- "SSA"
  
  
  dl_dist <- ggplot(temp_CB_Y4) +
    geom_density(aes(x=`1.2.2) Average fixed broadband download speed_E5.y`, fill = Region),alpha=0.8)+facet_grid(rows = vars(Region))+theme(legend.position = "none")+scale_fill_economist()+theme(title = element_text(size = 10, color = "black", face = "bold"))+xlab("Mbps")+theme(axis.title.y = element_blank())+
    labs(title="Distribution of Fixed Download Speeds")
  
  dl_dist
  
  mdl_dist <- ggplot(temp_CB_Y4) +
    geom_density(aes(x=`1.2.5) Average mobile download speed_E5.y`, fill = Region),alpha=0.8)+facet_grid(rows = vars(Region))+theme(legend.position = "none")+scale_fill_economist()+theme(title = element_text(size = 10, color = "black", face = "bold"))+xlab("Mbps")+theme(axis.title.y = element_blank())+
    labs(title="Distribution of Mobile Download Speeds")
  
  mdl_dist
  
  temp_CB_Y4 <- CB_Y4
  temp_CB_Y4$dlplotfill <- temp_CB_Y4$Country==CB_Y4$Country[counter]
  if(CB_Y4$Region[counter]!="North America (NA)"){
    temp_CB_Y4 <- temp_CB_Y4[temp_CB_Y4$Region==CB_Y4$Region[counter],]  
  } else{
    temp_CB_Y4 <- temp_CB_Y4[temp_CB_Y4$`Income group`==CB_Y4$`Income group`[counter],]
  }
  
  
  
  dlorderedbar <- ggplot() +
    scale_fill_manual(values = c("#9a9597", "red"))+
    geom_col(data=temp_CB_Y4, aes(fill=temp_CB_Y4$dlplotfill, x=reorder(temp_CB_Y4$Country, -temp_CB_Y4$`1.2.2) Average fixed broadband download speed_E5.y`), y=temp_CB_Y4$`1.2.2) Average fixed broadband download speed_E5.y`))+coord_flip()+theme_classic() +theme(axis.title.x = element_text(face = "plain", size=10), axis.text.x = element_text(face="plain",size=10),axis.title.y = element_blank(), axis.text.y = element_text(size=9))+
    theme(legend.position="none")+labs(y= "Mbps")+
    theme(title = element_text(size = 10, color = "black", face = "bold"))+
    labs(title = "Fixed Download Speeds")
  dlorderedbar
  
  mdlorderedbar <- ggplot() +
    scale_fill_manual(values = c("#9a9597", "red"))+
    geom_col(data=temp_CB_Y4, aes(fill=temp_CB_Y4$dlplotfill, x=reorder(temp_CB_Y4$Country, -temp_CB_Y4$`1.2.5) Average mobile download speed_E5.y`), y=temp_CB_Y4$`1.2.5) Average mobile download speed_E5.y`))+coord_flip()+theme_classic() +theme(axis.title.x = element_text(face = "plain", size=9), axis.text.x = element_text(face="plain",size=10),axis.title.y = element_blank(), axis.text.y = element_text(size=10))+
    theme(legend.position="none")+labs(y= "Mbps")+
    theme(title = element_text(size = 10, color = "black", face = "bold"))+
    labs(title = "Mobile Download Speeds")
  mdlorderedbar
  
  
  
  
  affplot <- ggplot(Y4_data, 
                    aes(x=inth, y = `2.1.2`)) + 
    geom_point(position="jitter", 
               mapping = aes(color=`Income Group`))+
    geom_segment(x=-20, xend=110, y=2, yend=2, linetype = "dotted")+
    xlab("Household Connectivity (% Population)")+ylab("Price of 1Gb Prepaid data (% of monthly GNI per capita)")+labs(color="Country Income Groups", size="Gender Gap Ratio")+
    theme_classic() +
    geom_text_repel(aes(label=Country), size = 3, segment.color="grey", max.overlaps = Inf)+ 
    theme(legend.position="bottom", legend.title = element_text(size=9, face="bold"), legend.text=element_text(size=9, face="plain"))+
    scale_x_continuous(limits = c(-6, 101))+
    # geom_text(x=-1, y=2.05, label="Greater than\n2% share of\nmonthly GNI\nper capita is\nunaffordable", col = "grey40", hjust = "right", vjust = "bottom",  fontface = "italic", size = 3)+ 
    scale_y_continuous(breaks=c(0,2,5,10,15,20))+
    theme(panel.background = element_rect(fill = "#EFEDE5", colour = "#EFEDE5"))
  
  # affplot
  
  
  
  Y4_data$hsetplotfill <- Y4_data$Country==CB_Y4$Country[counter]
  
  hsetplot <- ggplot() +
    scale_fill_manual(values = c("#9a9597", "#ffb237"))+
    geom_col(data=Y4_data, aes(fill=Y4_data$hsetplotfill, x=reorder(Y4_data$Country, -Y4_data$`2.1.1`), y=Y4_data$`2.1.1`))+coord_flip()+theme_classic() +theme(axis.title.x = element_text(face = "bold", size=7), axis.text.x = element_text(face="plain",size=6),axis.title.y = element_blank(), axis.text.y = element_text(size=6), panel.background = element_rect(fill = "#EFEDE5", colour = "#EFEDE5"))+
    theme(legend.position="none")+labs(y= "Handset Affordability")
  hsetplot
  
  fbupspplot <- ggplot() +
    scale_fill_manual(values = c("#9a9597", "#ffb237"))+
    geom_col(data=Y4_data, aes(fill=Y4_data$hsetplotfill, x=reorder(Y4_data$Country, -Y4_data$`1.2.1`), y=Y4_data$`1.2.1`))+coord_flip()+theme_classic() +theme(axis.title.x = element_text(face = "bold", size=7), axis.text.x = element_text(face="plain",size=6),axis.title.y = element_blank(), axis.text.y = element_text(size=6), panel.background = element_rect(fill = "#EFEDE5", colour = "#EFEDE5"))+
    theme(legend.position="none")+labs(y= "Average Fixed Broadband Upload Speed")
  fbupspplot
  
  
  #
  # Gender Gap Flextable 2 & 3
  ggtb2 <- data.frame(GGR_metrics = c(CB_Y4$fmai[counter], "Gender Gap Ratio"))
  
  ggtb3 <- data.frame(GGR_text = c(paste0(CB_Y4$PP_DIFF[counter], " pp"), "Δ E4 to E5"))
  ggtb4 <- data.frame(Rlbls = c("Gender Gap\nInternet Access", "Female Pop (%)\nWith Access", "Male Pop (%)\nWith Access"),
                      E2 = c(CB_Y4$`1.1.4) Gender gap in internet access_E2`[counter], CB_Y4$`BG21) Female internet users_E2`[counter], CB_Y4$`BG20) Male internet users_E2`[counter]),
                      E3 = c(CB_Y4$`1.1.4) Gender gap in internet access_E3`[counter], CB_Y4$`BG21) Female internet users_E3`[counter], CB_Y4$`BG20) Male internet users_E3`[counter]),
                      E4 = c(CB_Y4$`1.1.4) Gender gap in internet access_E4`[counter], CB_Y4$`BG21) Female internet users_E4`[counter], CB_Y4$`BG20) Male internet users_E4`[counter]),
                      E5 = c(CB_Y4$`1.1.4) Gender gap in internet access_E5`[counter], CB_Y4$`BG21) Female internet users_E5`[counter], CB_Y4$`BG20) Male internet users_E5`[counter]))
  
  
  ggtb5 <- data.frame(Rlbls = c("Gender Gap\nMobile Access", "Female Pop (%)\nWith Access", "Male Pop (%)\nWith Access"),
                      E2 = c(CB_Y4$`1.1.5) Gender gap in mobile phone access_E2`[counter], CB_Y4$`BG23) Female mobile phone subscribers_E2`[counter], CB_Y4$`BG22) Male mobile phone subscribers_E2`[counter]),
                      E3 = c(CB_Y4$`1.1.5) Gender gap in mobile phone access_E3`[counter], CB_Y4$`BG23) Female mobile phone subscribers_E3`[counter], CB_Y4$`BG22) Male mobile phone subscribers_E3`[counter]),
                      E4 = c(CB_Y4$`1.1.5) Gender gap in mobile phone access_E4`[counter], CB_Y4$`BG23) Female mobile phone subscribers_E4`[counter], CB_Y4$`BG22) Male mobile phone subscribers_E4`[counter]),
                      E5 = c(CB_Y4$`1.1.5) Gender gap in mobile phone access_E5`[counter], CB_Y4$`BG23) Female mobile phone subscribers_E5`[counter], CB_Y4$`BG22) Male mobile phone subscribers_E5`[counter]))
                      
                    
  ggtbft2 <- flextable(ggtb2)
  ggtbft3 <- flextable(ggtb3)
  ggtbft4 <- flextable(ggtb4)
  ggtbft5 <- flextable(ggtb5)
  
  
  ggtbft2 <- color(ggtbft2, part = "header", color = "#FFFFFF")
  ggtbft2 <- fontsize(ggtbft2, i = 1, part = "body", size = 44)
  ggtbft2 <- fontsize(ggtbft2, i = 2, part = "body", size = 24)
  ggtbft2 <- bold(ggtbft2, j = 1, part = "body")
  ggtbft2 <- border_remove(x = ggtbft2)
  ggtbft2 <- width(ggtbft2, j=1, width = 5.6)
  ggtbft2 <- delete_part(ggtbft2, part = "header")
  
  ggtbft3 <- color(ggtbft3, part = "header", color = "#FFFFFF")
  ggtbft3 <- fontsize(ggtbft3, i = 1, part = "body", size = 44)
  ggtbft3 <- fontsize(ggtbft3, i = 2, part = "body", size = 24)
  ggtbft3 <- bold(ggtbft3, i = 1, part = "body")
  ggtbft3 <- border_remove(x = ggtbft3)
  ggtbft3 <- width(ggtbft3, j=1, width = 5.6)
  ggtbft3 <- delete_part(ggtbft3, part = "header")
  
  ggtbft4 <- color(ggtbft4, part = "header", color = "#FFFFFF")
  ggtbft4 <- color(ggtbft4, part = "header", i = 1, j=1, color = "#2B458F")
  ggtbft4 <- bg(ggtbft4, bg = "#2B458F", part = "header")
  ggtbft4 <- bg(ggtbft4, bg = "#EFEDE5", part = "body")
  ggtbft4 <- bold(ggtbft4, part = "body")
  ggtbft4 <- width(ggtbft4, j=1, width = 1.5)
  ggtbft4 <- border_remove(x = ggtbft4)
  ggtbft4 <- border_outer(ggtbft4, part="all", border = fp_border(color="white", width = 1) )
  ggtbft4 <- border_inner_h(ggtbft4, part="all", border = fp_border(color="white", width = 1) )
  ggtbft4 <- border_inner_v(ggtbft4, part="all", border = fp_border(color="white", width = 1) )
  ggtbft4  <- colformat_double(
    x = ggtbft4 ,
    big.mark=",", digits = 1, na_str = "N/A")
  
  
  ggtbft5 <- color(ggtbft5, part = "header", color = "#FFFFFF")
  ggtbft5 <- color(ggtbft5, part = "header", i = 1, j=1, color = "#2B458F")
  ggtbft5 <- bg(ggtbft5, bg = "#2B458F", part = "header")
  ggtbft5 <- bg(ggtbft5, bg = "#EFEDE5", part = "body")
  ggtbft5 <- bold(ggtbft5, part = "body")
  ggtbft5 <- width(ggtbft5, j=1, width = 1.5)
  ggtbft5 <- border_remove(x = ggtbft5)
  ggtbft5 <- border_outer(ggtbft5, part="all", border = fp_border(color="white", width = 1) )
  ggtbft5 <- border_inner_h(ggtbft5, part="all", border = fp_border(color="white", width = 1) )
  ggtbft5 <- border_inner_v(ggtbft5, part="all", border = fp_border(color="white", width = 1) )
  ggtbft5  <- colformat_double(
    x = ggtbft5,
    big.mark=",", digits = 1, na_str = "N/A")
  
  ggtbft5 
  
  Y4_data_counter <- 0
  for(tempcounter in seq_along(Y4_data$Country)){
    if(Y4_data$Country[tempcounter]==CB_Y4$Country[counter]){
      Y4_data_counter <- tempcounter
    }
  }
  
  
  Y4_data$`2.1.2` <- round(Y4_data$`2.1.2`, 1)
  fbafftb <- data.frame(fbaffdf = c(paste0(Y4_data$`2.1.2`[Y4_data_counter],"%"), paste0("Cost of 1GB Mobile Data (Prepaid) as a share of GNI per capita in ", Y4_data$Country[Y4_data_counter])))
  
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
  flbsubs <- data.frame(flbsubs = c(paste0("1 for 2"), paste0("1GB of mobile broadband data for 2% or less of monthly GNI per capita is considered affordable")))
  
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
  
  
  srtitledata <- data.table(title = "")
  srtitle <- flextable(srtitledata)
  srtitle <- set_header_labels(srtitle, title="Index Score & Ranking")
  srtitle <- fontsize(srtitle, part = "all", size = 12)
  srtitle <- bold(srtitle, part = "all")
  srtitle <-  bg(srtitle, bg = "white", part = "all")
  srtitle <- width(srtitle, j=1, width = 5.6)
  srtitle <- border_remove(x = srtitle)
  srtitle <- padding(srtitle, padding=0, part = "all")
  
  mappath <- paste0("Silhouettes/",CB_Y4$Country[counter],".png")
  flagpath <- paste0("Flags-Icon-Set/",CB_Y4$ISO2[counter],".png")
  
  if(CB_Y4$Country[counter]=="Côte d'Ivoire"){
    mappath <- "Silhouettes/Cote d_Ivoire.png"
  }
  # print(flagpath)
  # print(mappath)
  doc <- add_slide(doc)
  doc <- ph_with(doc, value = paste("Country Briefing:", CB_Y4$Country[counter]), location = ph_location(left = 1.5, top = .3, width = 10, alignment = "l", height = .6))
  doc <- ph_with(doc, value = external_img(flagpath), location = ph_location(left = .5, top = .17, width = 1, height = .87))
  doc <- ph_with(doc, value = external_img(mappath), location = ph_location(left = 11.5, top = .17, width = 1, height = .84))
  doc <- phl_with_flextable(doc, 
                            olay = cb_offLayout, 4, srtitle)
  doc <- phl_with_flextable(doc, 
                            olay = cb_offLayout, 5, scorerank)
  srtitle <- set_header_labels(srtitle, title="Largest YoY Changes")
  doc <- phl_with_flextable(doc, 
                            olay = cb_offLayout, 10, srtitle)
  doc <- phl_with_flextable(doc, 
                            olay = cb_offLayout, 11, bigYOY)
  srtitle <- set_header_labels(srtitle, title="Strengths")
  doc <- phl_with_flextable(doc, 
                            olay = cb_offLayout, 6, srtitle)
  doc <- phl_with_flextable(doc, 
                            olay = cb_offLayout, 7, strengths)
  srtitle <- set_header_labels(srtitle, title="Areas for Improvement")
  doc <- phl_with_flextable(doc, 
                            olay = cb_offLayout, 8, srtitle)
  doc <- phl_with_flextable(doc, 
                            olay = cb_offLayout, 9, weaknesses)
  srtitle <- set_header_labels(srtitle, title="Gender Gap")
  doc <- phl_with_flextable(doc, 
                            olay = cb_offLayout, 13, srtitle)
  doc <- phl_with_flextable(doc, 
                            olay = cb_offLayout, 14, gender)
  
  doc <- add_slide(doc)
  doc <- ph_with(doc, value = external_img(flagpath), location = ph_location(left = .5, top = .17, width = 1, height = .87))
  doc <- ph_with(doc, value = external_img(mappath), location = ph_location(left = 11.5, top = .17, width = 1, height = .84))
  
  doc <- ph_with(doc, value = paste("YoY Change:", CB_Y4$Country[counter]), location = ph_location(left = 1.5, top = .3, width = 10, alignment = "l", height = .6))
  srtitle <- set_header_labels(srtitle, title="YoY Changes in Global and Regional Rank")
  doc <- phl_with_flextable(doc, 
                            olay = cb_offLayout2, 4, srtitle)
  doc <- phl_with_flextable(doc, 
                            olay = cb_offLayout2, 5, yoyrank)
  srtitle <- set_header_labels(srtitle, title="YoY Changes in Sub-Category Scores")
  doc <- phl_with_flextable(doc, 
                            olay = cb_offLayout, 6, srtitle)
  mypath <- paste("Y5_Radar_Charts/",CB_Y4$Country[counter], "3i Y5 Radar Chart", ".png")
  doc <- ph_with(doc, value = external_img(mypath),
                 location = ph_location(left = 7, top = 1.3, height=5, width=5))
  
  # legendpath <- paste("Ookla Quality Data - Legend.png")
  
  doc <- add_slide(doc)
  doc <- ph_with(doc, value = external_img(flagpath), location = ph_location(left = .5, top = .17, width = 1, height = .87))
  doc <- ph_with(doc, value = external_img(mappath), location = ph_location(left = 11.5, top = .17, width = 1, height = .84))
  
  doc <- ph_with(doc, value = paste("Fixed Quality:", CB_Y4$Country[counter]), location = ph_location(left = 1.5, top = .3, width = 10, alignment = "l", height = .6))
  doc <- ph_with(doc, value = ULbar, location = ph_location(left = 0.5, top = 1, width = 3.2, height = 4.25))
  
  doc <- ph_with(doc, value = DLbar, location = ph_location(left = 3.67, top = 1, width = 3.62, height = 4.25))
  
  doc <- ph_with(doc, value = obc_latency, location = ph_location(left = 0.5, top = 5.25, width = 6.79, height = 1.5))
  # doc <- ph_with(doc, value = dl_dist, location = ph_location(left = 8, top = 1, width = 4.8, height = 6))
  doc <- ph_with(doc, value = dlorderedbar, location = ph_location(left = 7.5, top = 1, width = 5.3, height = 6))
  
  
  
  doc <- add_slide(doc)
  doc <- ph_with(doc, value = external_img(flagpath), location = ph_location(left = .5, top = .17, width = 1, height = .87))
  doc <- ph_with(doc, value = external_img(mappath), location = ph_location(left = 11.5, top = .17, width = 1, height = .84))
  
  doc <- ph_with(doc, value = paste("Mobile Quality:", CB_Y4$Country[counter]), location = ph_location(left = 1.5, top = .3, width = 10, alignment = "l", height = .6))
  doc <- ph_with(doc, value = MULbar, location = ph_location(left = 0.5, top = 1, width = 3.2, height = 4.25))
  doc <- ph_with(doc, value = MDLbar, location = ph_location(left = 3.67, top = 1, width = 3.62, height = 4.25))
  doc <- ph_with(doc, value = mobc_latency, location = ph_location(left = 0.5, top = 5.25, width = 6.79, height = 1.5))
  # doc <- ph_with(doc, value = mdl_dist, location = ph_location(left = 8, top = 1, width = 4.8, height = 6))
  doc <- ph_with(doc, value = mdlorderedbar, location = ph_location(left = 7.5, top = 1, width = 5.3, height = 6))
  
  
  doc <- add_slide(doc)
  doc <- ph_with(doc, value = external_img(flagpath), location = ph_location(left = .5, top = .17, width = 1, height = .87))
  doc <- ph_with(doc, value = external_img(mappath), location = ph_location(left = 11.5, top = .17, width = 1, height = .84))
  
  # doc <- ph_with(doc, ggtbft2, 
  #                 location = ph_location(left = 4, top = 1.5, height=2, width=5))
  # doc <- ph_with(doc, ggtbft3, 
  #                 location = ph_location(left = 4, top = 3.3, height=2, width=5))
  #                 
  doc <- ph_with(doc, value = paste("Gender Gap:", CB_Y4$Country[counter]), location = ph_location(left = 1.5, top = .3, width = 10, alignment = "l", height = .6))
  srtitle <- set_header_labels(srtitle, title="Internet Access")
  doc <- ph_with(doc, srtitle, 
                 location = ph_location(left = 1, top = 1.08, height=0.25, width=2))
  srtitle <- set_header_labels(srtitle, title="Mobile Access")
  doc <- ph_with(doc, srtitle, 
                 location = ph_location(left = 7.5, top = 1.08, height=0.25, width=2))
  doc <- ph_with(doc, value = mobGGRbarplot,
                 location = ph_location(left = 7, top = 1.32, height=4, width=mbwidth))
  doc <- ph_with(doc, value = GGRbarplot,
                 location = ph_location(left = 0.5, top = 1.32, height=4, width=fbwidth))
  doc <- ph_with(doc, ggtbft4, 
                 location = ph_location(left = 1, top = 5.17, height=2, width=4.5))
  doc <- ph_with(doc, ggtbft5, 
                 location = ph_location(left = 7.5, top = 5.17, height=2, width=4.5))
  
  
  doc <- add_slide(doc)
  doc <- ph_with(doc, value = external_img(flagpath), location = ph_location(left = .5, top = .17, width = 1, height = .87))
  doc <- ph_with(doc, value = external_img(mappath), location = ph_location(left = 11.5, top = .17, width = 1, height = .84))
  
  doc <- ph_with(doc, value = paste("Affordability:", CB_Y4$Country[counter]), location = ph_location(left = 1.5, top = .3, width = 10, alignment = "l", height = .6))
  doc <- ph_with(doc, value = affplot,
                 location = ph_location(left = .45, top = 1, height=5.6, width=10))
  doc <- ph_with(doc, value = hsetplot,
                 location = ph_location(left = 10.68, top = 1, height=3, width=2.5))
  doc <- ph_with(doc, fbaffft,
                 location = ph_location(left = 10.68, top = 4.04, height=2, width=2.5))
  doc <- ph_with(doc, flbsubsft,
                 location = ph_location(left = 10.68, top = 5.56, height=2, width=2.5))
  
  print(CB_Y4$Country[counter])
  }
  doc <- add_slide(doc)
  
  
  
  print(doc, target = "3i_Y5_CB Tables.pptx")
  
  
  
  
  
