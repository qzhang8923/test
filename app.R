
#===========================================================================================================
#
#                                 CPM Project
#                                 
#                                 Last Update:  Feb 2021
#
#
#===========================================================================================================

#-----------------------------------------------------------------------------------------------------------
# Notes: Download function must work in an external browser 
#
# Themes: http://rstudio.github.io/shinythemes/
#-----------------------------------------------------------------------------------------------------------

#-----------------------------------------------------------------------------------------------------------
#install.packages("shiny",type="binary")
#install.packages("shinythemes")
#install.packages("readxl")
#install.packages("lubridate")
#install.packages("dplyr")
#install.packages("openxlsx")
#install.packages("officer")
#install.packages("flextable",dependencies=T)
#install.packages("plotrix")
#install.packages("plotly")
#install.packages("tidyr")
#install.packages("gridExtra")
#install.packages("rvg")
#install.packages("magick")
#install.packages("imager")
#install.packages("png")
#install.packages("zoo")
library(shiny)
library(shinythemes)
library(readxl)
library(lubridate) # function floor_date
library(dplyr) # function %>%
library(zoo) #function 'as.yearmon'
library(openxlsx)
library(officer)
library(flextable)
library(plotrix)
library(gridExtra)
library(grid)
library(rvg)
library(ggplot2)
library(ggrepel)
#library(magick)
#library(imager)
#library(png)
#library(tidyr)
#library(plotly)


options(shiny.maxRequestSize = 30*1024^10)#change uploading file's size


### read in image
# titlepageImage <- "https://raw.githubusercontent.com/alexsilve/pics/master/title_page.PNG"
# leftCornerImage <- "https://raw.githubusercontent.com/alexsilve/pics/master/left_corner.PNG"
# logoImage <-"https://raw.githubusercontent.com/alexsilve/pics/master/logo.PNG"
#MCImage <- "https://raw.githubusercontent.com/alexsilve/pics/master/ConvaTec.PNG"
MCImage <- "https://raw.githubusercontent.com/qzhang15403/pics/main/ConvaTec.PNG"

# titlepageImage <- tempfile()
# leftCornerImage <- tempfile()
# logoImage <- tempfile()
# download.file("https://raw.githubusercontent.com/alexsilve/pics/master/title_page.PNG", 
#              titlepageImage, mode = "wb")
# download.file("https://raw.githubusercontent.com/alexsilve/pics/master/left_corner.PNG", leftCornerImage, mode = "wb")
# download.file("https://raw.githubusercontent.com/alexsilve/pics/master/logo.PNG", logoImage, mode = "wb")




#loc_path <- "C:/Users/QI15403/OneDrive - ConvaTec Limited/Desktop/Quality Review Project/R Shiny/www/"
#-----------------------------------------------------------------------------------------------------------
#-----------------------------------------------------------------------------------------------------------
#                                        Functions 
#-----------------------------------------------------------------------------------------------------------
#Function to exract malfunction.code and harm code
f<-function(s)strsplit(as.character(s)," ")[[1]][1] # give value before 1st blank
t<-function(s)strsplit(as.character(s)," ")[[1]][2] # give value between 1st and 2nd blanks

#Function to row bind 
rbind.match.columns <- function(input1, input2) {
                                n.input1 <- ncol(input1)
                                n.input2 <- ncol(input2)
  
                                if (n.input2 < n.input1) {TF.names <- which(names(input2) %in% names(input1))
                                                          column.names <- names(input2[, TF.names])
                                                  } else {TF.names <- which(names(input1) %in% names(input2))
                                                          column.names <- names(input1[, TF.names])
                                                  }
                                return(rbind(input1[, column.names], input2[, column.names]))
}

#-----------------------------------------------------------------------------------------------------------
#
#                                         Define UI 
#
#-----------------------------------------------------------------------------------------------------------
ui <- navbarPage(
  
  theme = shinytheme("cerulean"),
  "CPM Calculation", # App title
  tabsetPanel(
    
    
    tabPanel("CPM Project", #tab title
             
             sidebarPanel(width = 3,
                          dateRangeInput('YTD', label = 'YTD Date Range (Current Year Jan 1 - Last Day of Reported Month)',
                                         start = cut(Sys.Date(), "year"), end = Sys.Date()),
                          #fileInput("sales","Choose Sales File (.csv)", multiple = FALSE,
                           #         accept = c("text/csv","text/comma-separated-values,text/plain",".csv")),
                          fileInput("sales","Choose 2018/2019/2020/2021 Non-IC Sales Eaches File (.xlsx)", multiple = TRUE,
                                    accept = c(".xlsx")),
                          fileInput("complaint","Choose Non-IC Complaint File (.xlsx)", multiple = FALSE,
                                    accept = c(".xlsx")),
                          fileInput("IC","Choose IC and Eurotec Complaints/Sales File (.xlsx)", multiple = FALSE,
                                    accept = c(".xlsx")),
                          fileInput("ref","Choose Region Ref Files (.xlsx)", multiple = TRUE,
                                    accept = c(".xlsx")),
                          fileInput("exclusion","Choose exclusive list (.xlsx)", multiple = TRUE,
                                    accept = c(".xlsx")),
                          fileInput("images", "Choose Images", multiple = TRUE, accept = c('image/png', 'image/jpeg')),
                          dateInput("report", label = "Report Date", value = Sys.Date()),
                          selectInput("xx", label = "Report Person", 
                                      choices = list("Qi Zhang" = "Qi Zhang", "Yaxin Zheng" = "Yaxin Zheng"), 
                                      selected = "Qi Zhang"),
                          sliderInput("plot_range", "Choose Time Range for Trend Chart (most recent xx months):",
                                      min = 6, max = 36,
                                      value = 12),
                          radioButtons("reduction_IC", label = "% of complaints reduction for IC",
                                       choices = list("15 %" = 0.85, "12 %" = 0.88, "10 %" = 0.9, "8 %" = 0.92), 
                                       selected = 0.9),
                          radioButtons("reduction_AWC", label = "% of complaints reduction for AWC",
                                       choices = list("15 %" = 0.85, "12 %" = 0.88, "10 %" = 0.9, "8 %" = 0.92), 
                                       selected = 0.9),
                          radioButtons("reduction_CC", label = "% of complaints reduction for CC",
                                       choices = list("15 %" = 0.85, "12 %" = 0.88, "10 %" = 0.9, "8 %" = 0.92), 
                                       selected = 0.9),
                          radioButtons("reduction_C", label = "% of complaints reduction for C",
                                       choices = list("15 %" = 0.85, "12 %" = 0.88, "10 %" = 0.9, "8 %" = 0.92), 
                                       selected = 0.9),
                          radioButtons("reduction_OC", label = "% of complaints reduction for OC",
                                       choices = list("15 %" = 0.85, "12 %" = 0.88, "10 %" = 0.9, "8 %" = 0.92), 
                                       selected = 0.9),
                          radioButtons("reduction_Eurotec", label = "% of complaints reduction for Eurotec",
                                       choices = list("15 %" = 0.85, "12 %" = 0.88, "10 %" = 0.9, "8 %" = 0.92), 
                                       selected = 0.9),
                          radioButtons("reduction_AP", label = "% of complaints reduction for GEM_APAP",
                                       choices = list("15 %" = 0.85, "12 %" = 0.88, "10 %" = 0.9, "8 %" = 0.92), 
                                       selected = 0.9),
                          radioButtons("reduction_LA", label = "% of complaints reduction for GEM_LATAM",
                                       choices = list("15 %" = 0.85, "12 %" = 0.88, "10 %" = 0.9, "8 %" = 0.92), 
                                       selected = 0.9),
                          radioButtons("reduction_MEA", label = "% of complaints reduction for GEM_MEA",
                                       choices = list("15 %" = 0.85, "12 %" = 0.88, "10 %" = 0.9, "8 %" = 0.92), 
                                       selected = 0.9),
                          radioButtons("reduction_EUR", label = "% of complaints reduction for EUR",
                                       choices = list("15 %" = 0.85, "12 %" = 0.88, "10 %" = 0.9, "8 %" = 0.92), 
                                       selected = 0.9),
                          radioButtons("reduction_NAM", label = "% of complaints reduction for NAM",
                                       choices = list("15 %" = 0.85, "12 %" = 0.88, "10 %" = 0.9, "8 %" = 0.92), 
                                       selected = 0.9),
                          downloadButton("slides", "Download BU and Region Dashboard"),
                          #downloadButton("slides_r", "Download Region Dashboard"),
                          downloadButton("file", "Download Summary File by BU"),
                          downloadButton("file_r", "Download Summary File by Region")
                          #downloadButton("TW_complaint", "Download TW file")
             ),
             
             mainPanel(
               p("=================================================================================================================================="),
               
               fluidRow(
                 column(5, h1("CPM Project"),
                        p("Copyright belong to ConvaTec Metrics Central team"),
                        p("Updated Date: Feb 2021")),
                 column(4, img(src = MCImage, height = 120,width = 400))
  
               ),
               p("=================================================================================================================================="),
               
               tabsetPanel(
                 tabPanel("Most Recent 12 Months CPM by BU",
                          fluidRow(
                            column(9, 
                                   p("################################################################################################"),
                                   p(strong("All.BU:"), "All Business Units;", strong("Non.IC:"), "non-Infusion Care BU combined;",
                                     strong("IC:"), "Infusion Care;", strong("WC:"), "Wound Care;",
                                     strong("CC:"), "Continence Care;", strong("C:"), "Critical Care;", strong("OC:"), "Ostomy Care."),
                                   p(span(strong("*Note:"), "Complaints for SKC (Skin Care) have been regrouped to under corresponding business units for OC and  AWC.", 
                                          style = "background-color: orange")),
                                   p("################################################################################################")
                            )  
                          ),
                          tableOutput("Monthly_tb"),
                          tableOutput("target")
                 ),
                 
                 tabPanel("YTD CPM by BU",
                          fluidRow(
                            column(9, 
                                   p("################################################################################################"),
                                   p(strong("All.BU:"), "All Business Units;", strong("Non.IC:"), "non-Infusion Care BU combined;",
                                     strong("IC:"), "Infusion Care;", strong("WC:"), "Wound Care;",
                                     strong("CC:"), "Continence Care;", strong("C:"), "Critical Care;", strong("OC:"), "Ostomy Care."),
                                   p(span(strong("*Note:"), "Complaints for SKC (Skin Care) have been regrouped to under corresponding business units for OC and  AWC.", 
                                          style = "background-color: orange")),
                                   p("################################################################################################")
                            )  
                          ),
                          uiOutput("target_ui"),
                          uiOutput("YTD_ui")
                 ),
                 

                 tabPanel("Trend Chart by BU",
                          radioButtons(inputId = "plot_option", 
                                       label = "Choose Business Unit", 
                                       choices = list("All BU Complaint" = "all_com", "All BU CPM" = "all", 
                                                      "All B.U. Combined" = "all_bu", "Non-IC CPM" = "NIC", 
                                                      "IC CPM" = "IC", "WC CPM" = "AWC", "OC CPM" = "OC", 
                                                      "Continence Care CPM" = "CC", "Critical Care CPM" = "C", 
                                                      "Eurotec CPM" = "Eurotec",
                                                      "Non-IC Complaint & CPM" = "Bar_Line"), 
                                       selected = "all",
                                       inline = T), 
                          plotOutput("plot")
                        
                 ),
                 tabPanel("Most Recent 12 Months CPM by Region",
                          fluidRow(
                            column(7, 
                                   p("###############################################################################"),
                                   p(strong("GEM:"), "Global Emerging Market (APAC+LATAM+MEA+TURKEY);"),
                                   p(strong("EUR:"), "European Countries (excluding Turkey);"),
                                   p(strong("NAM:"), "North America (US +PUERTO RICO+ CANADA)."),
                                   p(span(strong("*Note:"), "Complaints only include non-Infusion Care BU combined.", 
                                          style = "background-color: orange")),
                                   p("###############################################################################")
                            )  
                          ),
                          tableOutput("Monthly_tb_r"),
                          tableOutput("Monthly_tb_r2")
                 ),
               
                 tabPanel("YTD CPM by Region",
                          fluidRow(
                            column(7, 
                                   p("###############################################################################"),
                                   p(strong("GEM:"), "Global Emerging Market (APAC+LATAM+MEA+TURKEY);"),
                                   p(strong("EUR:"), "European Countries (excluding Turkey);"),
                                   p(strong("NAM:"), "North America (US +PUERTO RICO+ CANADA)."),
                                   p(span(strong("*Note:"), "Complaints only include non-Infusion Care.", 
                                          style = "background-color: orange")),
                                   p("###############################################################################")
                            )  
                          ),
                          uiOutput("target_ui_r"),
                          uiOutput("target_ui_r2"),
                          uiOutput("YTD_ui_r"),
                          uiOutput("YTD_ui_r2")
                 ),
               
                 tabPanel("Trend Chart by Region",
                          radioButtons(inputId = "plot_option_r", 
                                       label = "Choose Business Unit", 
                                       choices = list("GEM" = "GEM", "Europe" = "EUR", "North American" = "NAM", "APAC" = "AP", "LATAM" = "LA", "MEA" = "MEA"), 
                                       selected = "GEM",
                                       inline = T), 
                          plotOutput("plot_r") 
               ))
             )
    )
  )
)

#-----------------------------------------------------------------------------------------------------------
#
#                                          Define Server
#
#-----------------------------------------------------------------------------------------------------------
server <- function(input, output) {
  
  complaint0 <- reactive({if(is.null(input$complaint))return(NULL)
                         complaint <- read_excel(paste(input$complaint$datapath), 1)
                         complaint$`Date Created` <- as.Date(complaint$`Date Created`, format="%m/%d/%Y")
                         return(complaint)
  })

   
  complaint_BU <- reactive({if(is.null(complaint0())|is.null(input$IC))return(NULL) 
                            complaint <- complaint0()
                            IC <- read_excel(paste(input$IC$datapath), sheet = "IC Complaints")
                            Eurotec <- read_excel(paste(input$IC$datapath), sheet = "Eurotec Complaints")
                            summary <- complaint %>% group_by(month = floor_date(`Date Created`, "month"), `Business Unit (Hyperion)`)  %>% dplyr::summarize(count = n())
                            colnames(summary) <- c("month", "BU", "complaints")
                            summary[which(summary$BU=="Total Continence Care"),]$BU <- "Continence"
                            summary[which(summary$BU=="Total Critical Care"),]$BU <- "Critical"
                            summary[which(summary$BU=="Ostomy"),]$BU <- "OSTOMY"
                            summary[which(summary$BU=="Wound"),]$BU <- "WOUND"
                            IC$BU <- "IC"
                            IC <- IC[c("month", "BU", "complaints")]
                            Eurotec$BU <- "Eurotec"
                            Eurotec <- Eurotec[c("month", "BU", "complaints")]
                            summary <- data.frame(summary)
                            summary_all <- rbind(summary, IC, Eurotec)
                            summary_all <- summary_all[(summary_all$month >= "2018-01-01" & summary_all$month <= input$YTD[2]),]
                            summary_all[summary_all$BU=="OSTOMY",]$complaints <- summary_all[summary_all$BU=="OSTOMY",]$complaints +summary_all[summary_all$BU=="Eurotec",]$complaints
                            
                            dates <- seq(min(summary_all$month), max(summary_all$month), by = "month")
                            catg <- unique(summary_all$BU)
                            full_list <- as.data.frame(lapply(expand.grid(dates, catg), unlist))
                            colnames(full_list) <- c("month","BU")
                            summary_all <- merge(full_list, summary_all, by = c('month','BU'), all.x = T)
                            summary_all$complaints[is.na(summary_all$complaints)] <- 0
                            return(summary_all)
                            })
  
  complaint_region <- reactive({if(is.null(complaint0())|is.null(input$ref))return(NULL) 
                                complaint <- complaint0()
                                region_ref <- read_excel(input$ref$datapath[input$ref$name=="Region LookUp.xlsx"], sheet = 1)
                                region_merge <- merge(complaint[which(complaint$`Event Country`!="Unknown"),], region_ref, by.x = "Event Country", by.y = "Country/Area", all.x = T)
                                ifelse(nrow(region_merge[is.na(region_merge$Region),]) > 0, print(region_merge[is.na(region_merge$Region),]), print("Great! No new countries that not in mapping file!"))
                                validate(need(nrow(region_merge[is.na(region_merge$Region),]) == 0,
                                              "Warning Message: New Event County not in Region lookup file."))
                                summary <- region_merge %>% group_by(month = floor_date(`Date Created`, "month"), `Region`)  %>% dplyr::summarize(count = n())
                                colnames(summary) <- c("month", "Region", "complaints")
                                summary[which(summary$Region=="Global Emerging Markets"),]$Region <- "GEM"
                                summary[which(summary$Region=="Europe"),]$Region <- "EUR"
                                summary[which(summary$Region=="North America"),]$Region <- "NAM"
                                summary <- summary[(summary$month >= "2018-01-01" & summary$month <= input$YTD[2]),]
                                dates <- seq(min(summary$month), max(summary$month), by = "month")
                                catg <- unique(summary$Region)
                                full_list <- as.data.frame(lapply(expand.grid(dates, catg), unlist))
                                colnames(full_list) <- c("month","Region")
                                summary <- merge(full_list, summary, by = c('month','Region'), all.x = T)
                                summary$complaints[is.na(summary$complaints)] <- 0
                                return(summary)
  })
  
  complaint_region2 <- reactive({if(is.null(complaint_region()))return(NULL)
                                complaint <- complaint0()
                                region_ref <- read_excel(input$ref$datapath[input$ref$name=="Region LookUp.xlsx"], sheet = 1)
                                region_ref2 <- read_excel(input$ref$datapath[input$ref$name=="Region LookUp.xlsx"], sheet = 2)
                                region_merge <- merge(complaint[which(complaint$`Event Country`!="Unknown"),], region_ref, by.x = "Event Country", by.y = "Country/Area", all.x = T)
                                region_merge <- region_merge[region_merge$Region=="Global Emerging Markets",]
                                region_merge <- merge(region_merge[which(region_merge$`Event Country`!="Unknown"),], region_ref2, by.x = "Event Country", by.y = "Country/Area", all.x = T)
                                summary <- region_merge %>% group_by(month = floor_date(`Date Created`, "month"), `Region2`)  %>% dplyr::summarize(count = n())
                                colnames(summary) <- c("month", "Region2", "complaints")
                                summary[which(summary$Region2=="Asia Pacific"),]$Region2 <- "AP"
                                summary[which(summary$Region2=="Latin America"),]$Region2 <- "LA"
                                summary[which(summary$Region2=="MEA"),]$Region2 <- "MEA"
                                summary <- summary[(summary$month >= "2018-01-01" & summary$month <= input$YTD[2]),]
                                dates <- seq(min(summary$month), max(summary$month), by = "month")
                                catg <- unique(summary$Region2)
                                full_list <- as.data.frame(lapply(expand.grid(dates, catg), unlist))
                                colnames(full_list) <- c("month","Region2")
                                summary <- merge(full_list, summary, by = c('month','Region2'), all.x = T)
                                summary$complaints[is.na(summary$complaints)] <- 0
                                return(summary)
    })
  
  complaint_NIC <- reactive({if(is.null(complaint_BU()))return(NULL)
                   complaint <- complaint_BU()[which(complaint_BU()$BU!="IC" & complaint_BU()$BU!="Eurotec"),]
                   complaint_NIC <- aggregate(complaint$complaints,
                                              by = list(complaint$month),
                                              sum)
                   colnames(complaint_NIC) <- c("month", "complaints")
                   return(complaint_NIC)  
  })
  
  complaint_all <- reactive({if(is.null(complaint_BU()))return(NULL)
                   complaint <- complaint_BU()[which(complaint_BU()$BU!="Eurotec"),]
                   complaint_all <- aggregate(complaint$complaints,
                                              by = list(complaint$month),
                                              sum)
                   colnames(complaint_all) <- c("month", "complaints")
                   return(complaint_all)  
  })
  
  sales <- reactive({if(is.null(input$sales)|is.null(input$IC)|is.null(input$exclusion))return(NULL) 
                     excl_SKC <- read_excel(paste(input$exclusion$datapath), sheet = "SKC")
                     excl_Cobranding <- read_excel(paste(input$exclusion$datapath), sheet = "Co-branding")
                     excl_Eurotec <- read_excel(paste(input$exclusion$datapath), sheet = "Eurotec")
                     sales0 <- read_excel(input$sales$datapath[input$sales$name=="Global Eaches 2021 By Region.xlsx"], 1)
                     sales1 <- read_excel(input$sales$datapath[input$sales$name=="Global Eaches 2020 By Region.xlsx"], 1)
                     sales2 <- read_excel(input$sales$datapath[input$sales$name=="Global Eaches 2019 By Region.xlsx"], 1)
                     sales3 <- read_excel(input$sales$datapath[input$sales$name=="Global Eaches 2018 By Region.xlsx"], 1)
                     #sales0[which(sales0$`Hyperion Business Unit`=="Continence Care"),]$`Hyperion Business Unit` <- "Total Continence Care"
                     #sales1[which(sales1$`Hyperion Business Unit`=="Continence Care"),]$`Hyperion Business Unit` <- "Total Continence Care"
                     sales2[which(sales2$`Hyperion Business Unit`=="Continence Care"),]$`Hyperion Business Unit` <- "Total Continence Care"
                     sales3[which(sales3$`Hyperion Business Unit`=="Continence Care"),]$`Hyperion Business Unit` <- "Total Continence Care"
                     sales_IC <- read_excel(input$IC$datapath, sheet = "IC Sales")
                     sales_Eurotec <- read_excel(input$IC$datapath, sheet = "Eurotec Sales")
                     #sales_IC <- sales_IC[2,]
                     sales0$ICC <- gsub("F", "", sales0$ICC)
                     sales0$ICC <- gsub("UNO_", "", sales0$ICC) # delete "F" and "UNO_" from ICC
                     sales0$ICC <- ifelse(substr(sales0$ICC, 1, 1) =="0" & !grepl("\\.| |-", sales0$ICC), 
                                             sub("^0+", "", sales0$ICC), sales0$ICC)
                     sales1$ICC <- gsub("F", "", sales1$ICC)
                     sales1$ICC <- gsub("UNO_", "", sales1$ICC) # delete "F" and "UNO_" from ICC
                     sales1$ICC <- ifelse(substr(sales1$ICC, 1, 1) =="0" & !grepl("\\.| |-", sales1$ICC), 
                                          sub("^0+", "", sales1$ICC), sales1$ICC)
                     sales2$ICC <- gsub("F", "", sales2$ICC)
                     sales2$ICC <- gsub("UNO_", "", sales2$ICC) # delete "F" and "UNO_" from ICC
                     sales2$ICC <- ifelse(substr(sales2$ICC, 1, 1) =="0" & !grepl("\\.| |-", sales2$ICC), 
                                          sub("^0+", "", sales2$ICC), sales2$ICC)
                     sales3$ICC <- gsub("F", "", sales3$ICC)
                     sales3$ICC <- gsub("UNO_", "", sales3$ICC) # delete "F" and "UNO_" from ICC
                     sales3$ICC <- ifelse(substr(sales3$ICC, 1, 1) =="0" & !grepl("\\.| |-", sales3$ICC), 
                                          sub("^0+", "", sales3$ICC), sales3$ICC)
                     
                     sales0 <- sales0[which(!sales0$`ICC`%in% excl_SKC$ICC  & !sales0$`ICC`%in% excl_Cobranding$ICC  &  !sales0$`ICC`%in% excl_Eurotec$ICC),]
                     sales1 <- sales1[which(!sales1$`ICC`%in% excl_SKC$ICC  & !sales1$`ICC`%in% excl_Cobranding$ICC  &  !sales1$`ICC`%in% excl_Eurotec$ICC),]
                     sales2 <- sales2[which(!sales2$`ICC`%in% excl_SKC$ICC  & !sales2$`ICC`%in% excl_Cobranding$ICC  &  !sales2$`ICC`%in% excl_Eurotec$ICC),]
                     sales3 <- sales3[which(!sales3$`ICC`%in% excl_SKC$ICC  & !sales3$`ICC`%in% excl_Cobranding$ICC  &  !sales3$`ICC`%in% excl_Eurotec$ICC),]
                     
                     
                     
                     tb_s0 <- aggregate(sales0[4], by = list(sales0$`Hyperion Business Unit`), sum)
                     for (i in 5:ncol(sales0)){
                       each <- aggregate(sales0[i], by = list(sales0$`Hyperion Business Unit`), sum)
                       tb_s0 <- cbind(tb_s0, each[,2])
                     }
                     
                     tb_s1 <- aggregate(sales1[4], by = list(sales1$`Hyperion Business Unit`), sum)
                     for (i in 5:ncol(sales1)){
                       each <- aggregate(sales1[i], by = list(sales1$`Hyperion Business Unit`), sum)
                       tb_s1 <- cbind(tb_s1, each[,2])
                     }

                     tb_s2 <- aggregate(sales2[4], by = list(sales2$`Hyperion Business Unit`), sum)
                     for (i in 5:ncol(sales2)){
                       each <- aggregate(sales2[i], by = list(sales2$`Hyperion Business Unit`), sum)
                       tb_s2 <- cbind(tb_s2, each[,2])
                     }
                     
                     tb_s3 <- aggregate(sales3[4], by = list(sales3$`Hyperion Business Unit`), sum)
                     for (i in 5:ncol(sales3)){
                       each <- aggregate(sales3[i], by = list(sales3$`Hyperion Business Unit`), sum)
                       tb_s3 <- cbind(tb_s3, each[,2])
                     }

                     
                     tb_s <- cbind(tb_s3, tb_s2[,-1], tb_s1[,-1], tb_s0[,-1])
                     n <- length(seq(from = input$YTD[1] - years(3), to = input$YTD[2], by='month')) 
                     tb_s <- tb_s[,1:(n + 1)] #remove empty cols
                     names(tb_s) <- c("BU", seq(as.Date(input$YTD[1] - years(3)), by = "month", length.out = n))
                     tb_s[which(tb_s$BU=="Infusion devices"|tb_s$BU=="Infusion Devices"),]$BU <- "IC"
                     tb_s[which(tb_s$BU=="Ostomy"),]$BU <- "OSTOMY"
                     tb_s[which(tb_s$BU=="Total Continence Care"),]$BU <- "Continence"
                     tb_s[which(tb_s$BU=="Total Critical Care"),]$BU <- "Critical"
                     tb_s[which(tb_s$BU=="Wound"),]$BU <- "WOUND"
                     
                     tb_s <- tb_s[which(tb_s$BU!="IC"),]
                     sales_IC <- sales_IC[,-c(2:13)] # delete 2017 data
                     names(sales_IC) <- names(tb_s)
                     names(sales_Eurotec) <- names(tb_s)
                     tb_s <- rbind(tb_s, sales_IC, sales_Eurotec) 
                     tb_s[which(tb_s$BU=="OSTOMY"),2:ncol(tb_s)] <- tb_s[which(tb_s$BU=="OSTOMY"),2:ncol(tb_s)] + tb_s[which(tb_s$BU=="Eurotec"),2:ncol(tb_s)]
                     return(tb_s)
  })
  
  sales_download <- reactive({if(is.null(sales()))return(NULL) 
                              sales <- sales()
                              n <- length(seq(from = input$YTD[1] - years(3), to = input$YTD[2], by='month')) 
                              names(sales) <- c("BU", as.character(seq(as.Date(input$YTD[1] - years(3)), by = "month", length.out = n)))
                              return(sales)
  })
  
  sales_region <- reactive({if(is.null(input$sales)|is.null(input$exclusion))return(NULL) 
                  excl_SKC <- read_excel(paste(input$exclusion$datapath), sheet = "SKC")
                  excl_Cobranding <- read_excel(paste(input$exclusion$datapath), sheet = "Co-branding")
                  excl_Eurotec <- read_excel(paste(input$exclusion$datapath), sheet = "Eurotec")
                  GEM0 <- read_excel(input$sales$datapath[input$sales$name=="Global Eaches 2021 By Region.xlsx"], 3)
                  GEM1 <- read_excel(input$sales$datapath[input$sales$name=="Global Eaches 2020 By Region.xlsx"], 3)
                  GEM2 <- read_excel(input$sales$datapath[input$sales$name=="Global Eaches 2019 By Region.xlsx"], 3)
                  EUR0 <- read_excel(input$sales$datapath[input$sales$name=="Global Eaches 2021 By Region.xlsx"], 7)
                  EUR1 <- read_excel(input$sales$datapath[input$sales$name=="Global Eaches 2020 By Region.xlsx"], 7)
                  EUR2 <- read_excel(input$sales$datapath[input$sales$name=="Global Eaches 2019 By Region.xlsx"], 7)
                  NAM0 <- read_excel(input$sales$datapath[input$sales$name=="Global Eaches 2021 By Region.xlsx"], 8)
                  NAM1 <- read_excel(input$sales$datapath[input$sales$name=="Global Eaches 2020 By Region.xlsx"], 8)
                  NAM2 <- read_excel(input$sales$datapath[input$sales$name=="Global Eaches 2019 By Region.xlsx"], 8)
                  GEM0$Region <- "GEM"
                  GEM1$Region <- "GEM"
                  GEM2$Region <- "GEM"
                  EUR0$Region <- "EUR"
                  EUR1$Region <- "EUR"
                  EUR2$Region <- "EUR"
                  NAM0$Region <- "NAM"
                  NAM1$Region <- "NAM"
                  NAM2$Region <- "NAM"
                  year0 <- rbind(GEM0, EUR0, NAM0)
                  year1 <- rbind(GEM1, EUR1, NAM1)
                  year2 <- rbind(GEM2, EUR2, NAM2)
                  year0 <- year0[which(year0$`Hyperion Business Unit`!="Infusion Devices" & year0$`Hyperion Business Unit`!="0" & !is.na(year0$`Hyperion Business Unit`)),]
                  year1 <- year1[which(year1$`Hyperion Business Unit`!="Infusion Devices" & year1$`Hyperion Business Unit`!="0" & !is.na(year1$`Hyperion Business Unit`)),]
                  year2 <- year2[which(year2$`Hyperion Business Unit`!="Infusion Devices" & year2$`Hyperion Business Unit`!="0" & !is.na(year2$`Hyperion Business Unit`)),]
               
                  year0$ICC <- gsub("F", "", year0$ICC)
                  year0$ICC <- gsub("UNO_", "", year0$ICC) # delete "F" and "UNO_" from ICC
                  year0$ICC <- ifelse(substr(year0$ICC, 1, 1) =="0" & !grepl("\\.| |-", year0$ICC), 
                                       sub("^0+", "", year0$ICC), year0$ICC)
                  year1$ICC <- gsub("F", "", year1$ICC)
                  year1$ICC <- gsub("UNO_", "", year1$ICC) # delete "F" and "UNO_" from ICC
                  year1$ICC <- ifelse(substr(year1$ICC, 1, 1) =="0" & !grepl("\\.| |-", year1$ICC), 
                                       sub("^0+", "", year1$ICC), year1$ICC)
                  year2$ICC <- gsub("F", "", year2$ICC)
                  year2$ICC <- gsub("UNO_", "", year2$ICC) # delete "F" and "UNO_" from ICC
                  year2$ICC <- ifelse(substr(year2$ICC, 1, 1) =="0" & !grepl("\\.| |-", year2$ICC), 
                                       sub("^0+", "", year2$ICC), year2$ICC)
                  
                  year0 <- year0[which(!year0$`ICC`%in% excl_SKC$ICC  & !year0$`ICC`%in% excl_Cobranding$ICC  &  !year0$`ICC`%in% excl_Eurotec$ICC),]
                  year1 <- year1[which(!year1$`ICC`%in% excl_SKC$ICC  & !year1$`ICC`%in% excl_Cobranding$ICC  &  !year1$`ICC`%in% excl_Eurotec$ICC),]
                  year2 <- year2[which(!year2$`ICC`%in% excl_SKC$ICC  & !year2$`ICC`%in% excl_Cobranding$ICC  &  !year2$`ICC`%in% excl_Eurotec$ICC),]   
                  sales0 <- year0[,-c(1:3)]
                  sales1 <- year1[,-c(1:3)]
                  sales2 <- year2[,-c(1:3)]
                  sales0[1:(ncol(sales0)-1)] <- sapply(sales0[1:(ncol(sales0) - 1)], as.numeric)
                  sales1[1:(ncol(sales1)-1)] <- sapply(sales1[1:(ncol(sales1) - 1)], as.numeric)
                  sales2[1:(ncol(sales2)-1)] <- sapply(sales2[1:(ncol(sales2) - 1)], as.numeric)
                  
                  tb_s0 <- aggregate(sales0[1], by = list(sales0$Region), sum)
                  for (i in 2:(ncol(sales0)-1)){
                    each <- aggregate(sales0[i], by = list(sales0$Region), sum)
                    tb_s0 <- cbind(tb_s0,each[,2])
                  }
                  
                  tb_s1 <- aggregate(sales1[1], by = list(sales1$Region), sum)
                  for (i in 2:(ncol(sales1)-1)){
                    each <- aggregate(sales1[i], by = list(sales1$Region), sum)
                    tb_s1 <- cbind(tb_s1,each[,2])
                  }
                  
                  tb_s2 <- aggregate(sales2[1], by = list(sales2$Region), sum)
                  for (i in 2:(ncol(sales2)-1)){
                    each <- aggregate(sales2[i], by = list(sales2$Region), sum)
                    tb_s2 <- cbind(tb_s2,each[,2])
                  }
                  
                  sales <- cbind(tb_s2, tb_s1[,-1], tb_s0[,-1])
                  n <- length(seq(from = input$YTD[1] - years(2), to = input$YTD[2], by='month')) 
                  sales <- sales[,1:(n + 1)]
                  names(sales) <- c("Region", seq(as.Date(input$YTD[1] - years(2)), by = "month", length.out = n))
                  return(sales)
  })
  
  sales_region_download <- reactive({if(is.null(sales_region()))return(NULL) 
                                    sales <- sales_region()
                                    n <- length(seq(from = input$YTD[1] - years(2), to = input$YTD[2], by='month')) 
                                    names(sales) <- c("BU", as.character(seq(as.Date(input$YTD[1] - years(2)), by = "month", length.out = n)))
                                    return(sales)
  })
  sales_region2 <- reactive({if(is.null(input$sales)|is.null(input$exclusion))return(NULL) 
                             excl_SKC <- read_excel(paste(input$exclusion$datapath), sheet = "SKC")
                             excl_Cobranding <- read_excel(paste(input$exclusion$datapath), sheet = "Co-branding")
                             excl_Eurotec <- read_excel(paste(input$exclusion$datapath), sheet = "Eurotec")
                             AP0 <- read_excel(input$sales$datapath[input$sales$name=="Global Eaches 2021 By Region.xlsx"], 4)
                             AP1 <- read_excel(input$sales$datapath[input$sales$name=="Global Eaches 2020 By Region.xlsx"], 4)
                             AP2 <- read_excel(input$sales$datapath[input$sales$name=="Global Eaches 2019 By Region.xlsx"], 4)
                             LA0 <- read_excel(input$sales$datapath[input$sales$name=="Global Eaches 2021 By Region.xlsx"], 5)
                             LA1 <- read_excel(input$sales$datapath[input$sales$name=="Global Eaches 2020 By Region.xlsx"], 5)
                             LA2 <- read_excel(input$sales$datapath[input$sales$name=="Global Eaches 2019 By Region.xlsx"], 5)
                             MEA0 <- read_excel(input$sales$datapath[input$sales$name=="Global Eaches 2021 By Region.xlsx"], 6)
                             MEA1 <- read_excel(input$sales$datapath[input$sales$name=="Global Eaches 2020 By Region.xlsx"], 6)
                             MEA2 <- read_excel(input$sales$datapath[input$sales$name=="Global Eaches 2019 By Region.xlsx"], 6)
                             AP0$Region <- "AP"
                             AP1$Region <- "AP"
                             AP2$Region <- "AP"
                             LA0$Region <- "LA"
                             LA1$Region <- "LA"
                             LA2$Region <- "LA"
                             MEA0$Region <- "MEA"
                             MEA1$Region <- "MEA"
                             MEA2$Region <- "MEA"
                             year0 <- rbind(AP0, LA0, MEA0)
                             year1 <- rbind(AP1, LA1, MEA1)
                             year2 <- rbind(AP2, LA2, MEA2)
                             year0 <- year0[which(year0$`Hyperion Business Unit`!="Infusion Devices" & year0$`Hyperion Business Unit`!="0" & !is.na(year0$`Hyperion Business Unit`)),]
                             year1 <- year1[which(year1$`Hyperion Business Unit`!="Infusion Devices" & year1$`Hyperion Business Unit`!="0" & !is.na(year1$`Hyperion Business Unit`)),]
                             year2 <- year1[which(year2$`Hyperion Business Unit`!="Infusion Devices" & year2$`Hyperion Business Unit`!="0" & !is.na(year2$`Hyperion Business Unit`)),]
                             
                             year0$ICC <- gsub("F", "", year0$ICC)
                             year0$ICC <- gsub("UNO_", "", year0$ICC) # delete "F" and "UNO_" from ICC
                             year0$ICC <- ifelse(substr(year0$ICC, 1, 1) =="0" & !grepl("\\.| |-", year0$ICC), 
                                                 sub("^0+", "", year0$ICC), year0$ICC)
                             year1$ICC <- gsub("F", "", year1$ICC)
                             year1$ICC <- gsub("UNO_", "", year1$ICC) # delete "F" and "UNO_" from ICC
                             year1$ICC <- ifelse(substr(year1$ICC, 1, 1) =="0" & !grepl("\\.| |-", year1$ICC), 
                                                 sub("^0+", "", year1$ICC), year1$ICC)
                             year2$ICC <- gsub("F", "", year2$ICC)
                             year2$ICC <- gsub("UNO_", "", year2$ICC) # delete "F" and "UNO_" from ICC
                             year2$ICC <- ifelse(substr(year2$ICC, 1, 1) =="0" & !grepl("\\.| |-", year2$ICC), 
                                                 sub("^0+", "", year2$ICC), year2$ICC)
                             
                             year0 <- year0[which(!year0$`ICC`%in% excl_SKC$ICC  & !year0$`ICC`%in% excl_Cobranding$ICC  &  !year0$`ICC`%in% excl_Eurotec$ICC),]
                             year1 <- year1[which(!year1$`ICC`%in% excl_SKC$ICC  & !year1$`ICC`%in% excl_Cobranding$ICC  &  !year1$`ICC`%in% excl_Eurotec$ICC),]
                             year2 <- year2[which(!year2$`ICC`%in% excl_SKC$ICC  & !year2$`ICC`%in% excl_Cobranding$ICC  &  !year2$`ICC`%in% excl_Eurotec$ICC),]
                             
                             sales0 <- year0[,-c(1:3)]
                             sales1 <- year1[,-c(1:3)]
                             sales2 <- year2[,-c(1:3)]
                             sales0[1:(ncol(sales0)-1)] <- sapply(sales0[1:(ncol(sales0) - 1)], as.numeric)
                             sales1[1:(ncol(sales1)-1)] <- sapply(sales1[1:(ncol(sales1) - 1)], as.numeric)
                             sales2[1:(ncol(sales2)-1)] <- sapply(sales2[1:(ncol(sales2) - 1)], as.numeric)
    
                             tb_s0 <- aggregate(sales0[1], by = list(sales0$Region), sum)
                             for (i in 2:(ncol(sales0)-1)){
                             each <- aggregate(sales0[i], by = list(sales0$Region), sum)
                             tb_s0 <- cbind(tb_s0,each[,2])
                              }
    
                             tb_s1 <- aggregate(sales1[1], by = list(sales1$Region), sum)
                             for (i in 2:(ncol(sales1)-1)){
                             each <- aggregate(sales1[i], by = list(sales1$Region), sum)
                             tb_s1 <- cbind(tb_s1,each[,2])
                              }
    
                             tb_s2 <- aggregate(sales2[1], by = list(sales2$Region), sum)
                             for (i in 2:(ncol(sales2)-1)){
                             each <- aggregate(sales2[i], by = list(sales2$Region), sum)
                             tb_s2 <- cbind(tb_s2,each[,2])
                              }
    
                             sales <- cbind(tb_s2, tb_s1[,-1], tb_s0[,-1])
                             n <- length(seq(from = input$YTD[1] - years(2), to = input$YTD[2], by='month')) 
                             sales <- sales[,1:(n + 1)]
                             names(sales) <- c("Region2", seq(as.Date(input$YTD[1] - years(2)), by = "month", length.out = n))
                             return(sales)
                              })
  
  sales_region2_download <- reactive({if(is.null(sales_region2()))return(NULL) 
                                      sales <- sales_region2()
                                      n <- length(seq(from = input$YTD[1] - years(2), to = input$YTD[2], by='month')) 
                                      names(sales) <- c("BU", as.character(seq(as.Date(input$YTD[1] - years(2)), by = "month", length.out = n)))
                                      return(sales)
  })
  
  sales1 <- reactive({if(is.null(sales()))return(NULL)
                     sales <- sales()
                     colnames(sales)[1] <- "BU"
                     sales1 <- sales[,-1] # delete col 'BU'
                     sales1 <- as.list(sales1)
                     sales1 <- data.frame(matrix(unlist(sales1), nrow = length(sales1), byrow=T))
                     #sales1 <- data.frame(t(as.matrix(sales1)))
                     colnames(sales1) <- sales$BU
                     rownames(sales1) <- NULL
                     date <- seq(as.Date(input$YTD[1] - years(3)), by = "month", length.out = nrow(sales1))
                     sales1$month <- as.Date(as.yearmon(date))
                     return(sales1)
                    })
  
  sales1_region <- reactive({if(is.null(sales_region()))return(NULL)
                             sales <- sales_region()
                             sales1 <- sales[,-1] # delete col 'Region'
                             sales1 <- as.list(sales1)
                             sales1 <- data.frame(matrix(unlist(sales1), nrow = length(sales1), byrow=T))
                             #sales1 <- data.frame(t(as.matrix(sales1)))
                             colnames(sales1) <- sales$Region
                             rownames(sales1) <- NULL
                             date <- seq(as.Date(input$YTD[1] - years(2)), by = "month", length.out = nrow(sales1))
                             sales1$month <- as.Date(as.yearmon(date))
                             return(sales1)
  })
  
  sales1_region2 <- reactive({if(is.null(sales_region2()))return(NULL)
                             sales <- sales_region2()
                             sales1 <- sales[,-1] # delete col 'Region'
                             sales1 <- as.list(sales1)
                             sales1 <- data.frame(matrix(unlist(sales1), nrow = length(sales1), byrow=T))
                             #sales1 <- data.frame(t(as.matrix(sales1)))
                             colnames(sales1) <- sales$Region2
                             rownames(sales1) <- NULL
                             date <- seq(as.Date(input$YTD[1] - years(2)), by = "month", length.out = nrow(sales1))
                             sales1$month <- as.Date(as.yearmon(date))
                             return(sales1)
  })

  
  sales6 <- reactive({if(is.null(sales1()))return(NULL)
                      sales1 <- sales1()[,-ncol(sales1())]
                      sales6 <- data.frame()
                      n <- nrow(sales1)
                      for (j in 1:ncol(sales1)){
                           for(i in 1:(n-5)){sales6[i,j] <- round(colMeans(sales1[i:(i+5),]), 1)[j]}
                           }
                      colnames(sales6) <- colnames(sales1)
                      date <- seq(as.Date(input$YTD[1] - years(3)), by = "month", length.out = nrow(sales1) + 1)
                      sales6$month <- as.Date(as.yearmon(date[-c(1:6)]))
                      return(sales6)
    })
  
  sales6_region <- reactive({if(is.null(sales1_region()))return(NULL)
                   sales1 <- sales1_region()[,-ncol(sales1_region())]
                   sales6 <- data.frame()
                   n <- nrow(sales1)
                   for (j in 1:ncol(sales1)){
                       for(i in 1:(n-5)){sales6[i,j] <- round(colMeans(sales1[i:(i+5),]), 1)[j]}
                   }
                   colnames(sales6) <- colnames(sales1)
                   date <- seq(as.Date(input$YTD[1] - years(2)), by = "month", length.out = nrow(sales1) + 1)
                   sales6$month <- as.Date(as.yearmon(date[-c(1:6)]))
                   return(sales6)
  })
  
  sales6_region2 <- reactive({if(is.null(sales1_region2()))return(NULL)
                              sales1 <- sales1_region2()[,-ncol(sales1_region2())]
                              sales6 <- data.frame()
                              n <- nrow(sales1)
                              for (j in 1:ncol(sales1)){
                              for(i in 1:(n-5)){sales6[i,j] <- round(colMeans(sales1[i:(i+5),]), 1)[j]}
                              }
                              colnames(sales6) <- colnames(sales1)
                              date <- seq(as.Date(input$YTD[1] - years(2)), by = "month", length.out = nrow(sales1) + 1)
                              sales6$month <- as.Date(as.yearmon(date[-c(1:6)]))
                              return(sales6)
  })
  
  sales_avg6 <- reactive({if(is.null(sales6()))return(NULL)
                          sales6 <- sales6()
                          sales_avg6 <- data.frame('month' = rep(sales6$month, 6), 
                                                   'BU' = c(rep("Continence", nrow(sales6)), 
                                                            rep("Critical", nrow(sales6)),
                                                            rep("OSTOMY", nrow(sales6)), 
                                                            rep("WOUND", nrow(sales6)), 
                                                            rep("IC", nrow(sales6)),
                                                            rep("Eurotec", nrow(sales6))),
                                                   'sales' = rbind(as.matrix(sales6$Continence), 
                                                                   as.matrix(sales6$Critical),
                                                                   as.matrix(sales6$OSTOMY), 
                                                                   as.matrix(sales6$WOUND), 
                                                                   as.matrix(sales6$IC),
                                                                   as.matrix(sales6$Eurotec)))
                          return(sales_avg6)
  })
  
  sales_region_avg6 <- reactive({if(is.null(sales6_region()))return(NULL)
                                 sales6 <- sales6_region()
                                 sales_avg6 <- data.frame('month' = rep(sales6$month, 3), 
                                                          'Region' = c(rep("GEM", nrow(sales6)), 
                                                                       rep("EUR", nrow(sales6)),
                                                                       rep("NAM", nrow(sales6))),
                                                          'sales' = rbind(as.matrix(sales6$GEM), 
                                                                          as.matrix(sales6$EUR),
                                                                          as.matrix(sales6$NAM)
                                                 ))
                                 return(sales_avg6)
  })
  
  sales_region_avg62 <- reactive({if(is.null(sales6_region2()))return(NULL)
                                  sales6 <- sales6_region2()
                                  sales_avg6 <- data.frame('month' = rep(sales6$month, 3), 
                                                           'Region2' = c(rep("AP", nrow(sales6)), 
                                                                         rep("LA", nrow(sales6)),
                                                                         rep("MEA", nrow(sales6))),
                                                           'sales' = rbind(as.matrix(sales6$AP), 
                                                                           as.matrix(sales6$LA),
                                                                           as.matrix(sales6$MEA)
                                            ))
                                   return(sales_avg6)
  })

  
  CPM <- reactive({if(is.null(sales_avg6())|is.null(complaint_BU()))return(NULL)
                          sales_avg6 <- sales_avg6()
                          complaint <- complaint_BU()
                          sales <- sales_avg6[(sales_avg6$month >= "2018-01-01" & sales_avg6$month <= input$YTD[2]),]
                          both_monthly <- merge(x = complaint, y = sales, by = c("month", "BU"), all = T)
                          both_monthly$CPM <- round((both_monthly$complaint/both_monthly$sales)*1000000, 1)
                          both_monthly <- both_monthly[which(both_monthly$month >= min(sales_avg6$month)),]
                          return(both_monthly)
                          })
  
  CPM_region <- reactive({if(is.null(sales_region_avg6())|is.null(complaint_region()))return(NULL)
                          sales_avg6 <- sales_region_avg6()
                          complaint <- complaint_region()
                          sales <- sales_avg6[(sales_avg6$month >= "2019-01-01" & sales_avg6$month <= input$YTD[2]),]
                          both_monthly <- merge(x = complaint, y = sales, by = c("month", "Region"), all = T)
                          both_monthly$CPM <- round((both_monthly$complaint/both_monthly$sales)*1000000, 1)
                          both_monthly <- both_monthly[which(both_monthly$month >= min(sales_avg6$month)),]
                          return(both_monthly)
  })
  
  CPM_region2 <- reactive({if(is.null(sales_region_avg62())|is.null(complaint_region2()))return(NULL)
                           sales_avg6 <- sales_region_avg62()
                           complaint <- complaint_region2()
                           sales <- sales_avg6[(sales_avg6$month >= "2019-01-01" & sales_avg6$month <= input$YTD[2]),]
                           both_monthly <- merge(x = complaint, y = sales, by = c("month", "Region2"), all = T)
                           both_monthly$CPM <- round((both_monthly$complaint/both_monthly$sales)*1000000, 1)
                           both_monthly <- both_monthly[which(both_monthly$month >= min(sales_avg6$month)),]
                           return(both_monthly)
  })
  
  CPM_monthly <- reactive({if(is.null(CPM()))return(NULL)
                  both_monthly <- CPM()
                  comp_CPM <- data.frame(Month = unique(both_monthly$month),
                                         AWC.RAW.CPLT = both_monthly[both_monthly$BU == "WOUND",]$complaint,
                                         CC.RAW.CPLT = both_monthly[both_monthly$BU == "Continence",]$complaint,
                                         C.RAW.CPLT = both_monthly[both_monthly$BU == "Critical",]$complaint,
                                         OC.RAW.CPLT = both_monthly[both_monthly$BU == "OSTOMY",]$complaint,
                                         IC.RAW.CPLT = both_monthly[both_monthly$BU == "IC",]$complaint,
                                         Eurotec.RAW.CPLT = both_monthly[both_monthly$BU == "Eurotec",]$complaint,
                                         AWC.CPM = both_monthly[both_monthly$BU == "WOUND",]$CPM,
                                         CC.CPM = both_monthly[both_monthly$BU == "Continence",]$CPM,
                                         C.CPM = both_monthly[both_monthly$BU == "Critical",]$CPM,
                                         OC.CPM = both_monthly[both_monthly$BU == "OSTOMY",]$CPM,
                                         IC.CPM = both_monthly[both_monthly$BU == "IC",]$CPM,
                                         Eurotec.CPM = both_monthly[both_monthly$BU == "Eurotec",]$CPM)
                                         comp_CPM$Month <- as.yearmon(comp_CPM$Month)
                  comp_CPM <- comp_CPM[-c(1:6),]
                  return(comp_CPM)
  })
  
  CPM_region_monthly <- reactive({if(is.null(CPM_region()))return(NULL)
                                  both_monthly <- CPM_region()
                                  comp_CPM <- data.frame(Month = unique(both_monthly$month),
                                                         GEM.RAW.CPLT = both_monthly[both_monthly$Region == "GEM",]$complaint,
                                                         EUR.RAW.CPLT = both_monthly[both_monthly$Region == "EUR",]$complaint,
                                                         NAM.RAW.CPLT = both_monthly[both_monthly$Region == "NAM",]$complaint,
                                                         GEM.CPM = both_monthly[both_monthly$Region == "GEM",]$CPM,
                                                         EUR.CPM = both_monthly[both_monthly$Region == "EUR",]$CPM,
                                                         NAM.CPM = both_monthly[both_monthly$Region == "NAM",]$CPM)
                                                         comp_CPM$Month <- as.yearmon(comp_CPM$Month)
                                                         comp_CPM <- comp_CPM[-c(1:6),]
                                                         return(comp_CPM)
  })
  
  CPM_region_monthly2 <- reactive({if(is.null(CPM_region2()))return(NULL)
                                   both_monthly <- CPM_region2()
                                   comp_CPM <- data.frame(Month = unique(both_monthly$month),
                                   AP.RAW.CPLT = both_monthly[both_monthly$Region2 == "AP",]$complaint,
                                   LA.RAW.CPLT = both_monthly[both_monthly$Region2 == "LA",]$complaint,
                                   MEA.RAW.CPLT = both_monthly[both_monthly$Region2 == "MEA",]$complaint,
                                   AP.CPM = both_monthly[both_monthly$Region2 == "AP",]$CPM,
                                   LA.CPM = both_monthly[both_monthly$Region2 == "LA",]$CPM,
                                   MEA.CPM = both_monthly[both_monthly$Region2 == "MEA",]$CPM)
                                   comp_CPM$Month <- as.yearmon(comp_CPM$Month)
                                   comp_CPM <- comp_CPM[-c(1:6),]
                                   return(comp_CPM)
  })
  
  reduction_ALL <- reactive({if(is.null(CPM()))return(NULL)
                             CPM_monthly <- CPM()
                             
                             pastyear1 <- format((input$YTD[1] - years(1)), "%Y")
                             complaint_yearly <- aggregate(CPM_monthly$complaint, 
                                                           by = list(format(CPM_monthly$month, "%Y"), CPM_monthly$BU),
                                                           sum)
                             names(complaint_yearly) <- c("Year", "BU", "Complaint")
                             complaint_yearly <- complaint_yearly[complaint_yearly$BU!="Eurotec",]
                             IC <- complaint_yearly[which(complaint_yearly$Year==pastyear1 & complaint_yearly$BU=="IC"),]$Complaint * as.numeric(input$reduction_IC)
                             OC <- complaint_yearly[which(complaint_yearly$Year==pastyear1 & complaint_yearly$BU=="OSTOMY"),]$Complaint * as.numeric(input$reduction_OC)
                             CC <- complaint_yearly[which(complaint_yearly$Year==pastyear1 & complaint_yearly$BU=="Continence"),]$Complaint * as.numeric(input$reduction_CC)
                             C <- complaint_yearly[which(complaint_yearly$Year==pastyear1 & complaint_yearly$BU=="Critical"),]$Complaint * as.numeric(input$reduction_C)
                             AWC <- complaint_yearly[which(complaint_yearly$Year==pastyear1 & complaint_yearly$BU=="WOUND"),]$Complaint * as.numeric(input$reduction_AWC)
                             ALL <- aggregate(complaint_yearly$Complaint, 
                                              by = list(complaint_yearly$Year),
                                              sum)
                             names(ALL) <- c("Year", "Complaint")
                             ALL <- ALL[which(ALL$Year==pastyear1),]$Complaint
                             reduction <- round(sum(IC, OC, CC, C, AWC)/ALL,2)
                             return(reduction)
  })
  
  reduction_NIC <- reactive({if(is.null(CPM()))return(NULL)
                             CPM_monthly <- CPM()
                             pastyear1 <- format((input$YTD[1] - years(1)), "%Y")
                             complaint_yearly <- aggregate(CPM_monthly$complaint, 
                                                           by = list(format(CPM_monthly$month, "%Y"), CPM_monthly$BU),
                                                           sum)
                             names(complaint_yearly) <- c("Year", "BU", "Complaint")
                             complaint_yearly <- complaint_yearly[which(complaint_yearly$BU!="IC" & complaint_yearly$BU!="Eurotec"),]
                             OC <- complaint_yearly[which(complaint_yearly$Year==pastyear1 & complaint_yearly$BU=="OSTOMY"),]$Complaint * as.numeric(input$reduction_OC)
                             CC <- complaint_yearly[which(complaint_yearly$Year==pastyear1 & complaint_yearly$BU=="Continence"),]$Complaint * as.numeric(input$reduction_CC)
                             C <- complaint_yearly[which(complaint_yearly$Year==pastyear1 & complaint_yearly$BU=="Critical"),]$Complaint * as.numeric(input$reduction_C)
                             AWC <- complaint_yearly[which(complaint_yearly$Year==pastyear1 & complaint_yearly$BU=="WOUND"),]$Complaint * as.numeric(input$reduction_AWC)
                             ALL <- aggregate(complaint_yearly$Complaint, 
                                              by = list(complaint_yearly$Year),
                                              sum)
                             names(ALL) <- c("Year", "Complaint")
                             ALL <- ALL[which(ALL$Year==pastyear1),]$Complaint
                             reduction <- round(sum(OC, CC, C, AWC)/ALL, 2)
                             return(reduction)
  })
  
  reduction_region_GEM <- reactive({if(is.null(CPM_region2()))return(NULL)
                                    CPM_monthly <- CPM_region2()
    
                                    pastyear1 <- format((input$YTD[1] - years(1)), "%Y")
                                    complaint_yearly <- aggregate(CPM_monthly$complaint, 
                                                                  by = list(format(CPM_monthly$month, "%Y"), CPM_monthly$Region2),
                                                                  sum)
                                    names(complaint_yearly) <- c("Year", "Region2", "Complaint")
                                    AP <- complaint_yearly[which(complaint_yearly$Year==pastyear1 & complaint_yearly$Region2=="AP"),]$Complaint * as.numeric(input$reduction_AP)
                                    LA <- complaint_yearly[which(complaint_yearly$Year==pastyear1 & complaint_yearly$Region2=="LA"),]$Complaint * as.numeric(input$reduction_LA)
                                    MEA <- complaint_yearly[which(complaint_yearly$Year==pastyear1 & complaint_yearly$Region2=="MEA"),]$Complaint * as.numeric(input$reduction_MEA)
    
                                    GEM <- aggregate(complaint_yearly$Complaint, 
                                                     by = list(complaint_yearly$Year),
                                                     sum)
                                    names(GEM) <- c("Year", "Complaint")
                                    GEM <- GEM[which(GEM$Year==pastyear1),]$Complaint 
                                    reduction <- round(sum(AP, LA, MEA)/GEM,2)
                                    return(reduction)
  })
  
  reduction_region_ALL <- reactive({if(is.null(CPM_region()))return(NULL)
                                    CPM_monthly <- CPM_region()
    
                                    pastyear1 <- format((input$YTD[1] - years(1)), "%Y")
                                    complaint_yearly <- aggregate(CPM_monthly$complaint, 
                                                                  by = list(format(CPM_monthly$month, "%Y"), CPM_monthly$Region),
                                                                  sum)
                                    names(complaint_yearly) <- c("Year", "Region", "Complaint")
                                    GEM <- complaint_yearly[which(complaint_yearly$Year==pastyear1 & complaint_yearly$Region=="GEM"),]$Complaint * as.numeric(reduction_region_GEM())
                                    EUR <- complaint_yearly[which(complaint_yearly$Year==pastyear1 & complaint_yearly$Region=="EUR"),]$Complaint * as.numeric(input$reduction_EUR)
                                    NAM <- complaint_yearly[which(complaint_yearly$Year==pastyear1 & complaint_yearly$Region=="NAM"),]$Complaint * as.numeric(input$reduction_NAM)
    
                                    ALL <- aggregate(complaint_yearly$Complaint, 
                                                     by = list(complaint_yearly$Year),
                                                     sum)
                                    names(ALL) <- c("Year", "Complaint")
                                    ALL <- ALL[which(ALL$Year==pastyear1),]$Complaint 
                                    reduction <- round(sum(GEM, EUR, NAM)/ALL,2)
                                    return(reduction)
  })
  
   CPM_yearly <- reactive({if(is.null(CPM()))return(NULL)
                           CPM_monthly <- CPM()
                           CPM_yearly <- aggregate(CPM_monthly$CPM, 
                                                   by = list(format(CPM_monthly$month, "%Y"), CPM_monthly$BU),
                                                   mean)
                           complaint_yearly <- aggregate(CPM_monthly$complaint, 
                                                         by = list(format(CPM_monthly$month, "%Y"), CPM_monthly$BU),
                                                         sum)
                           colnames(complaint_yearly) <- c("Year", "BU", "complaints")
                           colnames(CPM_yearly) <- c("Year", "BU", "CPM")
                           both_yearly <- merge(x = complaint_yearly, y = CPM_yearly, by = c("Year", "BU"), all = T)
                           colnames(both_yearly) <- c("Year", "BU", "complaints", "CPM")
                           both_yearly$CPM <- round(both_yearly$CPM, 1)
                           comp_CPM_yearly <- data.frame(Year = unique(both_yearly$Year),
                                                         AWC.RAW.CPLT = both_yearly[both_yearly$BU == "WOUND",]$complaints,
                                                         CC.RAW.CPLT = both_yearly[both_yearly$BU == "Continence",]$complaints,
                                                         C.RAW.CPLT = both_yearly[both_yearly$BU == "Critical",]$complaints,
                                                         OC.RAW.CPLT = both_yearly[both_yearly$BU == "OSTOMY",]$complaints,
                                                         IC.RAW.CPLT = both_yearly[both_yearly$BU == "IC",]$complaints,
                                                         Eurotec.RAW.CPLT = both_yearly[both_yearly$BU == "Eurotec",]$complaints,
                                                         AWC.CPM = both_yearly[both_yearly$BU == "WOUND",]$CPM,
                                                         CC.CPM = both_yearly[both_yearly$BU == "Continence",]$CPM,
                                                         C.CPM = both_yearly[both_yearly$BU == "Critical",]$CPM,
                                                         OC.CPM = both_yearly[both_yearly$BU == "OSTOMY",]$CPM,
                                                         IC.CPM = both_yearly[both_yearly$BU == "IC",]$CPM,
                                                         Eurotec.CPM = both_yearly[both_yearly$BU == "Eurotec",]$CPM)
                           comp_CPM_yearly <- comp_CPM_yearly[-1,]
                           comp_CPM_yearly$Year <- as.character(comp_CPM_yearly$Year)
                           comp_CPM_yearly[nrow(comp_CPM_yearly),]$Year <- paste0(format((input$YTD[1]), "%Y")," YTD")
                           return(comp_CPM_yearly)
   })
   
   CPM_region_yearly <- reactive({if(is.null(CPM_region()))return(NULL)
                                  CPM_monthly <- CPM_region()
                                  CPM_yearly <- aggregate(CPM_monthly$CPM, 
                                                          by = list(format(CPM_monthly$month, "%Y"), CPM_monthly$Region),
                                                          mean)
                                  complaint_yearly <- aggregate(CPM_monthly$complaint, 
                                                                by = list(format(CPM_monthly$month, "%Y"), CPM_monthly$Region),
                                                                sum)
                                  colnames(complaint_yearly) <- c("Year", "Region", "complaints")
                                  colnames(CPM_yearly) <- c("Year", "Region", "CPM")
                                  both_yearly <- merge(x = complaint_yearly, y = CPM_yearly, by = c("Year", "Region"), all = T)
                                  colnames(both_yearly) <- c("Year", "Region", "complaints", "CPM")
                                  both_yearly$CPM <- round(both_yearly$CPM, 1)
                                  comp_CPM_yearly <- data.frame(Year = unique(both_yearly$Year),
                                                                GEM.RAW.CPLT = both_yearly[both_yearly$Region == "GEM",]$complaints,
                                                                EUR.RAW.CPLT = both_yearly[both_yearly$Region == "EUR",]$complaints,
                                                                NAM.RAW.CPLT = both_yearly[both_yearly$Region == "NAM",]$complaints,
                                                                GEM.CPM = both_yearly[both_yearly$Region == "GEM",]$CPM,
                                                                EUR.CPM = both_yearly[both_yearly$Region == "EUR",]$CPM,
                                                                NAM.CPM = both_yearly[both_yearly$Region == "NAM",]$CPM)
                                  comp_CPM_yearly <- comp_CPM_yearly[-1,] # remove year2
                                  comp_CPM_yearly$Year <- as.character(comp_CPM_yearly$Year)
                                  comp_CPM_yearly[nrow(comp_CPM_yearly),]$Year <- paste0(format((input$YTD[1]), "%Y")," YTD")
                                  return(comp_CPM_yearly)
   })
   
   CPM_region_yearly2 <- reactive({if(is.null(CPM_region2()))return(NULL)
                                   CPM_monthly <- CPM_region2()
                                   CPM_yearly <- aggregate(CPM_monthly$CPM, 
                                                           by = list(format(CPM_monthly$month, "%Y"), CPM_monthly$Region2),
                                                           mean)
                                   complaint_yearly <- aggregate(CPM_monthly$complaint, 
                                                                 by = list(format(CPM_monthly$month, "%Y"), CPM_monthly$Region2),
                                                                 sum)
                                   colnames(complaint_yearly) <- c("Year", "Region2", "complaints")
                                   colnames(CPM_yearly) <- c("Year", "Region2", "CPM")
                                   both_yearly <- merge(x = complaint_yearly, y = CPM_yearly, by = c("Year", "Region2"), all = T)
                                   colnames(both_yearly) <- c("Year", "Region2", "complaints", "CPM")
                                   both_yearly$CPM <- round(both_yearly$CPM, 1)
                                   comp_CPM_yearly <- data.frame(Year = unique(both_yearly$Year),
                                   AP.RAW.CPLT = both_yearly[both_yearly$Region2 == "AP",]$complaints,
                                   LA.RAW.CPLT = both_yearly[both_yearly$Region2 == "LA",]$complaints,
                                   MEA.RAW.CPLT = both_yearly[both_yearly$Region2 == "MEA",]$complaints,
                                   AP.CPM = both_yearly[both_yearly$Region2 == "AP",]$CPM,
                                   LA.CPM = both_yearly[both_yearly$Region2 == "LA",]$CPM,
                                   MEA.CPM = both_yearly[both_yearly$Region2 == "MEA",]$CPM)
                                   comp_CPM_yearly <- comp_CPM_yearly[-1,] # remove year2
                                   comp_CPM_yearly$Year <- as.character(comp_CPM_yearly$Year)
                                   comp_CPM_yearly[nrow(comp_CPM_yearly),]$Year <- paste0(format((input$YTD[1]), "%Y")," YTD")
                                   return(comp_CPM_yearly)
   })
   
   CPM_monthly_NIC <- reactive({if(is.null(sales_avg6())|is.null(complaint_NIC()))return(NULL)
                                sales_avg6 <- sales_avg6()
                                complaint <- complaint_NIC()
                                sales_avg6 <- sales_avg6[which(sales_avg6$BU!="IC" & sales_avg6$BU!="Eurotec"),]
                                sales_avg6 <- sales_avg6[sales_avg6$month >= "2018-01-01" & sales_avg6$month <= input$YTD[2],]
                                sales_monthly <- aggregate(sales_avg6$sales, 
                                                           by = list(sales_avg6$month),
                                                           sum)
                                colnames(sales_monthly) <- c("month", "sales")
                                both_monthly <- merge(x = complaint, y = sales_monthly, by = c("month"), all = T)
                                both_monthly$CPM <- round((both_monthly$complaint/both_monthly$sales)*1000000, 1)
                                both_monthly$sales <- round(both_monthly$sales, 0)
                                colnames(both_monthly) <- c("Month", "complaint", "6M_avg_sales", "CPM")
                                both_monthly$Month <- as.yearmon(both_monthly$Month)
                                both_monthly <- both_monthly[-c(1:12),]
                                return(both_monthly) 
   })
   
   
   
   CPM_yearly_NIC <- reactive({if(is.null(CPM_monthly_NIC()))return(NULL)
                               CPM_monthly <- CPM_monthly_NIC()
                               CPM_yearly <- aggregate(CPM_monthly$CPM, 
                                                       by = list(format(CPM_monthly$Month, "%Y")),
                                                       mean)
                               complaint_yearly <- aggregate(CPM_monthly$complaint, 
                                                             by = list(format(CPM_monthly$Month, "%Y")),
                                                             sum)
                               sales_yearly <- aggregate(CPM_monthly$`6M_avg_sales`, 
                                                         by = list(format(CPM_monthly$Month, "%Y")),
                                                         sum)
                               colnames(complaint_yearly) <- c("Year", "complaints")
                               colnames(CPM_yearly) <- c("Year", "CPM")
                               colnames(sales_yearly) <- c("Year", "6M_avg_sales")
                               both_yearly <- merge(x = complaint_yearly, y = CPM_yearly, by = c("Year"), all = T)
                               both_yearly <- merge(x = both_yearly, y = sales_yearly, by = c("Year"), all = T)
                               both_yearly$CPM <- round(both_yearly$CPM, 1)
                               both_yearly$`6M_avg_sales` <- round(both_yearly$`6M_avg_sales`, 0)
                               both_yearly <- both_yearly[c("Year", "complaints", "6M_avg_sales", "CPM")]
                               both_yearly[nrow(both_yearly),]$Year <- paste0(format((input$YTD[1]), "%Y")," YTD")
                               return(both_yearly)
   })
   
   CPM_monthly_all <- reactive({if(is.null(sales_avg6())|is.null(complaint_all()))return(NULL)
                                sales_avg6 <- sales_avg6()
                                complaint <- complaint_all()
                                sales_avg6 <- sales_avg6[which(sales_avg6$BU!="Eurotec"),]
                                sales_avg6 <- sales_avg6[sales_avg6$month >= "2018-01-01" & sales_avg6$month <= input$YTD[2],]
                                sales_monthly <- aggregate(sales_avg6$sales, 
                                                           by = list(sales_avg6$month),
                                                           sum)
                                colnames(sales_monthly) <- c("month", "sales")
                                both_monthly <- merge(x = complaint, y = sales_monthly, by = c("month"), all = T)
                                both_monthly$CPM <- round((both_monthly$complaint/both_monthly$sales)*1000000, 1)
                                both_monthly$sales <- round(both_monthly$sales, 0)
                                colnames(both_monthly) <- c("Month", "complaint", "6M_avg_sales", "CPM")
                                both_monthly <- both_monthly[both_monthly$Month >= as.Date(input$YTD[1] - years(2)),]
                                both_monthly$Month <- as.yearmon(both_monthly$Month)
                                #both_monthly <- both_monthly[-c(1:6),]
                                return(both_monthly) 
   })
   
   CPM_yearly_all <- reactive({if(is.null(CPM_monthly_all()))return(NULL)
                               CPM_monthly <- CPM_monthly_all()
                               CPM_yearly <- aggregate(CPM_monthly$CPM, 
                                                       by = list(format(CPM_monthly$Month, "%Y")),
                                                       mean)
                               complaint_yearly <- aggregate(CPM_monthly$complaint, 
                                                       by = list(format(CPM_monthly$Month, "%Y")),
                                                       sum)
                               sales_yearly <- aggregate(CPM_monthly$`6M_avg_sales`, 
                                                         by = list(format(CPM_monthly$Month, "%Y")),
                                                         sum)
                               colnames(complaint_yearly) <- c("Year", "complaints")
                               colnames(CPM_yearly) <- c("Year", "CPM")
                               colnames(sales_yearly) <- c("Year", "6M_avg_sales")
                               both_yearly <- merge(x = complaint_yearly, y = CPM_yearly, by = c("Year"), all = T)
                               both_yearly <- merge(x = both_yearly, y = sales_yearly, by = c("Year"), all = T)
                               both_yearly$CPM <- round(both_yearly$CPM, 1)
                               both_yearly$`6M_avg_sales` <- round(both_yearly$`6M_avg_sales`, 0)
                               both_yearly <- both_yearly[c("Year", "complaints", "6M_avg_sales", "CPM")]
                               both_yearly[nrow(both_yearly),]$Year <- paste0(format((input$YTD[1]), "%Y")," YTD")
                               return(both_yearly)
   })
   
   OC <- reactive({if(is.null(CPM_yearly())|is.null(CPM_monthly()))return(NULL)
                    CPM_monthly <- CPM_monthly()
                    CPM_yearly <- CPM_yearly()
                    validate(need(input$YTD[2]-input$YTD[1] < 365,
                                  "Warning Message: Selected Date Range should be less than 12 months."))
                    reduction <- as.numeric(input$reduction_OC)
                    OC <- data.frame()
                    
                    pastyear2 <- format((input$YTD[1] - years(2)), "%Y")
                    pastyear1 <- format((input$YTD[1] - years(1)), "%Y")
                    currentyear <- format(input$YTD[1], "%Y")
                    target <- round(CPM_yearly[which(CPM_yearly$Year==pastyear1),]$OC.CPM * reduction, 1)
                    OC[1:13, 1] <- c(as.character(as.yearmon(seq(input$YTD[1], by = "month", length.out = 12))), "YTD")
                    OC[1:13, 2] <- rep("", 13)
                    #OC[1:13, 3] <- c(CPM_yearly[which(CPM_yearly$Year==pastyear2),]$OC.CPM, rep("", 12))
                    OC[1:13, 3] <- c(CPM_yearly[which(CPM_yearly$Year==pastyear1),]$OC.CPM, rep("", 12))
                    OC[1:13, 4] <- c(target, rep("", 12))
                    
                    YTD_monthly <- CPM_monthly[which(CPM_monthly$Month >= as.yearmon(input$YTD[1]) & CPM_monthly$Month <= as.yearmon(input$YTD[2])),]
                    YTD_yearly <- CPM_yearly[which(CPM_yearly$Year==paste0(currentyear, " YTD")),]

                    OC[1:13, 5] <- c(if(nrow(YTD_monthly) == 12){YTD_monthly$OC.CPM
                                                            }else{c(YTD_monthly$OC.CPM, rep("-", (12-nrow(YTD_monthly))))},
                                      if(nrow(YTD_yearly) == 0){"-"}else{YTD_yearly$OC.CPM}
                                      )
                    OC[,5] <- as.character(OC[,5])
                    names(OC) <- c("na",
                                    "na1", 
                                    "pastyear1",
                                    "currentyear_target",
                                    "currentyear_actual") 
                    return(OC)
   })
   
   AWC <- reactive({if(is.null(CPM_yearly())|is.null(CPM_monthly()))return(NULL)
                    CPM_monthly <- CPM_monthly()
                    CPM_yearly <- CPM_yearly()
                    validate(need(input$YTD[2]-input$YTD[1] < 365,
                             "Warning Message: Selected Date Range should be less than 12 months."))
                    reduction <- as.numeric(input$reduction_AWC)
                    AWC <- data.frame()
     
                    #pastyear2 <- format((input$YTD[1] - years(2)), "%Y")
                    pastyear1 <- format((input$YTD[1] - years(1)), "%Y")
                    currentyear <- format(input$YTD[1], "%Y")
                    target <- round(CPM_yearly[which(CPM_yearly$Year==pastyear1),]$AWC.CPM * reduction, 1)
                    AWC[1:13, 1] <- c(as.character(as.yearmon(seq(input$YTD[1], by = "month", length.out = 12))), "YTD")
                    AWC[1:13, 2] <- rep("", 13)
                    #AWC[1:13, 3] <- c(CPM_yearly[which(CPM_yearly$Year==pastyear2),]$AWC.CPM, rep("", 12))
                    AWC[1:13, 3] <- c(CPM_yearly[which(CPM_yearly$Year==pastyear1),]$AWC.CPM, rep("", 12))
                    AWC[1:13, 4] <- c(target, rep("", 12))
     
                    YTD_monthly <- CPM_monthly[which(CPM_monthly$Month >= as.yearmon(input$YTD[1]) & CPM_monthly$Month <= as.yearmon(input$YTD[2])),]
                    YTD_yearly <- CPM_yearly[which(CPM_yearly$Year==paste0(currentyear, " YTD")),]
     
                    AWC[1:13, 5] <- c(if(nrow(YTD_monthly) == 12){YTD_monthly$AWC.CPM
                                     }else{c(YTD_monthly$AWC.CPM, rep("-", (12-nrow(YTD_monthly))))},
                                     if(nrow(YTD_yearly) == 0){"-"}else{YTD_yearly$AWC.CPM}
                                     )
                    AWC[,5] <- as.character(AWC[,5])
                    names(AWC) <- c("na",
                                    "na1", 
                                    "pastyear1",
                                    "currentyear_target",
                                    "currentyear_actual") 
                    return(AWC)
   })
   
   CC <- reactive({if(is.null(CPM_yearly())|is.null(CPM_monthly()))return(NULL)
                    CPM_monthly <- CPM_monthly()
                    CPM_yearly <- CPM_yearly()
                    validate(need(input$YTD[2]-input$YTD[1] < 365,
                      "Warning Message: Selected Date Range should be less than 12 months."))
                    reduction <- as.numeric(input$reduction_CC)
                    CC <- data.frame()
     
                    #pastyear2 <- format((input$YTD[1] - years(2)), "%Y")
                    pastyear1 <- format((input$YTD[1] - years(1)), "%Y")
                    currentyear <- format(input$YTD[1], "%Y")
                    target <- round(CPM_yearly[which(CPM_yearly$Year==pastyear1),]$CC.CPM * reduction, 1)
                    CC[1:13, 1] <- c(as.character(as.yearmon(seq(input$YTD[1], by = "month", length.out = 12))), "YTD")
                    CC[1:13, 2] <- rep("", 13)
                    #CC[1:13, 3] <- c(CPM_yearly[which(CPM_yearly$Year==pastyear2),]$CC.CPM, rep("", 12))
                    CC[1:13, 3] <- c(CPM_yearly[which(CPM_yearly$Year==pastyear1),]$CC.CPM, rep("", 12))
                    CC[1:13, 4] <- c(target, rep("", 12))
     
                    YTD_monthly <- CPM_monthly[which(CPM_monthly$Month >= as.yearmon(input$YTD[1]) & CPM_monthly$Month <= as.yearmon(input$YTD[2])),]
                    YTD_yearly <- CPM_yearly[which(CPM_yearly$Year==paste0(currentyear, " YTD")),]
     
                    CC[1:13, 5] <- c(if(nrow(YTD_monthly) == 12){YTD_monthly$CC.CPM
                                      }else{c(YTD_monthly$CC.CPM, rep("-", (12-nrow(YTD_monthly))))},
                                     if(nrow(YTD_yearly) == 0){"-"}else{YTD_yearly$CC.CPM}
                                     )
                    CC[,5] <- as.character(CC[,5])
                    names(CC) <- c("na",
                                   "na1",
                                   "pastyear1",
                                   "currentyear_target",
                                   "currentyear_actual") 
                    return(CC)
   })
   
   
   C <- reactive({if(is.null(CPM_yearly())|is.null(CPM_monthly()))return(NULL)
                  CPM_monthly <- CPM_monthly()
                  CPM_yearly <- CPM_yearly()
                  validate(need(input$YTD[2]-input$YTD[1] < 365,
                           "Warning Message: Selected Date Range should be less than 12 months."))
                  reduction <- as.numeric(input$reduction_C)
                  C <- data.frame()
     
                  #pastyear2 <- format((input$YTD[1] - years(2)), "%Y")
                  pastyear1 <- format((input$YTD[1] - years(1)), "%Y")
                  currentyear <- format(input$YTD[1], "%Y")
                  target <- round(CPM_yearly[which(CPM_yearly$Year==pastyear1),]$C.CPM * reduction, 1)
                  C[1:13, 1] <- c(as.character(as.yearmon(seq(input$YTD[1], by = "month", length.out = 12))), "YTD")
                  C[1:13, 2] <- rep("", 13)
                  #C[1:13, 3] <- c(CPM_yearly[which(CPM_yearly$Year==pastyear2),]$C.CPM, rep("", 12))
                  C[1:13, 3] <- c(CPM_yearly[which(CPM_yearly$Year==pastyear1),]$C.CPM, rep("", 12))
                  C[1:13, 4] <- c(target, rep("", 12))
     
                  YTD_monthly <- CPM_monthly[which(CPM_monthly$Month >= as.yearmon(input$YTD[1]) & CPM_monthly$Month <= as.yearmon(input$YTD[2])),]
                  YTD_yearly <- CPM_yearly[which(CPM_yearly$Year==paste0(currentyear, " YTD")),]
     
                  C[1:13, 5] <- c(if(nrow(YTD_monthly) == 12){YTD_monthly$C.CPM
                                                        }else{c(YTD_monthly$C.CPM, rep("-", (12-nrow(YTD_monthly))))},
                                                             if(nrow(YTD_yearly) == 0){"-"}else{YTD_yearly$C.CPM}
                                 )
                  C[,5] <- as.character(C[,5])
                  names(C) <- c("na",
                                "na1", 
                                "pastyear1",
                                "currentyear_target",
                                "currentyear_actual") 
                  return(C)
    })
   
   
   IC <- reactive({if(is.null(CPM_yearly())|is.null(CPM_monthly()))return(NULL)
                   CPM_monthly <- CPM_monthly()
                   CPM_yearly <- CPM_yearly()
                   validate(need(input$YTD[2]-input$YTD[1] < 365,
                            "Warning Message: Selected Date Range should be less than 12 months."))
                   reduction <- as.numeric(input$reduction_IC)
                   IC <- data.frame()
     
                   #pastyear2 <- format((input$YTD[1] - years(2)), "%Y")
                   pastyear1 <- format((input$YTD[1] - years(1)), "%Y")
                   currentyear <- format(input$YTD[1], "%Y")
                   target <- round(CPM_yearly[which(CPM_yearly$Year==pastyear1),]$IC.CPM * reduction, 1)
                   IC[1:13, 1] <- c(as.character(as.yearmon(seq(input$YTD[1], by = "month", length.out = 12))), "YTD")
                   IC[1:13, 2] <- rep("", 13)
                   #IC[1:13, 3] <- c(CPM_yearly[which(CPM_yearly$Year==pastyear2),]$IC.CPM, rep("", 12))
                   IC[1:13, 3] <- c(CPM_yearly[which(CPM_yearly$Year==pastyear1),]$IC.CPM, rep("", 12))
                   IC[1:13, 4] <- c(target, rep("", 12))
     
                   YTD_monthly <- CPM_monthly[which(CPM_monthly$Month >= as.yearmon(input$YTD[1]) & CPM_monthly$Month <= as.yearmon(input$YTD[2])),]
                   YTD_yearly <- CPM_yearly[which(CPM_yearly$Year==paste0(currentyear, " YTD")),]
     
                   IC[1:13, 5] <- c(if(nrow(YTD_monthly) == 12){YTD_monthly$IC.CPM
                                   }else{c(YTD_monthly$IC.CPM, rep("-", (12-nrow(YTD_monthly))))},
                                    if(nrow(YTD_yearly) == 0){"-"}else{YTD_yearly$IC.CPM}
                                    )
                   IC[,5] <- as.character(IC[,5])
                   names(IC) <- c("na",
                                  "na1",
                                  "pastyear1",
                                  "currentyear_target",
                                  "currentyear_actual") 
                   return(IC)
   })
   
   Eurotec <- reactive({if(is.null(CPM_yearly())|is.null(CPM_monthly()))return(NULL)
                         CPM_monthly <- CPM_monthly()
                         CPM_yearly <- CPM_yearly()
                         reduction <- as.numeric(input$reduction_Eurotec)
                         Eurotec <- data.frame()
     
                        #pastyear2 <- format((input$YTD[1] - years(2)), "%Y")
                         pastyear1 <- format((input$YTD[1] - years(1)), "%Y")
                         currentyear <- format(input$YTD[1], "%Y")
                         target <- round(CPM_yearly[which(CPM_yearly$Year==pastyear1),]$Eurotec.CPM * reduction, 1)
                         Eurotec[1:13, 1] <- c(as.character(as.yearmon(seq(input$YTD[1], by = "month", length.out = 12))), "YTD")
                         Eurotec[1:13, 2] <- rep("", 13)
                        #Eurotec[1:13, 3] <- c(CPM_yearly[which(CPM_yearly$Year==pastyear2),]$Eurotec.CPM, rep("", 12))
                         Eurotec[1:13, 3] <- c(CPM_yearly[which(CPM_yearly$Year==pastyear1),]$Eurotec.CPM, rep("", 12))
                         Eurotec[1:13, 4] <- c(target, rep("", 12))
     
                         YTD_monthly <- CPM_monthly[which(CPM_monthly$Month >= as.yearmon(input$YTD[1]) & CPM_monthly$Month <= as.yearmon(input$YTD[2])),]
                         YTD_yearly <- CPM_yearly[which(CPM_yearly$Year==paste0(currentyear, " YTD")),]
     
                         Eurotec[1:13, 5] <- c(if(nrow(YTD_monthly) == 12){YTD_monthly$Eurotec.CPM
                                            }else{c(YTD_monthly$Eurotec.CPM, rep("-", (12-nrow(YTD_monthly))))},
                                                  if(nrow(YTD_yearly) == 0){"-"}else{YTD_yearly$Eurotec.CPM}
                           )
                         Eurotec[,5] <- as.character(Eurotec[,5])
                         names(Eurotec) <- c("na",
                                             "na1",
                                             "pastyear1",
                                             "currentyear_target",
                                             "currentyear_actual") 
                                              return(Eurotec)
   })
   
   NIC <- reactive({if(is.null(CPM_yearly_NIC())|is.null(CPM_monthly_NIC()))return(NULL)
                    CPM_monthly <- CPM_monthly_NIC()
                    CPM_yearly <- CPM_yearly_NIC()
                    validate(need(input$YTD[2]-input$YTD[1] < 365,
                            "Warning Message: Selected Date Range should be less than 12 months."))
                    reduction <- as.numeric(reduction_NIC())
                    NIC <- data.frame()
     
                    #pastyear2 <- format((input$YTD[1] - years(2)), "%Y")
                    pastyear1 <- format((input$YTD[1] - years(1)), "%Y")
                    currentyear <- format(input$YTD[1], "%Y")
                    target <- round(CPM_yearly[which(CPM_yearly$Year==pastyear1),]$CPM * reduction, 1)
                    NIC[1:13, 1] <- c(as.character(as.yearmon(seq(input$YTD[1], by = "month", length.out = 12))), "YTD")
                    NIC[1:13, 2] <- rep("", 13)
                    #NIC[1:13, 3] <- c(CPM_yearly[which(CPM_yearly$Year==pastyear2),]$CPM, rep("", 12))
                    NIC[1:13, 3] <- c(CPM_yearly[which(CPM_yearly$Year==pastyear1),]$CPM, rep("", 12))
                    NIC[1:13, 4] <- c(target, rep("", 12))
     
                    YTD_monthly <- CPM_monthly[which(CPM_monthly$Month >= as.yearmon(input$YTD[1]) & CPM_monthly$Month <= as.yearmon(input$YTD[2])),]
                    YTD_yearly <- CPM_yearly[which(CPM_yearly$Year==paste0(currentyear, " YTD")),]
     
                    NIC[1:13, 5] <- c(if(nrow(YTD_monthly) == 12){YTD_monthly$CPM
                                      }else{c(YTD_monthly$CPM, rep("-", (12-nrow(YTD_monthly))))},
                                            if(nrow(YTD_yearly) == 0){"-"}else{YTD_yearly$CPM}
                    )
                    NIC[,5] <- as.character(NIC[,5])
                    names(NIC) <- c("na",
                                   "na1",
                                   "pastyear1",
                                   "currentyear_target",
                                   "currentyear_actual") 
                    return(NIC)
   })
   
   ALL <- reactive({if(is.null(CPM_yearly_all())|is.null(CPM_monthly_all()))return(NULL)
                    CPM_monthly <- CPM_monthly_all()
                    CPM_yearly <- CPM_yearly_all()
                    validate(need(input$YTD[2]-input$YTD[1] < 365,
                            "Warning Message: Selected Date Range should be less than 12 months."))
                    reduction <- as.numeric(reduction_ALL())
                    ALL <- data.frame()
     
                    #pastyear2 <- format((input$YTD[1] - years(2)), "%Y")
                    pastyear1 <- format((input$YTD[1] - years(1)), "%Y")
                    currentyear <- format(input$YTD[1], "%Y")
                    target <- round(CPM_yearly[which(CPM_yearly$Year==pastyear1),]$CPM * reduction, 1)
                    ALL[1:13, 1] <- c(as.character(as.yearmon(seq(input$YTD[1], by = "month", length.out = 12))), "YTD")
                    ALL[1:13, 2] <- rep("", 13)
                    #ALL[1:13, 3] <- c(CPM_yearly[which(CPM_yearly$Year==pastyear2),]$CPM, rep("", 12))
                    ALL[1:13, 3] <- c(CPM_yearly[which(CPM_yearly$Year==pastyear1),]$CPM, rep("", 12))
                    ALL[1:13, 4] <- c(target, rep("", 12))
     
                    YTD_monthly <- CPM_monthly[which(CPM_monthly$Month >= as.yearmon(input$YTD[1]) & CPM_monthly$Month <= as.yearmon(input$YTD[2])),]
                    YTD_yearly <- CPM_yearly[which(CPM_yearly$Year==paste0(currentyear, " YTD")),]
     
                    ALL[1:13, 5] <- c(if(nrow(YTD_monthly) == 12){YTD_monthly$CPM
                                          }else{c(YTD_monthly$CPM, rep("-", (12-nrow(YTD_monthly))))},
                                                if(nrow(YTD_yearly) == 0){"-"}else{YTD_yearly$CPM}
                                    )
                    ALL[,5] <- as.character(ALL[,5])
                    names(ALL) <- c("na",
                                    "na1",
                                    "pastyear1",
                                    "currentyear_target",
                                    "currentyear_actual") 
                    return(ALL)
   })
   
   GEM <- reactive({if(is.null(CPM_region_yearly())|is.null(CPM_region_monthly()))return(NULL)
                    CPM_monthly <- CPM_region_monthly()
                    CPM_yearly <- CPM_region_yearly()
                    validate(need(input$YTD[2]-input$YTD[1] < 365,
                             "Warning Message: Selected Date Range should be less than 12 months."))
                    reduction <- as.numeric(reduction_region_GEM())
                    GEM <- data.frame()
     
                    #pastyear2 <- format((input$YTD[1] - years(2)), "%Y")
                     pastyear1 <- format((input$YTD[1] - years(1)), "%Y")
                     currentyear <- format(input$YTD[1], "%Y")
                     target <- round(CPM_yearly[which(CPM_yearly$Year==pastyear1),]$GEM.CPM * reduction, 1)
                     GEM[1:13, 1] <- c(as.character(as.yearmon(seq(input$YTD[1], by = "month", length.out = 12))), "YTD")
                     GEM[1:13, 2] <- rep("", 13)
                     GEM[1:13, 3] <- c(CPM_yearly[which(CPM_yearly$Year==format((input$YTD[1]) - years(1), "%Y")),]$GEM.CPM, rep("", 12))
                     GEM[1:13, 4] <- c(target, rep("", 12))
     
                     YTD_monthly <- CPM_monthly[which(CPM_monthly$Month >= as.yearmon(input$YTD[1]) & CPM_monthly$Month <= as.yearmon(input$YTD[2])),]
                     YTD_yearly <- CPM_yearly[which(CPM_yearly$Year==paste0(format((input$YTD[1]), "%Y")," YTD")),]
     
                     GEM[1:13, 5] <- c(if(nrow(YTD_monthly) == 12){YTD_monthly$GEM.CPM
                                                             }else{c(YTD_monthly$GEM.CPM, rep("-", (12-nrow(YTD_monthly))))},
                                                                   if(nrow(YTD_yearly) == 0){"-"}else{YTD_yearly$GEM.CPM}
                                        )
                     GEM[,5] <- as.character(GEM[,5])
                     names(GEM) <- c("na",
                                     "na1",
                                     "pastyear1",
                                     "currentyear_target",
                                     "currentyear_actual") 
                     return(GEM)
   })
   
   EUR <- reactive({if(is.null(CPM_region_yearly())|is.null(CPM_region_monthly()))return(NULL)
                    CPM_monthly <- CPM_region_monthly()
                    CPM_yearly <- CPM_region_yearly()
                    validate(need(input$YTD[2]-input$YTD[1] < 365,
                             "Warning Message: Selected Date Range should be less than 12 months."))
                    reduction <- as.numeric(input$reduction_EUR)
                    EUR <- data.frame()
                    pastyear1 <- format((input$YTD[1] - years(1)), "%Y")
                    currentyear <- format(input$YTD[1], "%Y")
                    target <- round(CPM_yearly[which(CPM_yearly$Year==pastyear1),]$EUR.CPM * reduction, 1)
                    EUR[1:13, 1] <- c(as.character(as.yearmon(seq(input$YTD[1], by = "month", length.out = 12))), "YTD")
                    EUR[1:13, 2] <- rep("", 13)
                    EUR[1:13, 3] <- c(CPM_yearly[which(CPM_yearly$Year==format((input$YTD[1]) - years(1), "%Y")),]$EUR.CPM, rep("", 12))
                    EUR[1:13, 4] <- c(target, rep("", 12))
     
                    YTD_monthly <- CPM_monthly[which(CPM_monthly$Month >= as.yearmon(input$YTD[1]) & CPM_monthly$Month <= as.yearmon(input$YTD[2])),]
                    YTD_yearly <- CPM_yearly[which(CPM_yearly$Year==paste0(format((input$YTD[1]), "%Y")," YTD")),]
     
                    EUR[1:13, 5] <- c(if(nrow(YTD_monthly) == 12){YTD_monthly$EUR.CPM
                                                            }else{c(YTD_monthly$EUR.CPM, rep("-", (12-nrow(YTD_monthly))))},
                                                                 if(nrow(YTD_yearly) == 0){"-"}else{YTD_yearly$EUR.CPM}
                                    )
                    EUR[,5] <- as.character(EUR[,5])
                    names(EUR) <- c("na",
                                    "na1",
                                    "pastyear1",
                                    "currentyear_target",
                                    "currentyear_actual") 
                    return(EUR)
   })
   
   NAM <- reactive({if(is.null(CPM_region_yearly())|is.null(CPM_region_monthly()))return(NULL)
                    CPM_monthly <- CPM_region_monthly()
                    CPM_yearly <- CPM_region_yearly()
                    validate(need(input$YTD[2]-input$YTD[1] < 365,
                             "Warning Message: Selected Date Range should be less than 12 months."))
                    reduction <- as.numeric(input$reduction_NAM)
                    NAM <- data.frame()
                    pastyear1 <- format((input$YTD[1] - years(1)), "%Y")
                    currentyear <- format(input$YTD[1], "%Y")
                    target <- round(CPM_yearly[which(CPM_yearly$Year==pastyear1),]$NAM.CPM * reduction, 1)
                    NAM[1:13, 1] <- c(as.character(as.yearmon(seq(input$YTD[1], by = "month", length.out = 12))), "YTD")
                    NAM[1:13, 2] <- rep("", 13)
                    NAM[1:13, 3] <- c(CPM_yearly[which(CPM_yearly$Year==format((input$YTD[1]) - years(1), "%Y")),]$NAM.CPM, rep("", 12))
                    NAM[1:13, 4] <- c(target, rep("", 12))
     
                    YTD_monthly <- CPM_monthly[which(CPM_monthly$Month >= as.yearmon(input$YTD[1]) & CPM_monthly$Month <= as.yearmon(input$YTD[2])),]
                    YTD_yearly <- CPM_yearly[which(CPM_yearly$Year==paste0(format((input$YTD[1]), "%Y")," YTD")),]
     
                    NAM[1:13, 5] <- c(if(nrow(YTD_monthly) == 12){YTD_monthly$NAM.CPM
                                             }else{c(YTD_monthly$NAM.CPM, rep("-", (12-nrow(YTD_monthly))))},
                                                  if(nrow(YTD_yearly) == 0){"-"}else{YTD_yearly$NAM.CPM}
                     )
                    NAM[,5] <- as.character(NAM[,5])
                    names(NAM) <- c("na",
                                    "na1",
                                    "pastyear1",
                                    "currentyear_target",
                                    "currentyear_actual") 
                    return(NAM)
       })
   
   
   AP <- reactive({if(is.null(CPM_region_yearly2())|is.null(CPM_region_monthly2()))return(NULL)
                   CPM_monthly <- CPM_region_monthly2()
                   CPM_yearly <- CPM_region_yearly2()
                   validate(need(input$YTD[2]-input$YTD[1] < 365,
                            "Warning Message: Selected Date Range should be less than 12 months."))
                   reduction <- as.numeric(input$reduction_AP)
                   AP <- data.frame()
                   pastyear1 <- format((input$YTD[1] - years(1)), "%Y")
                   currentyear <- format(input$YTD[1], "%Y")
                   target <- round(CPM_yearly[which(CPM_yearly$Year==pastyear1),]$AP.CPM * reduction, 1)
                   AP[1:13, 1] <- c(as.character(as.yearmon(seq(input$YTD[1], by = "month", length.out = 12))), "YTD")
                   AP[1:13, 2] <- rep("", 13)
                   AP[1:13, 3] <- c(CPM_yearly[which(CPM_yearly$Year==format((input$YTD[1]) - years(1), "%Y")),]$AP.CPM, rep("", 12))
                   AP[1:13, 4] <- c(target, rep("", 12))
     
                   YTD_monthly <- CPM_monthly[which(CPM_monthly$Month >= as.yearmon(input$YTD[1]) & CPM_monthly$Month <= as.yearmon(input$YTD[2])),]
                   YTD_yearly <- CPM_yearly[which(CPM_yearly$Year==paste0(format((input$YTD[1]), "%Y")," YTD")),]
     
                   AP[1:13, 5] <- c(if(nrow(YTD_monthly) == 12){YTD_monthly$AP.CPM
                                   }else{c(YTD_monthly$AP.CPM, rep("-", (12-nrow(YTD_monthly))))},
                                     if(nrow(YTD_yearly) == 0){"-"}else{YTD_yearly$AP.CPM}
                                  )
                   AP[,5] <- as.character(AP[,5])
                   names(AP) <- c("na",
                                  "na1",
                                  "pastyear1",
                                  "currentyear_target",
                                  "currentyear_actual") 
                   return(AP)
   })
   
   LA <- reactive({if(is.null(CPM_region_yearly2())|is.null(CPM_region_monthly2()))return(NULL)
                   CPM_monthly <- CPM_region_monthly2()
                   CPM_yearly <- CPM_region_yearly2()
                   validate(need(input$YTD[2]-input$YTD[1] < 365,
                            "Warning Message: Selected Date Range should be less than 12 months."))
                   reduction <- as.numeric(input$reduction_LA)
                   LA <- data.frame()
                   pastyear1 <- format((input$YTD[1] - years(1)), "%Y")
                   currentyear <- format(input$YTD[1], "%Y")
                   target <- round(CPM_yearly[which(CPM_yearly$Year==pastyear1),]$LA.CPM * reduction, 1)
                   LA[1:13, 1] <- c(as.character(as.yearmon(seq(input$YTD[1], by = "month", length.out = 12))), "YTD")
                   LA[1:13, 2] <- rep("", 13)
                   LA[1:13, 3] <- c(CPM_yearly[which(CPM_yearly$Year==format((input$YTD[1]) - years(1), "%Y")),]$LA.CPM, rep("", 12))
                   LA[1:13, 4] <- c(target, rep("", 12))
     
                   YTD_monthly <- CPM_monthly[which(CPM_monthly$Month >= as.yearmon(input$YTD[1]) & CPM_monthly$Month <= as.yearmon(input$YTD[2])),]
                   YTD_yearly <- CPM_yearly[which(CPM_yearly$Year==paste0(format((input$YTD[1]), "%Y")," YTD")),]
     
                   LA[1:13, 5] <- c(if(nrow(YTD_monthly) == 12){YTD_monthly$LA.CPM
                                  }else{c(YTD_monthly$LA.CPM, rep("-", (12-nrow(YTD_monthly))))},
                                   if(nrow(YTD_yearly) == 0){"-"}else{YTD_yearly$LA.CPM}
                     )
                   LA[,5] <- as.character(LA[,5])
                   names(LA) <- c("na",
                                  "na1",
                                  "pastyear1",
                                  "currentyear_target",
                                  "currentyear_actual") 
                    return(LA)
   })
   
   MEA <- reactive({if(is.null(CPM_region_yearly2())|is.null(CPM_region_monthly2()))return(NULL)
                    CPM_monthly <- CPM_region_monthly2()
                    CPM_yearly <- CPM_region_yearly2()
                    validate(need(input$YTD[2]-input$YTD[1] < 365,
                             "Warning Message: Selected Date Range should be less than 12 months."))
                    reduction <- as.numeric(input$reduction_MEA)
                    MEA <- data.frame()
                    pastyear1 <- format((input$YTD[1] - years(1)), "%Y")
                    currentyear <- format(input$YTD[1], "%Y")
                    target <- round(CPM_yearly[which(CPM_yearly$Year==pastyear1),]$MEA.CPM * reduction, 1)
                    MEA[1:13, 1] <- c(as.character(as.yearmon(seq(input$YTD[1], by = "month", length.out = 12))), "YTD")
                    MEA[1:13, 2] <- rep("", 13)
                    MEA[1:13, 3] <- c(CPM_yearly[which(CPM_yearly$Year==format((input$YTD[1]) - years(1), "%Y")),]$MEA.CPM, rep("", 12))
                    MEA[1:13, 4] <- c(target, rep("", 12))
     
                    YTD_monthly <- CPM_monthly[which(CPM_monthly$Month >= as.yearmon(input$YTD[1]) & CPM_monthly$Month <= as.yearmon(input$YTD[2])),]
                    YTD_yearly <- CPM_yearly[which(CPM_yearly$Year==paste0(format((input$YTD[1]), "%Y")," YTD")),]
     
                    MEA[1:13, 5] <- c(if(nrow(YTD_monthly) == 12){YTD_monthly$MEA.CPM
                                     }else{c(YTD_monthly$MEA.CPM, rep("-", (12-nrow(YTD_monthly))))},
                                      if(nrow(YTD_yearly) == 0){"-"}else{YTD_yearly$MEA.CPM}
                                     )
                    MEA[,5] <- as.character(MEA[,5])
                    names(MEA) <- c("na",
                                    "na1",
                                    "pastyear1",
                                    "currentyear_target",
                                    "currentyear_actual") 
                     return(MEA)
   })

  target <- reactive({if(is.null(CPM_yearly())|is.null(CPM_yearly_NIC()))return(NULL)
                      CPM_yearly <- CPM_yearly()
                      CPM_yearly_NIC <- CPM_yearly_NIC()
                      CPM_yearly_all <- CPM_yearly_all()
                    
                      reduction <- as.numeric(reduction_ALL())
                      reduction_NIC <- as.numeric(reduction_NIC())
                      reduction_AWC <- as.numeric(input$reduction_AWC)
                      reduction_CC <- as.numeric(input$reduction_CC)
                      reduction_C <- as.numeric(input$reduction_C)
                      reduction_OC <- as.numeric(input$reduction_OC)
                      reduction_IC <- as.numeric(input$reduction_IC)
                      reduction_Eurotec <- as.numeric(input$reduction_Eurotec)
                      pastyear1 <- format((input$YTD[1] - years(1)), "%Y")
                      currentyear <- format(input$YTD[1], "%Y")
                      
                      target_AWC <- round(CPM_yearly[which(CPM_yearly$Year==pastyear1),]$AWC.CPM * reduction_AWC, 1)
                      target_CC <- round(CPM_yearly[which(CPM_yearly$Year==pastyear1),]$CC.CPM * reduction_CC, 1)
                      target_C <- round(CPM_yearly[which(CPM_yearly$Year==pastyear1),]$C.CPM * reduction_C, 1)
                      target_OC <- round(CPM_yearly[which(CPM_yearly$Year==pastyear1),]$OC.CPM * reduction_OC, 1)
                      target_IC <- round(CPM_yearly[which(CPM_yearly$Year==pastyear1),]$IC.CPM * reduction_IC, 1)
                      target_Eurotec <- round(CPM_yearly[which(CPM_yearly$Year==pastyear1),]$Eurotec.CPM * reduction_Eurotec, 1)
                      target_NIC <- round(CPM_yearly_NIC[which(CPM_yearly_NIC$Year==pastyear1),]$CPM * reduction_NIC, 1)
                      target_ALL <- round(CPM_yearly_all[which(CPM_yearly_all$Year==pastyear1),]$CPM * reduction, 1)
                      target <- data.frame("BU" = c("All BU", "NIC", "IC", "AWC", "CC", "C", "OC", "(Eurotec)"),
                                           "Target_CPM" = c(target_ALL, target_NIC, target_IC, target_AWC, target_CC, target_C, target_OC, target_Eurotec))
                      target$"Target_CPM" <- as.character(target$"Target_CPM")
                      names(target) <- c("BU", "Target")
                      return(target)
  })
  
  
  
  target_html <- reactive({if(is.null(target()))return(NULL)
                           target <- target()
                           CPM_yearly <- CPM_yearly()
                           CPM_yearly_NIC <- CPM_yearly_NIC()
                           CPM_yearly_all <- CPM_yearly_all()
                           
                           reduction <- as.numeric(reduction_ALL())
                           reduction_NIC <- as.numeric(reduction_NIC())
                           reduction_AWC <- as.numeric(input$reduction_AWC)
                           reduction_CC <- as.numeric(input$reduction_CC)
                           reduction_C <- as.numeric(input$reduction_C)
                           reduction_OC <- as.numeric(input$reduction_OC)
                           reduction_IC <- as.numeric(input$reduction_IC)
                           pastyear1 <- format((input$YTD[1] - years(1)), "%Y")
                           currentyear <- format(input$YTD[1], "%Y")
                           
                           tb <- paste0('<table class = "tf" id = "t01">
                                         <tr>
                                            <th> Business Unit </th>
                                            <td> All BU </td>
                                            <td> Non-IC </td>
                                            <td> IC </td>
                                            <td> AWC </td>
                                            <td> CC </td>
                                            <td> C </td>
                                            <td> OC </td>
                                         </tr>
                                         <tr>
                                            <th> Previous YE CPM </th> 
                                            <td>', round(CPM_yearly_all[which(CPM_yearly_all$Year==pastyear1),]$CPM, 1),'</td>
                                            <td>', round(CPM_yearly_NIC[which(CPM_yearly_NIC$Year==pastyear1),]$CPM, 1),' </td>
                                            <td>', round(CPM_yearly[which(CPM_yearly$Year==pastyear1),]$IC.CPM, 1),' </td>
                                            <td>', round(CPM_yearly[which(CPM_yearly$Year==pastyear1),]$AWC.CPM, 1),' </td>
                                            <td>', round(CPM_yearly[which(CPM_yearly$Year==pastyear1),]$CC.CPM, 1),' </td>
                                            <td>', round(CPM_yearly[which(CPM_yearly$Year==pastyear1),]$C.CPM, 1),' </td>
                                            <td>', round(CPM_yearly[which(CPM_yearly$Year==pastyear1),]$OC.CPM, 1),' </td>
                                        </tr>
                                        <tr>
                                            <th> % of Reduction </th> 
                                            <td>', paste0((100 - as.numeric(reduction_ALL())*100), "%"),'</td>
                                            <td>', paste0((100 - as.numeric(reduction_NIC())*100), "%"),' </td>
                                            <td>', paste0((100 - as.numeric(input$reduction_IC)*100), "%"),' </td>
                                            <td>', paste0((100 - as.numeric(input$reduction_AWC)*100), "%"),' </td>
                                            <td>', paste0((100 - as.numeric(input$reduction_CC)*100), "%"),' </td>
                                            <td>', paste0((100 - as.numeric(input$reduction_C)*100), "%"),' </td>
                                            <td>', paste0((100 - as.numeric(input$reduction_OC)*100), "%"),' </td>
                                        </tr>
                                        <tr>
                                            <th> New Year Target CPM </th> 
                                            <td>', target[which(target$BU == "All BU"),]$Target ,'</td>
                                            <td>', target[which(target$BU == "NIC"),]$Target ,' </td>
                                            <td>', target[which(target$BU == "IC"),]$Target ,' </td>
                                            <td>', target[which(target$BU == "AWC"),]$Target ,' </td>
                                            <td>', target[which(target$BU == "CC"),]$Target ,' </td>
                                            <td>', target[which(target$BU == "C"),]$Target ,' </td>
                                            <td>', target[which(target$BU == "OC"),]$Target ,' </td>
                                         </tr>
                                         </table>')
                           return(tb)
  })
  
  BSC <- reactive({if(is.null(target()))return(NULL)
                   target <- target()
                   target_region <- target_region()
                   pastyear1 <- format((input$YTD[1] - years(1)), "%Y")
                   currentyear <- format(input$YTD[1], "%Y")
                   CPM_yearly <- CPM_yearly()
                   CPM_region_yearly <- CPM_region_yearly()
                   CPM_yearly_NIC <- CPM_yearly_NIC()
                   CPM_yearly_all <- CPM_yearly_all()
                   YTD_yearly_region <- CPM_region_yearly[which(CPM_region_yearly$Year==paste0(currentyear, " YTD")),]
                   YTD_yearly <- CPM_yearly[which(CPM_yearly$Year==paste0(currentyear, " YTD")),]
                   YTD_yearly_NIC <- CPM_yearly_NIC[which(CPM_yearly_NIC$Year==paste0(currentyear, " YTD")),]
                   YTD_yearly_all <- CPM_yearly_all[which(CPM_yearly_all$Year==paste0(currentyear, " YTD")),]
                   

                   pastyear1 <- format((input$YTD[1] - years(1)), "%Y")
                   currentyear <- format(input$YTD[1], "%Y")
                   
                   BSC <- data.frame()
                   BSC[1:13, 1] <- c("Business Unit", "", "All-Combined", "Ostomy Care", "Wound Care",
                                     "Continence", "Critical", "Infusion Care","Region", "",
                                     "Global Emerging Market (GEM)", "North America (NA)", "Europe(EUR)")
                   BSC[1:13, 2] <- rep("", 13)
                   BSC[1:13, 3] <- c(paste0(pastyear1, " Baseline"), "",
                                     CPM_yearly_all[which(CPM_yearly_all$Year==pastyear1),]$CPM, 
                                     CPM_yearly[which(CPM_yearly$Year==pastyear1),]$OC.CPM, 
                                     CPM_yearly[which(CPM_yearly$Year==pastyear1),]$AWC.CPM, 
                                     CPM_yearly[which(CPM_yearly$Year==pastyear1),]$CC.CPM, 
                                     CPM_yearly[which(CPM_yearly$Year==pastyear1),]$C.CPM, 
                                     CPM_yearly[which(CPM_yearly$Year==pastyear1),]$IC.CPM,
                                     paste0(pastyear1, " Baseline"), "",
                                     CPM_region_yearly[which(CPM_region_yearly$Year==pastyear1),]$GEM.CPM,
                                     CPM_region_yearly[which(CPM_region_yearly$Year==pastyear1),]$NAM.CPM,
                                     CPM_region_yearly[which(CPM_region_yearly$Year==pastyear1),]$EUR.CPM
                                     )
                   BSC[1:13, 4] <- c("12% Target", "",
                                     target[which(target$BU == "All BU"),]$Target,
                                     target[which(target$BU == "OC"),]$Target,
                                     target[which(target$BU == "AWC"),]$Target,
                                     target[which(target$BU == "CC"),]$Target,
                                     target[which(target$BU == "C"),]$Target,
                                     target[which(target$BU == "IC"),]$Target,
                                     "12% Target", "",
                                     target_region[which(target_region$Region == "GEM"),]$Target,
                                     target_region[which(target_region$Region == "NAM"),]$Target,
                                     target_region[which(target_region$Region == "EUR"),]$Target
                                     )
                   BSC[1:13, 5] <- c("2021 Performance", "YTD",
                                     YTD_yearly_all$CPM,
                                     YTD_yearly$OC.CPM,
                                     YTD_yearly$AWC.CPM,
                                     YTD_yearly$CC.CPM,
                                     YTD_yearly$C.CPM,
                                     YTD_yearly$IC.CPM,
                                     "2021 Performance", "YTD",
                                     YTD_yearly_region$GEM.CPM,
                                     YTD_yearly_region$NAM.CPM,
                                     YTD_yearly_region$EUR.CPM
                                     )
                   BSC[1:13, 6] <- c("","Actually % Reduction",
                                     paste0((round(1 - as.numeric(BSC[3,5])/as.numeric(BSC[3,3]), 3))*100, "%"),
                                     paste0((round(1 - as.numeric(BSC[4,5])/as.numeric(BSC[4,3]), 3))*100, "%"),
                                     paste0((round(1 - as.numeric(BSC[5,5])/as.numeric(BSC[5,3]), 3))*100, "%"),
                                     paste0((round(1 - as.numeric(BSC[6,5])/as.numeric(BSC[6,3]), 3))*100, "%"),
                                     paste0((round(1 - as.numeric(BSC[7,5])/as.numeric(BSC[7,3]), 3))*100, "%"),
                                     paste0((round(1 - as.numeric(BSC[8,5])/as.numeric(BSC[8,3]), 3))*100, "%"),
                                     "","Actually % Reduction",
                                     paste0((round(1 - as.numeric(BSC[11,5])/as.numeric(BSC[11,3]), 3))*100, "%"),
                                     paste0((round(1 - as.numeric(BSC[12,5])/as.numeric(BSC[12,3]), 3))*100, "%"),
                                     paste0((round(1 - as.numeric(BSC[13,5])/as.numeric(BSC[13,3]), 3))*100, "%")
                                     )
                   
                
                   return(BSC)
    
  })
  
  BSC_region <- reactive({if(is.null(target()))return(NULL)
                          target_region <- target_region()
                          pastyear1 <- format((input$YTD[1] - years(1)), "%Y")
                          currentyear <- format(input$YTD[1], "%Y")
                          CPM_region_yearly <- CPM_region_yearly()
                          CPM_region_yearly2 <- CPM_region_yearly2()
                          YTD_yearly_region <- CPM_region_yearly[which(CPM_region_yearly$Year==paste0(currentyear, " YTD")),]
                          YTD_yearly_region2 <- CPM_region_yearly2[which(CPM_region_yearly2$Year==paste0(currentyear, " YTD")),]

                          BSC <- data.frame()
                          BSC[1:8, 1] <- c("Region", "",
                                            "Global Emerging Market (GEM)", "", "", "", "North America (NA)", "Europe(EUR)")
                          BSC[1:8, 2] <- rep("", 8)
                          BSC[1:8, 3] <- c("", "", "APAC", "LATAM", "MEA", "Total GEM", "", "")
                          BSC[1:8, 4] <- c(paste0(pastyear1, " Baseline"), "",
                                           CPM_region_yearly2[which(CPM_region_yearly2$Year==pastyear1),]$AP.CPM,
                                           CPM_region_yearly2[which(CPM_region_yearly2$Year==pastyear1),]$LA.CPM,
                                           CPM_region_yearly2[which(CPM_region_yearly2$Year==pastyear1),]$MEA.CPM,
                                           CPM_region_yearly[which(CPM_region_yearly$Year==pastyear1),]$GEM.CPM,
                                           CPM_region_yearly[which(CPM_region_yearly$Year==pastyear1),]$NAM.CPM,
                                           CPM_region_yearly[which(CPM_region_yearly$Year==pastyear1),]$EUR.CPM
                                          )
                          BSC[1:8, 5] <- c("10% Target", "",
                                           target_region[which(target_region$Region == "APAC"),]$Target,
                                           target_region[which(target_region$Region == "LATAM"),]$Target,
                                           target_region[which(target_region$Region == "MEA"),]$Target,
                                           target_region[which(target_region$Region == "GEM"),]$Target,
                                           target_region[which(target_region$Region == "NAM"),]$Target,
                                           target_region[which(target_region$Region == "EUR"),]$Target
                                          )
                          BSC[1:8, 6] <- c("2020 Performance", "YTD",
                                           YTD_yearly_region2$AP.CPM,
                                           YTD_yearly_region2$LA.CPM,
                                           YTD_yearly_region2$MEA.CPM,
                                           YTD_yearly_region$GEM.CPM,
                                           YTD_yearly_region$NAM.CPM,
                                           YTD_yearly_region$EUR.CPM
                                           )
                          BSC[1:8, 7] <- c( "","Actually % Reduction",
                                           paste0((round(1 - as.numeric(BSC[3,6])/as.numeric(BSC[3,4]), 3))*100, "%"),
                                           paste0((round(1 - as.numeric(BSC[4,6])/as.numeric(BSC[4,4]), 3))*100, "%"),
                                           paste0((round(1 - as.numeric(BSC[5,6])/as.numeric(BSC[5,4]), 3))*100, "%"),
                                           paste0((round(1 - as.numeric(BSC[6,6])/as.numeric(BSC[6,4]), 3))*100, "%"),
                                           paste0((round(1 - as.numeric(BSC[7,6])/as.numeric(BSC[7,4]), 3))*100, "%"),
                                           paste0((round(1 - as.numeric(BSC[8,6])/as.numeric(BSC[8,4]), 3))*100, "%")
                                           )
                           return(BSC)
    
  })
  
  GMSC <- reactive({if(is.null(target()))return(NULL)
                    target <- target()
                    CPM_yearly <- CPM_yearly()
                    CPM_yearly_NIC <- CPM_yearly_NIC()
                    CPM_yearly_all <- CPM_yearly_all()
                    CPM_monthly <- CPM_monthly()
                    CPM_monthly_NIC <- CPM_monthly_NIC()
                    CPM_monthly_all <- CPM_monthly_all()
                    
                    pastyear2 <- format((input$YTD[1] - years(2)), "%Y")
                    pastyear1 <- format((input$YTD[1] - years(1)), "%Y")
                    currentyear <- format(input$YTD[1], "%Y")
                    
                    YTD_monthly <- CPM_monthly[which(CPM_monthly$Month >= as.yearmon(input$YTD[1]) & CPM_monthly$Month <= as.yearmon(input$YTD[2])),]
                    YTD_yearly <- CPM_yearly[which(CPM_yearly$Year==paste0(currentyear, " YTD")),]
                    YTD_monthly_NIC <- CPM_monthly_NIC[which(CPM_monthly_NIC$Month >= as.yearmon(input$YTD[1]) & CPM_monthly_NIC$Month <= as.yearmon(input$YTD[2])),]
                    YTD_yearly_NIC <- CPM_yearly_NIC[which(CPM_yearly_NIC$Year==paste0(currentyear, " YTD")),]
                    YTD_monthly_all <- CPM_monthly_all[which(CPM_monthly_all$Month >= as.yearmon(input$YTD[1]) & CPM_monthly_all$Month <= as.yearmon(input$YTD[2])),]
                    YTD_yearly_all <- CPM_yearly_all[which(CPM_yearly_all$Year==paste0(currentyear, " YTD")),]
                    
                    
                    
                    GMSC <- data.frame()
                    GMSC[1:17,1] <- c( 
                                      paste0(pastyear2, " Baseline"), 
                                      paste0(pastyear1, " Baseline"), 
                                      paste0(currentyear, " target (10%)"), 
                                      as.character(as.yearmon(seq(input$YTD[1], by = "month", length.out = 12))),
                                      paste0(currentyear, " YTD CPM"),
                                      "Actual % reduction"
                                      ) 
                    GMSC[1:17,2] <- rep("", 17)
                    GMSC[1:17,3] <- rep("", 17)
                    GMSC[1:16,4] <- c(
                                      CPM_yearly[which(CPM_yearly$Year==pastyear2),]$OC.CPM,
                                      CPM_yearly[which(CPM_yearly$Year==pastyear1),]$OC.CPM,
                                      target[which(target$BU == "OC"),]$Target,
                                      c(if(nrow(YTD_monthly) == 12){YTD_monthly$OC.CPM
                                        }else{c(YTD_monthly$OC.CPM, rep("-", (12-nrow(YTD_monthly))))},
                                        
                                      if(nrow(YTD_yearly) == 0){"-"}else{YTD_yearly$OC.CPM}
                                       )
                    )
                    GMSC[17,4] <- paste0((round(1 - as.numeric(GMSC[16,4])/as.numeric(GMSC[2,4]), 3))*100, "%")
                    GMSC[1:17,5] <- rep("", 17)
                    GMSC[1:16,6] <- c(
                                      CPM_yearly[which(CPM_yearly$Year==pastyear2),]$AWC.CPM,
                                      CPM_yearly[which(CPM_yearly$Year==pastyear1),]$AWC.CPM,
                                      target[which(target$BU == "AWC"),]$Target,
                                      c(if(nrow(YTD_monthly) == 12){YTD_monthly$AWC.CPM
                                      }else{c(YTD_monthly$AWC.CPM, rep("-", (12-nrow(YTD_monthly))))},
                                      
                                      if(nrow(YTD_yearly) == 0){"-"}else{YTD_yearly$AWC.CPM}
                                      )
                    )
                    GMSC[17,6] <- paste0((round(1 - as.numeric(GMSC[16,6])/as.numeric(GMSC[2,6]), 3))*100, "%")
                    GMSC[1:17,7] <- rep("", 17)
                    GMSC[1:16,8] <- c(CPM_yearly[which(CPM_yearly$Year==pastyear2),]$CC.CPM,
                                      CPM_yearly[which(CPM_yearly$Year==pastyear1),]$CC.CPM,
                                      target[which(target$BU == "CC"),]$Target,
                                      c(if(nrow(YTD_monthly) == 12){YTD_monthly$CC.CPM
                                      }else{c(YTD_monthly$CC.CPM, rep("-", (12-nrow(YTD_monthly))))},
                                      
                                      if(nrow(YTD_yearly) == 0){"-"}else{YTD_yearly$CC.CPM}
                                      )
                    )
                    GMSC[17,8] <- paste0((1 - round(as.numeric(GMSC[16,8])/as.numeric(GMSC[2,8]),3))*100, "%")
                    GMSC[1:17,9] <- rep("", 17)
                    GMSC[1:16,10] <- c(CPM_yearly[which(CPM_yearly$Year==pastyear2),]$C.CPM,
                                       CPM_yearly[which(CPM_yearly$Year==pastyear1),]$C.CPM,
                                       target[which(target$BU == "C"),]$Target,
                                       c(if(nrow(YTD_monthly) == 12){YTD_monthly$C.CPM
                                       }else{c(YTD_monthly$C.CPM, rep("-", (12-nrow(YTD_monthly))))},
                                      
                                       if(nrow(YTD_yearly) == 0){"-"}else{YTD_yearly$C.CPM}
                                       )
                    )
                    GMSC[17,10] <- paste0((round(1 - as.numeric(GMSC[16,10])/as.numeric(GMSC[2,10]),3))*100, "%")
                    GMSC[1:17,11] <- rep("", 17)
                    GMSC[1:16,12] <- c(CPM_yearly_NIC[which(CPM_yearly_NIC$Year==pastyear2),]$CPM,
                                       CPM_yearly_NIC[which(CPM_yearly_NIC$Year==pastyear1),]$CPM,
                                       target[which(target$BU == "NIC"),]$Target,
                                       c(if(nrow(YTD_monthly_NIC) == 12){YTD_monthly_NIC$CPM
                                       }else{c(YTD_monthly_NIC$CPM, rep("-", (12-nrow(YTD_monthly_NIC))))},
                                      
                                       if(nrow(YTD_yearly_NIC) == 0){"-"}else{YTD_yearly_NIC$CPM}
                                       )
                    )
                    GMSC[17,12] <- paste0((1 - round(as.numeric(GMSC[16,12])/as.numeric(GMSC[2,12]),3))*100, "%")
                    GMSC[1:17,13] <- rep("", 17)
                    GMSC[1:16,14] <- c(CPM_yearly[which(CPM_yearly$Year==pastyear2),]$IC.CPM,
                                       CPM_yearly[which(CPM_yearly$Year==pastyear1),]$IC.CPM,
                                       target[which(target$BU == "IC"),]$Target,
                                       c(if(nrow(YTD_monthly) == 12){YTD_monthly$IC.CPM
                                           }else{c(YTD_monthly$IC.CPM, rep("-", (12-nrow(YTD_monthly))))},
                      
                                       if(nrow(YTD_yearly) == 0){"-"}else{YTD_yearly$IC.CPM}
                      )
                    )
                    GMSC[17,14] <- paste0((1 - round(as.numeric(GMSC[16,14])/as.numeric(GMSC[2,14]),3))*100, "%")
                    GMSC[1:17,15] <- rep("", 17)
                    GMSC[1:16,16] <- c(CPM_yearly_all[which(CPM_yearly_all$Year==pastyear2),]$CPM,
                                       CPM_yearly_all[which(CPM_yearly_all$Year==pastyear1),]$CPM,
                                       target[which(target$BU == "All BU"),]$Target,
                                       c(if(nrow(YTD_monthly_all) == 12){YTD_monthly_all$CPM
                                          }else{c(YTD_monthly_all$CPM, rep("-", (12-nrow(YTD_monthly_all))))},
                      
                                       if(nrow(YTD_yearly_all)== 0){"-"}else{YTD_yearly_all$CPM}
                      )
                    )
                    GMSC[17,16] <- paste0((1 - round(as.numeric(GMSC[16,16])/as.numeric(GMSC[2,16]),3))*100, "%")
                    GMSC[1:17,17] <- rep("", 17)
                    
                    names(GMSC) <- c("BU", "na1","na2",
                                    "OC", "na3",
                                    "AWC", "na4",
                                    "CC", "na5",
                                    "C", "na6",
                                    "NIC", "na7",
                                    "IC", "na8",
                                    "ALL", "na9")
                    return(GMSC)
                    
  })
  
  CPM_COM_NIC <- reactive({if(is.null(CPM_monthly_NIC()))return(NULL)
                 data <- CPM_monthly_NIC()
                 n <- nrow(data)
                 pastyear2 <- format((input$YTD[1] - years(2)), "%Y")
                 pastyear1 <- format((input$YTD[1] - years(1)), "%Y")
                 currentyear <- format(input$YTD[1], "%Y")
                 #data$Month <- seq(input$YTD[1]- years(2), by = "month", length.out = n)
                 #data <- data[data$Month >= (input$YTD[1]- years(2)),]
                 
                 #data2 <- data[data$Month >= (input$YTD[1]- years(2)) & data$Month < (input$YTD[1]- years(1)),]
                 #data1 <- data[data$Month >= (input$YTD[1]- years(1)) & data$Month < input$YTD[1],]
                 #data0 <- data[data$Month >= input$YTD[1],]
                 data <- data[which(data$Month >= as.yearmon(input$YTD[1]- years(2))),]
                 data2 <- data[which(data$Month >= as.yearmon(input$YTD[1]- years(2)) & data$Month < as.yearmon(input$YTD[1]- years(1))),]
                 data1 <- data[which(data$Month >= as.yearmon(input$YTD[1]- years(1)) & data$Month < as.yearmon(input$YTD[1])),]
                 data0 <- data[which(data$Month >= as.yearmon(input$YTD[1])),]
                 
                 
                 tb <- data.frame()
                 tb[1:6,1] <- c(pastyear2, pastyear2, pastyear1, pastyear1, currentyear, currentyear)
                 tb[1:6,2] <- rep(c("Complaints", "CPM"), 3)
                 tb[1, 3:14] <- as.character(round(data2$complaint, 0))
                 tb[2, 3:14] <- data2$CPM
                 tb[3, 3:14] <- data1$complaint
                 tb[4, 3:14] <- data1$CPM
                 if(nrow(data0)==12){tb[5, 3:14] <- data0$complaint
                 tb[6, 3:14] <- data0$CPM
                 }else{m <- nrow(data0)
                 data0[(m+1):12,] <- "-"
                 tb[5, 3:14] <- data0$complaint
                 tb[6, 3:14] <- data0$CPM
                 }
                 colnames(tb) <- c("Year", "-", "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")
                 
                 return(tb)
  })
  

  
  
  target_region <- reactive({if(is.null(CPM_region_yearly()))return(NULL)
                             CPM_yearly <- CPM_region_yearly()
                             CPM_yearly2 <- CPM_region_yearly2()
                             reduction <- as.numeric(reduction_region_ALL())
                             reduction_GEM <- as.numeric(reduction_region_GEM())
                             reduction_AP <- as.numeric(input$reduction_AP)
                             reduction_LA <- as.numeric(input$reduction_LA)
                             reduction_MEA <- as.numeric(input$reduction_MEA)
                             reduction_EUR <- as.numeric(input$reduction_EUR)
                             reduction_NAM <- as.numeric(input$reduction_NAM)
                             pastyear1 <- format((input$YTD[1] - years(1)), "%Y")
                             currentyear <- format(input$YTD[1], "%Y")
    
                             target_AP <- round(CPM_yearly2[which(CPM_yearly2$Year==pastyear1),]$AP.CPM * reduction_AP, 1)
                             target_LA <- round(CPM_yearly2[which(CPM_yearly2$Year==pastyear1),]$LA.CPM * reduction_LA, 1)
                             target_MEA <- round(CPM_yearly2[which(CPM_yearly2$Year==pastyear1),]$MEA.CPM * reduction_MEA, 1)
                             target_GEM <- round(CPM_yearly[which(CPM_yearly$Year==pastyear1),]$GEM.CPM * reduction_GEM, 1)
                             target_EUR <- round(CPM_yearly[which(CPM_yearly$Year==pastyear1),]$EUR.CPM * reduction_EUR, 1)
                             target_NAM <- round(CPM_yearly[which(CPM_yearly$Year==pastyear1),]$NAM.CPM * reduction_NAM, 1)
                            #target_ALL <- round(CPM_yearly_all[which(CPM_yearly_all$Year==pastyear1),]$CPM * reduction, 1)
                             target <- data.frame("Region" = c("GEM", "EUR", "NAM", "APAC", "LATAM", "MEA"),
                                                  "Target_CPM" = c(target_GEM, target_EUR, target_NAM, target_AP, target_LA, target_MEA))
                             target$"Target_CPM" <- as.character(target$"Target_CPM")
                             names(target) <- c("Region", "Target")
                             return(target)
  })
  
  target_region_html <- reactive({if(is.null(target_region()))return(NULL)
                                  target <- target_region()
                                  CPM_yearly <- CPM_region_yearly()
    
                                  reduction <- as.numeric(reduction_region_ALL())
                                  reduction_GEM <- as.numeric(reduction_region_GEM())
                                  reduction_EUR <- as.numeric(input$reduction_EUR)
                                  reduction_NAM <- as.numeric(input$reduction_NAM)
                                  pastyear1 <- format((input$YTD[1] - years(1)), "%Y")
                                  currentyear <- format(input$YTD[1], "%Y")
    
                                  tb <- paste0('<table class = "tf" id = "t01">
                                                <tr>
                                                <th> Region </th>
                                                <td> GEM </td>
                                                <td> EUR </td>
                                                <td> NAM </td>
                                                </tr>
                                                <tr>
                                                <th> Previous YE CPM </th> 
                                                <td>', round(CPM_yearly[which(CPM_yearly$Year==pastyear1),]$GEM.CPM, 1),' </td>
                                                <td>', round(CPM_yearly[which(CPM_yearly$Year==pastyear1),]$EUR.CPM, 1),' </td>
                                                <td>', round(CPM_yearly[which(CPM_yearly$Year==pastyear1),]$NAM.CPM, 1),' </td>
                                                </tr>
                                                <tr>
                                                <th> % of Reduction </th> 
                                                <td>', paste0((100 - as.numeric(reduction_region_GEM())*100), "%"),'</td>
                                                <td>', paste0((100 - as.numeric(input$reduction_EUR)*100), "%"),' </td>
                                                <td>', paste0((100 - as.numeric(input$reduction_NAM)*100), "%"),' </td>
                                                </tr>
                                                <tr>
                                                <th> New Year Target CPM </th> 
                                                <td>', target[which(target$Region == "GEM"),]$Target ,'</td>
                                                <td>', target[which(target$Region == "EUR"),]$Target ,' </td>
                                                <td>', target[which(target$Region == "NAM"),]$Target ,' </td>
                                                </tr>
                                                </table>')
                                                return(tb)
  })
  
  target_region_html2 <- reactive({if(is.null(target_region()))return(NULL)
                                   target <- target_region()
                                   CPM_yearly2 <- CPM_region_yearly2()
    
                                   reduction_AP <- as.numeric(input$reduction_AP)
                                   reduction_LA <- as.numeric(input$reduction_LA)
                                   reduction_MEA <- as.numeric(input$reduction_MEA)
                                   pastyear1 <- format((input$YTD[1] - years(1)), "%Y")
                                   currentyear <- format(input$YTD[1], "%Y")
    
                                   tb <- paste0('<table class = "tf" id = "t01">
                                                 <tr>
                                                 <th> Region under GEM </th>
                                                 <td> APAC </td>
                                                 <td> LATAM </td>
                                                 <td> MEA </td>
                                                 </tr>
                                                 <tr>
                                                 <th> Previous YE CPM </th> 
                                                 <td>', round(CPM_yearly2[which(CPM_yearly2$Year==pastyear1),]$AP.CPM, 1),' </td>
                                                 <td>', round(CPM_yearly2[which(CPM_yearly2$Year==pastyear1),]$LA.CPM, 1),' </td>
                                                 <td>', round(CPM_yearly2[which(CPM_yearly2$Year==pastyear1),]$MEA.CPM, 1),' </td>
                                                 </tr>
                                                 <tr>
                                                 <th> % of Reduction </th> 
                                                 <td>', paste0((100 - as.numeric(input$reduction_AP)*100), "%"),'</td>
                                                 <td>', paste0((100 - as.numeric(input$reduction_LA)*100), "%"),' </td>
                                                 <td>', paste0((100 - as.numeric(input$reduction_MEA)*100), "%"),' </td>
                                                 </tr>
                                                 <tr>
                                                 <th> New Year Target CPM </th> 
                                                 <td>', target[which(target$Region == "APAC"),]$Target ,'</td>
                                                 <td>', target[which(target$Region == "LATAM"),]$Target ,' </td>
                                                 <td>', target[which(target$Region == "MEA"),]$Target ,' </td>
                                                 </tr>
                                                 </table>')
                                                 return(tb)
  })
  
  
  Trend_Chart_all <- function()({if(is.null(CPM_monthly()))return(NULL)
                     CPM_monthly <- CPM_monthly()
                     month_n <- input$plot_range
                     validate(need((nrow(CPM_monthly) - (month_n - 1)) >= 0, "Warning Message: No previous data available"))
                     comp_CPM.plt <- CPM_monthly[(nrow(CPM_monthly) - (month_n - 1)):nrow(CPM_monthly),]
                     comp_CPM.plt$Month.Plot <- c(1:nrow(comp_CPM.plt))
                     #par(xpd=TRUE, mar=par()$mar + c(1, 1.2, 0, 0.8))
                     par(xpd=TRUE, mar=par()$mar + c(-2.3, 0.5, 0, 1.7)) 
                     color.scheme <- c("dodgerblue4", "darkgoldenrod1", "aquamarine4", "darkorange1", "cornsilk4") # corresponds to WT, CC, OC, IC and C respectively
                     
                     plot <- plot(comp_CPM.plt$Month.Plot, 
                                  comp_CPM.plt$AWC.CPM,
                                  type = "l",
                                  main = paste0("CPM by Business Unit for the Past ", month_n, " Months"),
                                  ylim = c(0, 100),
                                  col = color.scheme[1],
                                  lwd = 3,
                                  cex.main = 2,
                                  axes = F,
                                  ylab = expression(italic("CPM")),
                                  xlab = "",
                                  cex.lab = 2,
                                  yaxs = "i")
                     
                     #x-axis
                     labellist <- comp_CPM.plt$Month
                     axis(side = 1, at = seq(1, month_n, by = 1), labels = FALSE)
                     text(seq(1, month_n, by = 1), par("usr")[3], srt = 45, label = labellist, pos = 1, xpd = TRUE, cex = 0.7)
                     
                     #y-axis
                     axis(2, at = c(0, 5, 10, 15, 20), labels = c("0", "5", "10", "15", "20"), cex.axis = 0.7)
                     axis(2, at = c(25, 35, 45), labels = c("50", "60", "70"), cex.axis = 0.7, col = color.scheme[3])
                     axis(2, at = c(50, 60, 70, 80, 90, 100), labels = c("100", "200", "300", "400", "500", "600"), cex.axis = 0.7, col = color.scheme[4])
                     
                     # break axis
                     axis.break(2, 20, style = "zigzag")
                     axis.break(2, 22.5, style = "zigzag")
                     axis.break(2, 45, style = "zigzag")
                     axis.break(2, 47.5, style = "zigzag")
                     
                     #keep adding points/lines for other franchises
                     points(comp_CPM.plt$Month.Plot, comp_CPM.plt$AWC.CPM, col = color.scheme[1], cex = 1, pch = 1)
                     points(comp_CPM.plt$Month.Plot, comp_CPM.plt$CC.CPM, col = color.scheme[2], type="l", lwd = 3)
                     points(comp_CPM.plt$Month.Plot, comp_CPM.plt$CC.CPM, col = color.scheme[2], cex = 1, pch = 1)  
                     points(comp_CPM.plt$Month.Plot, comp_CPM.plt$C.CPM, col = color.scheme[5], type="l", lwd = 3)
                     points(comp_CPM.plt$Month.Plot, comp_CPM.plt$C.CPM, col = color.scheme[5], cex = 1, pch = 1)
                     
                     #Derive new scale for OC and IC
                     comp_CPM.plt$NEW.OC.CPM <- comp_CPM.plt$OC.CPM - 25
                     points(comp_CPM.plt$Month.Plot, comp_CPM.plt$NEW.OC.CPM, col = color.scheme[3], type = "l", lwd = 3)
                     points(comp_CPM.plt$Month.Plot, comp_CPM.plt$NEW.OC.CPM, col = color.scheme[3], cex = 1, pch = 1)  
                     
                     comp_CPM.plt$NEW.IC.CPM <- (as.numeric(comp_CPM.plt$IC.CPM) + 400) / 10
                     points(comp_CPM.plt$Month.Plot, comp_CPM.plt$NEW.IC.CPM,col = color.scheme[4], type = "l", lwd = 3)
                     points(comp_CPM.plt$Month.Plot, comp_CPM.plt$NEW.IC.CPM,col = color.scheme[4], cex = 1, pch = 1)  
                     
                     #Add complaint #s for the latest month
                     text(c(1:month_n), comp_CPM.plt$AWC.CPM[1:month_n] + 4, labels = comp_CPM.plt$AWC.CPM[1:month_n], col = color.scheme[1], cex = 0.6)
                     text(c(1:month_n), comp_CPM.plt$CC.CPM[1:month_n] + 3, labels = comp_CPM.plt$CC.CPM[1:month_n], col = color.scheme[2], cex = 0.6)
                     text(c(1:month_n), comp_CPM.plt$C.CPM[1:month_n] + 3, labels = comp_CPM.plt$C.CPM[1:month_n], col = color.scheme[5], cex = 0.6)
                     text(c(1:month_n), comp_CPM.plt$NEW.OC.CPM[1:month_n] + 3, labels = comp_CPM.plt$OC.CPM[1:month_n], col = color.scheme[3], cex = 0.6)
                     text(c(1:month_n), comp_CPM.plt$NEW.IC.CPM[1:month_n] + 2, labels =comp_CPM.plt$IC.CPM[1:month_n], col = color.scheme[4], cex = 0.6)
                     
                     #ADDING LEGENG
                     add_legend <- function(...){
                                           opar <- par(fig = c(0, 1, 0, 1), 
                                                       oma = c(0, 0, 0, 0), 
                                                       mar = c(0, 0, 0, 0), 
                                                       new = TRUE)
                                           on.exit(par(opar))
                                           plot(0, 0, type='n', bty='n', xaxt='n', yaxt='n')
                                           legend(...)
                     }
                     
                     # add_legend("bottom", 
                     #            legend = c("Infusion Care (IC)", "Ostomy Care (OC)", "Wound Care (WC)", "Continence Care (CC)", "Critical Care (C)"), 
                     #            pch = 1, 
                     #            text.col = c("darkorange1", "aquamarine4", "dodgerblue4", "darkgoldenrod1", "cornsilk4"),
                     #            lty = 1,
                     #            lwd = 2,
                     #            col = c("darkorange1", "aquamarine4", "dodgerblue4", "darkgoldenrod1", "cornsilk4"),
                     #            cex = 0.6,
                     #            horiz = T,
                     #            bty = "n")
                     
                     legend(month_n + 0.3, 60,
                            legend = c("Infusion Care (IC)", "Ostomy Care (OC)", "Wound Care (WC)", "Continence Care (CC)", "Critical Care (C)"),
                            pch = 1, 
                            text.col = c("darkorange1", "aquamarine4", "dodgerblue4", "darkgoldenrod1", "cornsilk4"),
                            lty = 1, 
                            lwd = 2,
                            col = c("darkorange1", "aquamarine4", "dodgerblue4", "darkgoldenrod1", "cornsilk4"),
                            cex = 0.5,  
                            horiz = F, 
                            x.intersp = 0.2,
                            bty="n")
                     return(plot)
  })
  
  Trend_Chart_all_com <- function()({if(is.null(CPM_monthly()))return(NULL)
                                     RAW.CPLT_monthly <- CPM_monthly()
                                     month_n <- input$plot_range
                                     validate(need((nrow(RAW.CPLT_monthly) - (month_n - 1)) >= 0, "Warning Message: No previous data available"))
                                     comp_RAW.CPLT.plt <- RAW.CPLT_monthly[(nrow(RAW.CPLT_monthly) - (month_n - 1)):nrow(RAW.CPLT_monthly),]
                                     comp_RAW.CPLT.plt$Month.Plot <- c(1:nrow(comp_RAW.CPLT.plt))
                                     par(xpd=TRUE, mar=par()$mar + c(-2.3, 0.5, 0, 1.7)) 
                                     color.scheme <- c("dodgerblue4", "darkgoldenrod1", "aquamarine4", "darkorange1", "cornsilk4") # corresponds to WT, CC, OC, IC and C respectively
    
                                     plot <- plot(comp_RAW.CPLT.plt$Month.Plot, 
                                                  comp_RAW.CPLT.plt$AWC.RAW.CPLT,
                                                  type = "l",
                                                  main = paste0("Complaint by Business Unit for the Past ", month_n, " Months"),
                                                  ylim = c(0, 1000),
                                                  col = color.scheme[1],
                                                  lwd = 3,
                                                  cex.main = 2,
                                                  axes = F,
                                                  ylab = expression(italic("Complaint")),
                                                  xlab = "",
                                                  cex.lab = 2,
                                                  yaxs = "i",
                                                  las = 1)
    
                                    #x-axis
                                     labellist <- comp_RAW.CPLT.plt$Month
                                     axis(side = 1, at = seq(1, month_n, by = 1), labels = FALSE)
                                     text(seq(1, month_n, by = 1), par("usr")[3]-2, srt = 45, label = labellist, pos = 1, xpd = TRUE, cex = 0.7)
    
                                    #y-axis
                                     axis(2, at = c(0, 50, 150, 250), labels = c("0", "50", "150", "250"), cex.axis = 0.6)
                                     axis(2, at = c(300, 383, 467, 633, 800, 967), labels = c("1000", "1500","2000","3000", "4000", "5000"), cex.axis = 0.6)
                                     #axis(2, at = c(550, 750, 950), labels = c("3000", "4000", "5000"), cex.axis = 0.6, col = color.scheme[4])
    
                                    #break axis
                                     axis.break(2, 250, style = "zigzag")
                                     axis.break(2, 300, style = "zigzag")
                                     #axis.break(2, 500, style = "zigzag")
                                     #axis.break(2, 550, style = "zigzag")
    
                                    #keep adding points/lines for other franchises
                                     points(comp_RAW.CPLT.plt$Month.Plot, comp_RAW.CPLT.plt$AWC.RAW.CPLT, col = color.scheme[1], cex = 1, pch = 1)
                                     points(comp_RAW.CPLT.plt$Month.Plot, comp_RAW.CPLT.plt$CC.RAW.CPLT, col = color.scheme[2], type="l", lwd = 3)
                                     points(comp_RAW.CPLT.plt$Month.Plot, comp_RAW.CPLT.plt$CC.RAW.CPLT, col = color.scheme[2], cex = 1, pch = 1)  
                                     points(comp_RAW.CPLT.plt$Month.Plot, comp_RAW.CPLT.plt$C.RAW.CPLT, col = color.scheme[5], type="l", lwd = 3)
                                     points(comp_RAW.CPLT.plt$Month.Plot, comp_RAW.CPLT.plt$C.RAW.CPLT, col = color.scheme[5], cex = 1, pch = 1)
    
                                    #Derive new scale for OC and IC
                                     comp_RAW.CPLT.plt$NEW.OC.RAW.CPLT <- (comp_RAW.CPLT.plt$OC.RAW.CPLT+800) / 6
                                     points(comp_RAW.CPLT.plt$Month.Plot, comp_RAW.CPLT.plt$NEW.OC.RAW.CPLT, col = color.scheme[3], type = "l", lwd = 3)
                                     points(comp_RAW.CPLT.plt$Month.Plot, comp_RAW.CPLT.plt$NEW.OC.RAW.CPLT, col = color.scheme[3], cex = 1, pch = 1)  
    
                                     comp_RAW.CPLT.plt$NEW.IC.RAW.CPLT <- (as.numeric(comp_RAW.CPLT.plt$IC.RAW.CPLT)+800) / 6 
                                     points(comp_RAW.CPLT.plt$Month.Plot, comp_RAW.CPLT.plt$NEW.IC.RAW.CPLT,col = color.scheme[4], type = "l", lwd = 3)
                                     points(comp_RAW.CPLT.plt$Month.Plot, comp_RAW.CPLT.plt$NEW.IC.RAW.CPLT,col = color.scheme[4], cex = 1, pch = 1)  
    
                                    #Add complaint #s for the latest month
                                     text(c(1:month_n), comp_RAW.CPLT.plt$AWC.RAW.CPLT[1:month_n] + 15, labels = comp_RAW.CPLT.plt$AWC.RAW.CPLT[1:month_n], col = color.scheme[1], cex = 0.6)
                                     text(c(1:month_n), comp_RAW.CPLT.plt$CC.RAW.CPLT[1:month_n] + 15, labels = comp_RAW.CPLT.plt$CC.RAW.CPLT[1:month_n], col = color.scheme[2], cex = 0.6)
                                     text(c(1:month_n), comp_RAW.CPLT.plt$C.RAW.CPLT[1:month_n] + 15, labels = comp_RAW.CPLT.plt$C.RAW.CPLT[1:month_n], col = color.scheme[5], cex = 0.6)
                                     text(c(1:month_n), comp_RAW.CPLT.plt$NEW.OC.RAW.CPLT[1:month_n] + 15, labels = comp_RAW.CPLT.plt$OC.RAW.CPLT[1:month_n], col = color.scheme[3], cex = 0.6)
                                     text(c(1:month_n), comp_RAW.CPLT.plt$NEW.IC.RAW.CPLT[1:month_n] + 15, labels =comp_RAW.CPLT.plt$IC.RAW.CPLT[1:month_n], col = color.scheme[4], cex = 0.6)
    
                                   #ADDING LEGENG
                                    add_legend <- function(...){
                                        opar <- par(fig = c(0, 1, 0, 1), 
                                                    oma = c(0, 0, 0, 0), 
                                                    mar = c(0, 0, 0, 0), 
                                                    new = TRUE)
                                        on.exit(par(opar))
                                        plot(0, 0, type='n', bty='n', xaxt='n', yaxt='n')
                                        legend(...)
                                   }
    
                                  # add_legend("bottom",
                                  #            legend = c("Infusion Care (IC)", "Ostomy Care (OC)", "Wound Care (WC)", "Continence Care (CC)", "Critical Care (C)"),
                                  #            pch = 1,
                                  #            text.col = c("darkorange1", "aquamarine4", "dodgerblue4", "darkgoldenrod1", "cornsilk4"),
                                  #            lty = 1,
                                  #            lwd = 2,
                                  #            col = c("darkorange1", "aquamarine4", "dodgerblue4", "darkgoldenrod1", "cornsilk4"),
                                  #            cex = 0.6,
                                  #            horiz = F,
                                  #            bty = "n")
                                  
                                  legend(month_n + 0.3, 600,
                                        legend = c("Infusion Care (IC)", "Ostomy Care (OC)", "Wound Care (WC)", "Continence Care (CC)", "Critical Care (C)"),
                                        pch = 1, 
                                        text.col = c("darkorange1", "aquamarine4", "dodgerblue4", "darkgoldenrod1", "cornsilk4"),
                                        lty = 1, 
                                        lwd = 2,
                                        col = c("darkorange1", "aquamarine4", "dodgerblue4", "darkgoldenrod1", "cornsilk4"),
                                        cex = 0.5,
                                        x.intersp = 0.2,
                                        horiz = F, 
                                        bty="n")
                                 return(plot)
                                 })
  
  Trend_Chart_NIC <- function()({if(is.null(CPM_monthly_NIC()))return(NULL)
                     CPM_monthly <- CPM_monthly_NIC()
                     CPM_yearly <- CPM_yearly_NIC()
                     target <- target()
                     target <- as.numeric(target[which(target$BU == "NIC"),]$Target)
                     pastyear1 <- format((input$YTD[1] - years(1)), "%Y")
                     baseline <- as.numeric(CPM_yearly[which(CPM_yearly$Year==pastyear1),]$CPM)
                     month_n <- input$plot_range
                     comp_CPM.plt <- CPM_monthly[(nrow(CPM_monthly) - (month_n - 1)):nrow(CPM_monthly),]
                     comp_CPM.plt$Month.Plot <- c(1:nrow(comp_CPM.plt))
                     
                     fit_NIC <- lm(CPM ~ Month.Plot, data = comp_CPM.plt)
                     m_NIC <- summary(fit_NIC)
                     
                     plot <- plot(comp_CPM.plt$Month.Plot, 
                                  comp_CPM.plt$CPM,
                                  type = "l",
                                  main = paste0("CPM of Non-Infusion Care for the Past ", month_n, " Months"),
                                  col.main = "darkcyan",
                                  ylim = c(15, 25),
                                  col = "darkcyan",
                                  lwd = 3,
                                  cex.main = 2,
                                  axes = F,
                                  ylab = expression(italic("CPM")),
                                  xlab = "",
                                  cex.lab = 2,
                                  yaxs = "i")
                     #x-axis
                     labellist <- comp_CPM.plt$Month
                     axis(side = 1, at = seq(1, month_n, by = 1), labels = FALSE)
                     text(seq(1, month_n, by = 1), par("usr")[3], srt = 45, label = labellist, pos = 1, xpd = TRUE, cex = 0.7)
                     
                     #y-axis
                     axis(side = 2, at = c(15, 20, 25), labels = c("15", "20", "25"), cex.axis = 0.7)
                     
                     text(c(1:month_n), comp_CPM.plt$CPM[1:month_n], labels = comp_CPM.plt$CPM[1:month_n], col = "darkcyan", cex = 0.7, pos = 3)
                     points(comp_CPM.plt$Month.Plot, comp_CPM.plt$CPM, col = "darkcyan", cex = 1, pch = 1) 
                     abline(h = target, col = "red", lty = 1)
                     abline(h = baseline, col = "blue", lty = 1)
                     text(x = 2, y = target - 0.4, paste0("Target CPM = ", target), cex = 0.7, col = "red")
                     text(x = 2, y = baseline + 0.3, paste0(pastyear1, " baseline = ", baseline), cex = 0.7, col = "blue")
                     lines(comp_CPM.plt$Month.Plot, fitted(fit_NIC), col = "orange", lwd = 2, lty = 2)
                     text(x = month_n - 2, y = 15.5, paste0("P-value: ", format.pval(round(pf(m_NIC$fstatistic[1], # F-statistic
                                                                                  m_NIC$fstatistic[2], # df
                                                                                  m_NIC$fstatistic[3], # df
                                                                                  lower.tail = FALSE),2))), cex = 0.7, col = "orange")
                     #text(x = 2, y = 24.6, paste0("Slope: ", round(m_NIC$coefficients[2], 2)), col = "orange")
                     #text(x = 2, y = 23.8, paste0("R square: ", round(m_NIC$r.squared, 2)), col = "orange")

                     return(plot)
  })
  
  Trend_Chart_ALL <- function()({if(is.null(CPM_monthly_all()))return(NULL)
                                 CPM_monthly <- CPM_monthly_all()
                                 CPM_yearly <- CPM_yearly_all()
                                 target <- target()
                                 target <- as.numeric(target[which(target$BU == "All BU"),]$Target)
                                 pastyear1 <- format((input$YTD[1] - years(1)), "%Y")
                                 baseline <- as.numeric(CPM_yearly[which(CPM_yearly$Year==pastyear1),]$CPM)
                                 month_n <- input$plot_range
                                 comp_CPM.plt <- CPM_monthly[(nrow(CPM_monthly) - (month_n - 1)):nrow(CPM_monthly),]
                                 comp_CPM.plt$Month.Plot <- c(1:nrow(comp_CPM.plt))
    
                                 fit_ALL <- lm(CPM ~ Month.Plot, data = comp_CPM.plt)
                                 m_ALL <- summary(fit_ALL)
    
                                 plot <- plot(comp_CPM.plt$Month.Plot, 
                                              comp_CPM.plt$CPM,
                                              type = "l",
                                              main = paste0("CPM of All B.U. Combined for the Past ", month_n, " Months"),
                                              col.main = "coral4",
                                              ylim = c(30, 75),
                                              col = "coral4",
                                              lwd = 3,
                                              cex.main = 2,
                                              axes = F,
                                              ylab = expression(italic("CPM")),
                                              xlab = "",
                                              cex.lab = 2,
                                              yaxs = "i")
                                #x-axis
                                 labellist <- comp_CPM.plt$Month
                                 axis(side = 1, at = seq(1, month_n, by = 1), labels = FALSE)
                                 text(seq(1, month_n, by = 1), par("usr")[3], srt = 45, label = labellist, pos = 1, xpd = TRUE, cex = 0.7)
    
                                #y-axis
                                 axis(side = 2, at = c(30, 35, 45, 55, 65, 75), labels = c("30", "35", "45", "55", "65", "75"), cex.axis = 0.7)
    
                                 text(c(1:month_n), comp_CPM.plt$CPM[1:month_n], labels = comp_CPM.plt$CPM[1:month_n], col = "coral4", cex = 0.7, pos = 3)
                                 points(comp_CPM.plt$Month.Plot, comp_CPM.plt$CPM, col = "coral4", cex = 1, pch = 1) 
                                 abline(h = target, col = "red", lty = 1)
                                 abline(h = baseline, col = "blue", lty = 1)
                                 text(x = 2, y = target - 2, paste0("Target CPM = ", target), cex = 0.7, col = "red")
                                 text(x = 1.8, y = baseline + 1.5, paste0(pastyear1, " baseline = ", baseline), cex = 0.7, col = "blue")
                                 lines(comp_CPM.plt$Month.Plot, fitted(fit_ALL), col = "orange", lwd = 2, lty = 2)
                                 #text(x = 2, y = 78.7, paste0("R square: ", round(m_ALL$r.squared, 2)), col = "orange")
                                 text(x = month_n - 2, y = 35, paste0("P-value: ", format.pval(round(pf(m_ALL$fstatistic[1], # F-statistic
                                                                                     m_ALL$fstatistic[2], # df
                                                                                     m_ALL$fstatistic[3], # df
                                                                                     lower.tail = FALSE),2))), cex = 0.7, col = "orange")
                                 #text(x = 2, y = 81, paste0("Slope: ", round(m_ALL$coefficients[2], 2)), col = "orange")
    
                                 return(plot)
  })
    
  
  Trend_Chart_OC <- function()({if(is.null(CPM_monthly()))return(NULL)
                     CPM_monthly <- CPM_monthly()
                     CPM_yearly <- CPM_yearly()
                     target <- target()
                     target <- as.numeric(target[which(target$BU == "OC"),]$Target)
                     pastyear1 <- format((input$YTD[1] - years(1)), "%Y")
                     baseline <- as.numeric(CPM_yearly[which(CPM_yearly$Year==pastyear1),]$OC.CPM)
                     month_n <- input$plot_range
                     comp_CPM.plt <- CPM_monthly[(nrow(CPM_monthly) - (month_n - 1)):nrow(CPM_monthly),]
                     comp_CPM.plt$Month.Plot <- c(1:nrow(comp_CPM.plt))
                     
                     fit_OC <- lm(OC.CPM ~ Month.Plot, data = comp_CPM.plt)
                     m_OC <- summary(fit_OC)
    
                     plot <- plot(comp_CPM.plt$Month.Plot, 
                                  comp_CPM.plt$OC.CPM,
                                  type = "l",
                                  main = paste0("CPM of Ostomy Care for the Past ", month_n, " Months"),
                                  col.main = "aquamarine4",
                                  ylim = c(45, 70),
                                  col = "aquamarine4",
                                  lwd = 3,
                                  cex.main = 2,
                                  axes = F,
                                  ylab = expression(italic("CPM")),
                                  xlab = "",
                                  cex.lab = 2,
                                  yaxs = "i")
                     #x-axis
                     labellist <- comp_CPM.plt$Month
                     axis(side = 1, at = seq(1, month_n, by = 1), labels = FALSE)
                     text(seq(1, month_n, by = 1), par("usr")[3], srt = 45, label = labellist, pos = 1, xpd = TRUE, cex = 0.7)
    
                     #y-axis
                     axis(side = 2, at = c(45, 50, 55, 60, 65, 70), 
                          labels = c("45", "50", "55", "60", "65", "70"), cex.axis = 0.7)
    
                     text(c(1:month_n), comp_CPM.plt$OC.CPM[1:month_n], labels = comp_CPM.plt$OC.CPM[1:month_n], col = "aquamarine4", cex = 0.7, pos = 3)
                     points(comp_CPM.plt$Month.Plot, comp_CPM.plt$OC.CPM, col = "black", cex = 1, pch = 1)
                     abline(h = target, col = "red", lty = 1)
                     abline(h = baseline, col = "blue", lty = 1)
                     text(x = 2, y = target - 1.6, paste0("Target CPM = ", target), cex = 0.7, col = "red")
                     text(x = 2, y = baseline + 1.4, paste0(pastyear1, " baseline = ", baseline), cex = 0.7, col = "blue")
                     lines(comp_CPM.plt$Month.Plot, fitted(fit_OC), col = "orange", lwd = 2, lty = 2)
                     #text(x = 2, y = 68.5, paste0("R square: ", round(m_OC$r.squared, 2)), col = "orange")
                     text(x = month_n - 2, y = 43, paste0("P-value: ", format.pval(round(pf(m_OC$fstatistic[1], # F-statistic
                                                                                  m_OC$fstatistic[2], # df
                                                                                  m_OC$fstatistic[3], # df
                                                                                  lower.tail = FALSE),2))), cex = 0.7, col = "orange")
                     #text(x = 2, y = 71, paste0("Slope: ", round(m_OC$coefficients[2],2)), col = "orange")
                     return(plot)
  })
  
  Trend_Chart_CC <- function()({if(is.null(CPM_monthly()))return(NULL)
                     CPM_monthly <- CPM_monthly()
                     CPM_yearly <- CPM_yearly()
                     target <- target()
                     target <- as.numeric(target[which(target$BU == "CC"),]$Target)
                     pastyear1 <- format((input$YTD[1] - years(1)), "%Y")
                     baseline <- as.numeric(CPM_yearly[which(CPM_yearly$Year==pastyear1),]$CC.CPM)
                     month_n <- input$plot_range
                     comp_CPM.plt <- CPM_monthly[(nrow(CPM_monthly) - (month_n - 1)):nrow(CPM_monthly),]
                     comp_CPM.plt$Month.Plot <- c(1:nrow(comp_CPM.plt))
                     fit_CC <- lm(CC.CPM ~ Month.Plot, data = comp_CPM.plt)
                     m_CC <- summary(fit_CC)
    
                     plot <- plot(comp_CPM.plt$Month.Plot, 
                                  comp_CPM.plt$CC.CPM,
                                  type = "l",
                                  main = paste0("CPM of Continence Care for the Past ", month_n, " Months"),
                                  col.main = "darkgoldenrod1",
                                  ylim = c(0, 5),
                                  col = "darkgoldenrod1",
                                  lwd = 3,
                                  
                                  
                                  cex.main = 2,
                                  axes = F,
                                  ylab = expression(italic("CPM")),
                                  xlab = "",
                                  cex.lab = 2,
                                  yaxs = "i")
                                  #x-axis
                                  labellist <- comp_CPM.plt$Month
                                  axis(side = 1, at = seq(1, month_n, by = 1), labels = FALSE)
                                  text(seq(1, month_n, by = 1), par("usr")[3], srt = 45, label = labellist, pos = 1, xpd = TRUE, cex = 0.7)
    
                                  #y-axis
                                  axis(side = 2, at = c(0, 1, 2, 3, 4, 5), 
                                  labels = c("0", "1", "2", "3", "4", "5"), cex.axis = 0.7)
    
                                  text(c(1:month_n), comp_CPM.plt$CC.CPM[1:month_n], labels = comp_CPM.plt$CC.CPM[1:month_n], col = "darkgoldenrod1", cex = 0.7, pos = 3)
                                  points(comp_CPM.plt$Month.Plot, comp_CPM.plt$CC.CPM, col = "darkgoldenrod1", cex = 1, pch = 1)
                                  abline(h = target, col = "red", lty = 1)
                                  abline(h = baseline, col = "blue", lty = 1)
                                  text(x = 2, y = target - 0.3, paste0("Target CPM = ", target), cex = 0.7, col = "red")
                                  text(x = 2, y = baseline + 0.3, paste0(pastyear1, " baseline = ", baseline), cex = 0.7, col = "blue")
                                  lines(comp_CPM.plt$Month.Plot, fitted(fit_CC), col = "orange", lwd = 2, lty = 2)
                                  #text(x = 2, y = 4.05, paste0("R square: ", round(m_CC$r.squared, 2)), col = "orange")
                                  text(x = month_n - 2, y = 0.4, paste0("P-value: ", format.pval(round(pf(m_CC$fstatistic[1], # F-statistic
                                                                                                m_CC$fstatistic[2], # df
                                                                                                m_CC$fstatistic[3], # df
                                                                                                lower.tail = FALSE),2))), cex = 0.7, col = "orange")
                                  #text(x = 2, y = 4.5, paste0("Slope: ", round(m_CC$coefficients[2],2)), col = "orange")
                                  return(plot)
  })
  
  Trend_Chart_C <- function()({if(is.null(CPM_monthly()))return(NULL)
                               CPM_monthly <- CPM_monthly()
                               CPM_yearly <- CPM_yearly()
                               target <- target()
                               target <- as.numeric(target[which(target$BU == "C"),]$Target)
                               pastyear1 <- format((input$YTD[1] - years(1)), "%Y")
                               baseline <- as.numeric(CPM_yearly[which(CPM_yearly$Year==pastyear1),]$C.CPM)
                               month_n <- input$plot_range
                               comp_CPM.plt <- CPM_monthly[(nrow(CPM_monthly) - (month_n - 1)):nrow(CPM_monthly),]
                               comp_CPM.plt$Month.Plot <- c(1:nrow(comp_CPM.plt))
                               fit_C <- lm(C.CPM ~ Month.Plot, data = comp_CPM.plt)
                               m_C <- summary(fit_C)
    
                               plot <- plot(comp_CPM.plt$Month.Plot, 
                                       comp_CPM.plt$C.CPM,
                                       type = "l",
                                       main = paste0("CPM of Critical Care for the Past ", month_n, " Months"),
                                       col.main = "cornsilk4",
                                       ylim = c(0.5, 3),
                                       col = "cornsilk4",
                                       lwd = 3,
                                       cex.main = 2,
                                       axes = F,
                                       ylab = expression(italic("CPM")),
                                       xlab = "",
                                       cex.lab = 2,
                                       yaxs = "i")
                               #x-axis
                               labellist <- comp_CPM.plt$Month
                               axis(side = 1, at = seq(1, month_n, by = 1), labels = FALSE)
                               text(seq(1, month_n, by = 1), par("usr")[3], srt = 45, label = labellist, pos = 1, xpd = TRUE, cex = 0.7)
    
                               #y-axis
                               axis(side = 2, at = c(0.5, 1, 1.5, 2, 2.5, 3), labels = c("0.5", "1", "1.5", "2", "2.5", "3"), cex.axis = 0.7)
    
                               text(c(1:month_n), comp_CPM.plt$C.CPM[1:month_n], labels = comp_CPM.plt$C.CPM[1:month_n], col = "cornsilk4", cex = 0.7, pos = 3)
                               points(comp_CPM.plt$Month.Plot, comp_CPM.plt$C.CPM, col = "cornsilk4", cex = 1, pch = 1)
                               abline(h = target, col = "red", lty = 1)
                               abline(h = baseline, col = "blue", lty = 1)
                               text(x = 2, y = target - 0.1, paste0("Target CPM = ", target), cex = 0.7, col = "red")
                               text(x = 2, y = baseline + 0.08, paste0(pastyear1, " baseline = ", baseline), cex = 0.7, col = "blue")
                               lines(comp_CPM.plt$Month.Plot, fitted(fit_C), col = "orange", lwd = 2, lty = 2)
                               #text(x = 2, y = 2.2, paste0("R square: ", round(m_C$r.squared, 2)), col = "orange")
                               text(x = month_n - 2, y = 0.6, paste0("P-value: ", format.pval(round(pf(m_C$fstatistic[1], # F-statistic
                                                                                             m_C$fstatistic[2], # df
                                                                                             m_C$fstatistic[3], # df
                                                                                             lower.tail = FALSE),2))), cex = 0.7, col = "orange")
                               #text(x = 2, y = 2.4, paste0("Slope: ", round(m_C$coefficients[2],2)), col = "orange")
                               return(plot)
  })
  
  Trend_Chart_AWC <- function()({if(is.null(CPM_monthly()))return(NULL)
                     CPM_monthly <- CPM_monthly()
                     CPM_yearly <- CPM_yearly()
                     target <- target()
                     target <- as.numeric(target[which(target$BU == "AWC"),]$Target)
                     pastyear1 <- format((input$YTD[1] - years(1)), "%Y")
                     baseline <- as.numeric(CPM_yearly[which(CPM_yearly$Year==pastyear1),]$AWC.CPM)
                     month_n <- input$plot_range
                     comp_CPM.plt <- CPM_monthly[(nrow(CPM_monthly) - (month_n - 1)):nrow(CPM_monthly),]
                     comp_CPM.plt$Month.Plot <- c(1:nrow(comp_CPM.plt))
                     fit_AWC <- lm(AWC.CPM ~ Month.Plot, data = comp_CPM.plt)
                     m_AWC <- summary(fit_AWC)
    
                     plot <- plot(comp_CPM.plt$Month.Plot, 
                                  comp_CPM.plt$AWC.CPM,
                                  type = "l",
                                  main = paste0("CPM of Wound Care for the Past ", month_n, " Months"),
                                  col.main = "dodgerblue4",
                                  ylim = c(5, 17),
                                  col = "dodgerblue4",
                                  lwd = 3,
                                  cex.main = 2,
                                  axes = F,
                                  ylab = expression(italic("CPM")),
                                  xlab = "",
                                  cex.lab = 2,
                                  yaxs = "i")
                    #x-axis
                    labellist <- comp_CPM.plt$Month
                    axis(side = 1, at = seq(1, month_n, by = 1), labels = FALSE)
                    text(seq(1, month_n, by = 1), par("usr")[3], srt = 45, label = labellist, pos = 1, xpd = TRUE, cex = 0.7)
    
                    #y-axis
                    axis(side = 2, at = c(5, 11, 17), 
                    labels = c("5", "11", "17"), cex.axis = 0.7)
                    abline(h = target, col = "red", lty = 1)
                    abline(h = baseline, col = "blue", lty = 1)
                    text(c(1:month_n), comp_CPM.plt$AWC.CPM[1:month_n], labels = comp_CPM.plt$AWC.CPM[1:month_n], col = "dodgerblue4", cex = 0.7, pos = 3)
                    points(comp_CPM.plt$Month.Plot, comp_CPM.plt$AWC.CPM, col = "dodgerblue4", cex = 1, pch = 1)
                    text(x = 2, y = target - 0.6, paste0("Target CPM = ", target), cex = 0.7, col = "red")
                    text(x = 2, y = baseline + 0.5, paste0(pastyear1, " baseline = ", baseline), cex = 0.7, col = "blue")
                    lines(comp_CPM.plt$Month.Plot, fitted(fit_AWC), col = "orange", lwd = 2, lty = 2)
                    #text(x = 2, y = 17.7, paste0("R square: ", round(m_AWC$r.squared, 2)), col = "orange")
                    text(x = month_n - 2, y = 5.9, paste0("P-value: ", format.pval(round(pf(m_AWC$fstatistic[1], # F-statistic
                                                                                   m_AWC$fstatistic[2], # df
                                                                                   m_AWC$fstatistic[3], # df
                                                                                   lower.tail = FALSE),2))), cex = 0.7, col = "orange")
                    #text(x = 2, y = 19, paste0("Slope: ", round(m_AWC$coefficients[2],2)), col = "orange")
                    return(plot)
  })
  
  
  Trend_Chart_IC <- function()({if(is.null(CPM_monthly()))return(NULL)
                    CPM_monthly <- CPM_monthly()
                    CPM_yearly <- CPM_yearly()
                    target <- target()
                    target <- as.numeric(target[which(target$BU == "IC"),]$Target)
                    pastyear1 <- format((input$YTD[1] - years(1)), "%Y")
                    baseline <- as.numeric(CPM_yearly[which(CPM_yearly$Year==pastyear1),]$IC.CPM)
                    month_n <- input$plot_range
                    comp_CPM.plt <- CPM_monthly[(nrow(CPM_monthly) - (month_n - 1)):nrow(CPM_monthly),]
                    comp_CPM.plt$Month.Plot <- c(1:nrow(comp_CPM.plt))
                    fit_IC <- lm(IC.CPM ~ Month.Plot, data = comp_CPM.plt)
                    m_IC <- summary(fit_IC)
    
                    plot <- plot(comp_CPM.plt$Month.Plot, 
                                 comp_CPM.plt$IC.CPM,
                                 type = "l",
                                 main = paste0("CPM of Infusion Care for the Past ", month_n, " Months"),
                                 col.main = "darkorange1",
                                 ylim = c(100, 600),
                                 col = "darkorange1",
                                 lwd = 3,
                                 cex.main = 2,
                                 axes = F,
                                 ylab = expression(italic("CPM")),
                                 xlab = "",
                                 cex.lab = 2,
                                 yaxs = "i")
                   #x-axis
                   labellist <- comp_CPM.plt$Month
                   axis(side = 1, at = seq(1, month_n, by = 1), labels = FALSE)
                   text(seq(1, month_n, by = 1), par("usr")[3], srt = 45, label = labellist, pos = 1, xpd = TRUE, cex = 0.7)
    
                   #y-axis
                   axis(side = 2, at = c(100, 150, 200, 250, 300, 350, 400, 450, 500, 550, 600), 
                                  labels = c("100", "150", "200", "250", "300", "350", "400", "450", "500", "550", "600"), cex.axis = 0.7)
                   abline(h = target, col = "red", lty = 1)
                   abline(h = baseline, col = "blue", lty = 1)
                   text(c(1:month_n), comp_CPM.plt$IC.CPM[1:month_n], labels = comp_CPM.plt$IC.CPM[1:month_n], col = "darkorange1", cex = 0.7, pos = 3)
                   points(comp_CPM.plt$Month.Plot, comp_CPM.plt$IC.CPM, col = "darkorange1", cex = 1, pch = 1)
                   text(x = 2, y = target - 10, paste0("Target CPM = ", target), cex = 0.7, col = "red")
                   text(x = 2.2, y = baseline + 10, paste0(pastyear1, " baseline = ", baseline), cex = 0.7, col = "blue")
                   lines(comp_CPM.plt$Month.Plot, fitted(fit_IC), col = "orange", lwd = 2, lty = 2)
                   #text(x = 2, y = 572, paste0("R square: ", round(m_IC$r.squared, 2)), col = "orange")
                   text(x = month_n - 2, y = 315, paste0("P-value: ", format.pval(round(pf(m_IC$fstatistic[1], # F-statistic
                                                                                 m_IC$fstatistic[2], # df
                                                                                 m_IC$fstatistic[3], # df
                                                                                 lower.tail = FALSE),2))), cex = 0.7, col = "orange")
                   #text(x = 2, y = 595, paste0("Slope: ", round(m_IC$coefficients[2],2)), col = "orange")
                   return(plot)
  })
  
  
  Trend_Chart_Eurotec <- function()({if(is.null(CPM_monthly()))return(NULL)
                                        CPM_monthly <- CPM_monthly()
                                        CPM_yearly <- CPM_yearly()
                                        target <- target()
                                        target <- as.numeric(target[which(target$BU == "(Eurotec)"),]$Target)
                                        pastyear1 <- format((input$YTD[1] - years(1)), "%Y")
                                        baseline <- as.numeric(CPM_yearly[which(CPM_yearly$Year==pastyear1),]$Eurotec.CPM)
                                        month_n <- input$plot_range
                                        comp_CPM.plt <- CPM_monthly[(nrow(CPM_monthly) - (month_n - 1)):nrow(CPM_monthly),]
                                        comp_CPM.plt$Month.Plot <- c(1:nrow(comp_CPM.plt))
                                        fit_Eurotec <- lm(Eurotec.CPM ~ Month.Plot, data = comp_CPM.plt)
                                        m_Eurotec <- summary(fit_Eurotec)
    
                                        plot <- plot(comp_CPM.plt$Month.Plot, 
                                                     comp_CPM.plt$Eurotec.CPM,
                                                     type = "l",
                                                     main = paste0("CPM of Eurotec for the Past ", month_n, " Months"),
                                                     col.main = "aquamarine3",
                                                     ylim = c(0, 100),
                                                     col = "aquamarine3",
                                                     lwd = 3,
                                                     cex.main = 2,
                                                     axes = F,
                                                     ylab = expression(italic("CPM")),
                                                     xlab = "",
                                                     cex.lab = 2,
                                                     yaxs = "i")
                                         #x-axis
                                         labellist <- comp_CPM.plt$Month
                                         axis(side = 1, at = seq(1, month_n, by = 1), labels = FALSE)
                                         text(seq(1, month_n, by = 1), par("usr")[3], srt = 45, label = labellist, pos = 1, xpd = TRUE, cex = 0.7)
    
                                        #y-axis
                                         axis(side = 2, at = c(0, 50, 100), 
                                         labels = c("0", "50", "100"), cex.axis = 0.7)
                                         abline(h = target, col = "red", lty = 1)
                                         abline(h = baseline, col = "blue", lty = 1)
                                         text(c(1:month_n), comp_CPM.plt$Eurotec.CPM[1:month_n], labels = comp_CPM.plt$Eurotec.CPM[1:month_n], col = "darkorange1", cex = 0.7, pos = 3)
                                         points(comp_CPM.plt$Month.Plot, comp_CPM.plt$Eurotec.CPM, col = "darkorange1", cex = 1, pch = 1)
                                         text(x = 2, y = target - 10, paste0("Target CPM = ", target), cex = 0.7, col = "red")
                                         text(x = 2.2, y = baseline + 10, paste0(pastyear1, " baseline = ", baseline), cex = 0.7, col = "blue")
                                         lines(comp_CPM.plt$Month.Plot, fitted(fit_Eurotec), col = "orange", lwd = 2, lty = 2)
                                        #text(x = 2, y = 572, paste0("R square: ", round(m_Eurotec$r.squared, 2)), col = "orange")
                                         text(x = month_n - 2, y = 315, paste0("P-value: ", format.pval(round(pf(m_Eurotec$fstatistic[1], # F-statistic
                                                                                                        m_Eurotec$fstatistic[2], # df
                                                                                                        m_Eurotec$fstatistic[3], # df
                                                                                                        lower.tail = FALSE),2))), cex = 0.7, col = "orange")
                                       #text(x = 2, y = 595, paste0("Slope: ", round(m_Eurotec$coefficients[2],2)), col = "orange")
                                        return(plot)
  })
  
  Trend_Chart6 <- function()({
                  month_n <- input$plot_range
                  CPM_monthly_all <- CPM_monthly_all()
                  comp_CPM_all.plt <- CPM_monthly_all[(nrow(CPM_monthly_all) - (month_n - 1)):nrow(CPM_monthly_all),]
                  comp_CPM_all.plt$Month.Plot <- c(1:nrow(comp_CPM_all.plt))
    
                  CPM_monthly <- CPM_monthly()
                  comp_CPM.plt <- CPM_monthly[(nrow(CPM_monthly) - (month_n - 1)):nrow(CPM_monthly),]
                  comp_CPM.plt$Month.Plot <- c(1:nrow(comp_CPM.plt))
                  
                  target <- target()
                  target_ALL <- as.numeric(target[which(target$BU == "All BU"),]$Target)
                  target_IC <- as.numeric(target[which(target$BU == "IC"),]$Target)
                  target_OC <- as.numeric(target[which(target$BU == "OC"),]$Target)
                  target_AWC <- as.numeric(target[which(target$BU == "AWC"),]$Target)
                  target_CC <- as.numeric(target[which(target$BU == "CC"),]$Target)
                  target_C <- as.numeric(target[which(target$BU == "C"),]$Target)
    
                  
                  fit_ALL <- lm(CPM ~ Month.Plot, data = comp_CPM_all.plt)
                  fit_IC <- lm(IC.CPM ~ Month.Plot, data = comp_CPM.plt)
                  fit_OC <- lm(OC.CPM ~ Month.Plot, data = comp_CPM.plt)
                  fit_AWC <- lm(AWC.CPM ~ Month.Plot, data = comp_CPM.plt)
                  fit_CC <- lm(CC.CPM ~ Month.Plot, data = comp_CPM.plt)
                  fit_C <- lm(C.CPM ~ Month.Plot, data = comp_CPM.plt)
                  
                  m_ALL <- summary(fit_ALL)
                  m_IC <- summary(fit_IC)
                  m_OC <- summary(fit_OC)
                  m_AWC <- summary(fit_AWC)
                  m_CC <- summary(fit_CC)
                  m_C <- summary(fit_C)
                  
                  labellist <- comp_CPM.plt$Month
                  
                  layout(matrix(c(1, 2, 3, 4, 5, 6), 3, 2, byrow = TRUE))   


                  plot_all <- plot(comp_CPM_all.plt$Month.Plot, 
                                   comp_CPM_all.plt$CPM,
                                   type = "l",
                                   main = paste0("All B.U. Combined"),
                                   col.main = "coral4",
                                   ylim = c(35, 75),
                                   col = "coral4",
                                   lwd = 3,
                                   cex.main = 1,
                                   axes = F,
                                   ylab = expression(italic("CPM")),
                                   xlab = "",
                                   cex.lab = 1,
                                   yaxs = "i")
                  #x-axis
                  axis(side = 1, at = seq(1, month_n, by = 1), labels = FALSE)
                  text(seq(1, month_n, by = 1), par("usr")[3], srt = 45, label = labellist, pos = 1, xpd = TRUE, cex = 0.7)
                  
                  #y-axis
                  axis(side = 2, at = c(35, 45, 55, 65, 75), labels = c("35", "45", "55", "65", "75"), cex.axis = 0.7)
                  
                  text(c(1:month_n), comp_CPM_all.plt$CPM[1:month_n], labels = comp_CPM_all.plt$CPM[1:month_n], col = "coral4", cex = 0.7, pos = 3)
                  points(comp_CPM_all.plt$Month.Plot, comp_CPM_all.plt$CPM, col = "coral4", cex = 1, pch = 1) 
                  abline(h = target_ALL, col = "red", lty = 1)
                  text(x = 2, y = target_ALL - 2, paste0("Target CPM = ", target_ALL), col = "red", cex = 0.7)
                  lines(comp_CPM.plt$Month.Plot, fitted(fit_ALL), col = "orange", lwd = 2, lty = 2)
                  text(x = 2, y = 67, paste0("R square: ", round(m_ALL$r.squared, 2)), col = "orange", cex = 0.7)
                  text(x = 2, y = 70, paste0("P-value: ", format.pval(round(pf(m_ALL$fstatistic[1], # F-statistic
                                                                               m_ALL$fstatistic[2], # df
                                                                               m_ALL$fstatistic[3], # df
                                                                               lower.tail = FALSE),2))), col = "orange", cex = 0.8)
                  text(x = 2, y = 72, paste0("Slope: ", round(m_ALL$coefficients[2], 2)), col = "orange", cex = 0.7)
    
                  plot_IC <- plot(comp_CPM.plt$Month.Plot, 
                                  comp_CPM.plt$IC.CPM,
                                  type = "l",
                                  main = paste0("Infusion Care"),
                                  col.main = "darkorange1",
                                  ylim = c(150, 500),
                                  col = "darkorange1",
                                  lwd = 3,
                                  cex.main = 1,
                                  axes = F,
                                  ylab = expression(italic("CPM")),
                                  xlab = "",
                                  cex.lab = 1,
                                  yaxs = "i")
                  #x-axis
                  axis(side = 1, at = seq(1, month_n, by = 1), labels = FALSE)
                  text(seq(1, month_n, by = 1), par("usr")[3], srt = 45, label = labellist, pos = 1, xpd = TRUE, cex = 0.8)
                  
                  #y-axis
                  axis(side = 2, at = c(150, 200, 250, 300, 350, 400, 450, 500), 
                       labels = c("150", "200", "250", "300", "350", "400", "450", "500"), cex.axis = 0.7)
                  abline(h = target_IC, col = "red", lty = 1)
                  text(c(1:month_n), comp_CPM.plt$IC.CPM[1:month_n], labels = comp_CPM.plt$IC.CPM[1:month_n], col = "darkorange1", cex = 0.7, pos = 3)
                  points(comp_CPM.plt$Month.Plot, comp_CPM.plt$IC.CPM, col = "darkorange1", cex = 1, pch = 1)
                  text(x = 2, y = target_IC - 16, paste0("Target CPM = ", target_IC), col = "red", cex = 0.7)
                  lines(comp_CPM.plt$Month.Plot, fitted(fit_IC), col = "orange", lwd = 2, lty = 2)
                  text(x = 2, y = 460, paste0("R square: ", round(m_IC$r.squared, 2)), col = "orange", cex = 0.7)
                  text(x = 2, y = 475, paste0("P-value: ", format.pval(round(pf(m_IC$fstatistic[1], # F-statistic
                                                                                m_IC$fstatistic[2], # df
                                                                                m_IC$fstatistic[3], # df
                                                                                lower.tail = FALSE),2))), col = "orange", cex = 0.7)
                  text(x = 2, y = 490, paste0("Slope: ", round(m_IC$coefficients[2],2)), col = "orange", cex = 0.7)
    
    
                  
                  plot_OC <- plot(comp_CPM.plt$Month.Plot, 
                               comp_CPM.plt$OC.CPM,
                               type = "l",
                               main = paste0("Ostomy"),
                               col.main = "aquamarine4",
                               ylim = c(40, 75),
                               col = "aquamarine4",
                               lwd = 3,
                               cex.main = 1,
                               axes = F,
                               ylab = expression(italic("CPM")),
                               xlab = "",
                               cex.lab = 1,
                               yaxs = "i")
                  #x-axis
                  axis(side = 1, at = seq(1, month_n, by = 1), labels = FALSE)
                  text(seq(1, month_n, by = 1), par("usr")[3], srt = 45, label = labellist, pos = 1, xpd = TRUE, cex = 0.8)
                  
                  #y-axis
                  axis(side = 2, at = c(40, 45, 50, 55, 60, 65, 70), 
                       labels = c("40", "45", "50", "55", "60", "65", "70"), cex.axis = 0.7)
                  
                  text(c(1:month_n), comp_CPM.plt$OC.CPM[1:month_n], labels = comp_CPM.plt$OC.CPM[1:month_n], col = "aquamarine4", cex = 0.7, pos = 3)
                  points(comp_CPM.plt$Month.Plot, comp_CPM.plt$OC.CPM, col = "black", cex = 1, pch = 1)
                  abline(h = target_OC, col = "red", lty = 1)
                  text(x = 2, y = target_OC - 1.8, paste0("Target CPM = ", target_OC), col = "red", cex = 0.7)
                  lines(comp_CPM.plt$Month.Plot, fitted(fit_OC), col = "orange", lwd = 2, lty = 2)
                  text(x = 2, y = 63, paste0("R square: ", round(m_OC$r.squared, 2)), col = "orange", cex = 0.7)
                  text(x = 2, y =65.5, paste0("P-value: ", format.pval(round(pf(m_OC$fstatistic[1], # F-statistic
                                                                              m_OC$fstatistic[2], # df
                                                                              m_OC$fstatistic[3], # df
                                                                              lower.tail = FALSE),2))), col = "orange", cex = 0.7)
                  text(x = 2, y = 68, paste0("Slope: ", round(m_OC$coefficients[2],2)), col = "orange", cex = 0.7)
                  
                  
                  plot_AWC <- plot(comp_CPM.plt$Month.Plot, 
                               comp_CPM.plt$AWC.CPM,
                               type = "l",
                               main = paste0("Wound Care"),
                               col.main = "dodgerblue4",
                               ylim = c(5, 20),
                               col = "dodgerblue4",
                               lwd = 3,
                               cex.main = 1,
                               axes = F,
                               ylab = expression(italic("CPM")),
                               xlab = "",
                               cex.lab = 1,
                               yaxs = "i")
                  #x-axis
                  axis(side = 1, at = seq(1, month_n, by = 1), labels = FALSE)
                  text(seq(1, month_n, by = 1), par("usr")[3], srt = 45, label = labellist, pos = 1, xpd = TRUE, cex = 0.8)
                  
                  #y-axis
                  axis(side = 2, at = c(5, 10, 15, 20), 
                       labels = c("5", "10", "15", "20"), cex.axis = 0.7)
                  abline(h = target_AWC, col = "red", lty = 1)
                  text(c(1:month_n), comp_CPM.plt$AWC.CPM[1:month_n], labels = comp_CPM.plt$AWC.CPM[1:month_n], col = "dodgerblue4", cex = 0.7, pos = 3)
                  points(comp_CPM.plt$Month.Plot, comp_CPM.plt$AWC.CPM, col = "dodgerblue4", cex = 1, pch = 1)
                  text(x = 2, y = target_AWC - 1, paste0("Target CPM = ", target_AWC), col = "red", cex = 0.7)
                  lines(comp_CPM.plt$Month.Plot, fitted(fit_AWC), col = "orange", lwd = 2, lty = 2, cex = 0.7)
                  text(x = 2, y = 17, paste0("R square: ", round(m_AWC$r.squared, 2)), col = "orange", cex = 0.7)
                  text(x = 2, y = 18.3, paste0("P-value: ", format.pval(round(pf(m_AWC$fstatistic[1], # F-statistic
                                                                                 m_AWC$fstatistic[2], # df
                                                                                 m_AWC$fstatistic[3], # df
                                                                                 lower.tail = FALSE),2))), col = "orange", cex = 0.7)
                  text(x = 2, y = 19.1, paste0("Slope: ", round(m_AWC$coefficients[2],2)), col = "orange", cex = 0.7)
    
    
                  plot_CC <- plot(comp_CPM.plt$Month.Plot, 
                               comp_CPM.plt$CC.CPM,
                               type = "l",
                               main = paste0("Continence Care"),
                               col.main = "darkgoldenrod1",
                               ylim = c(0, 7),
                               col = "darkgoldenrod1",
                               lwd = 3,
                               cex.main = 1,
                               axes = F,
                               ylab = expression(italic("CPM")),
                               xlab = "",
                               cex.lab = 1,
                               yaxs = "i")
                  #x-axis
                  axis(side = 1, at = seq(1, month_n, by = 1), labels = FALSE)
                  text(seq(1, month_n, by = 1), par("usr")[3], srt = 45, label = labellist, pos = 1, xpd = TRUE, cex = 0.8)
                  
                  #y-axis
                  axis(side = 2, at = c(0, 1, 2, 3, 4, 5, 6), 
                       labels = c("0","1", "2", "3", "4", "5", "6"), cex.axis = 0.7)
                  
                  text(c(1:month_n), comp_CPM.plt$CC.CPM[1:month_n], labels = comp_CPM.plt$CC.CPM[1:month_n], col = "darkgoldenrod1", cex = 0.7, pos = 3)
                  points(comp_CPM.plt$Month.Plot, comp_CPM.plt$CC.CPM, col = "darkgoldenrod1", cex = 1, pch = 1)
                  abline(h = target_CC, col = "red", lty = 1)
                  text(x = 2, y = target_CC - 1, paste0("Target CPM = ", target_CC), col = "red", cex = 0.7)
                  lines(comp_CPM.plt$Month.Plot, fitted(fit_CC), col = "orange", lwd = 2, lty = 2)
                  text(x = 2, y = 4, paste0("R square: ", round(m_CC$r.squared, 2)), col = "orange", cex = 0.7)
                  text(x = 2, y = 4.5, paste0("P-value: ", format.pval(round(pf(m_CC$fstatistic[1], # F-statistic
                                                                                m_CC$fstatistic[2], # df
                                                                                m_CC$fstatistic[3], # df
                                                                                lower.tail = FALSE),2))),col = "orange", cex = 0.7)
                  text(x = 2, y = 5, paste0("Slope: ", round(m_CC$coefficients[2],2)), col = "orange", cex = 0.7)
                  
                  
                  plot_C <- plot(comp_CPM.plt$Month.Plot, 
                               comp_CPM.plt$C.CPM,
                               type = "l",
                               main = paste0("Critical Care"),
                               col.main = "cornsilk4",
                               ylim = c(0.5, 2.5),
                               col = "cornsilk4",
                               lwd = 3,
                               cex.main = 1,
                               axes = F,
                               ylab = expression(italic("CPM")),
                               xlab = "",
                               cex.lab = 1,
                               yaxs = "i")
                  #x-axis
                  axis(side = 1, at = seq(1, month_n, by = 1), labels = FALSE)
                  text(seq(1, month_n, by = 1), par("usr")[3], srt = 45, label = labellist, pos = 1, xpd = TRUE, cex = 0.8)
                  
                  #y-axis
                  axis(side = 2, at = c(0.5, 1, 1.5, 2, 2.5), labels = c("0.5", "1", "1.5", "2", "2.5"), cex.axis = 0.7)
                  
                  text(c(1:month_n), comp_CPM.plt$C.CPM[1:month_n], labels = comp_CPM.plt$C.CPM[1:month_n], col = "cornsilk4", cex = 0.7, pos = 3)
                  points(comp_CPM.plt$Month.Plot, comp_CPM.plt$C.CPM, col = "cornsilk4", cex = 1, pch = 1)
                  abline(h = target_C, col = "red", lty = 1)
                  text(x = 2, y = target_C - 0.2, paste0("Target CPM = ", target_C), col = "red", cex = 0.7)
                  lines(comp_CPM.plt$Month.Plot, fitted(fit_C), col = "orange", lwd = 2, lty = 2)
                  text(x = 2, y = 2.16, paste0("R square: ", round(m_C$r.squared, 2)), col = "orange", cex = 0.7)
                  text(x = 2, y = 2.3, paste0("P-value: ", format.pval(round(pf(m_C$fstatistic[1], # F-statistic
                                                                                m_C$fstatistic[2], # df
                                                                                m_C$fstatistic[3], # df
                                                                                lower.tail = FALSE),2))), col = "orange", cex = 0.7)
                  text(x = 2, y = 2.41, paste0("Slope: ", round(m_C$coefficients[2],2)), col = "orange", cex = 0.7)
                  
    
  })     
  
  Bar_Line <- function()({if(is.null(CPM_monthly_NIC()))return(NULL)
                          data <- CPM_monthly_NIC()
                          n <- nrow(data)
                          pastyear2 <- format((input$YTD[1] - years(2)), "%Y")
                          pastyear1 <- format((input$YTD[1] - years(1)), "%Y")
                          currentyear <- format(input$YTD[1], "%Y")
                          data$Month <- seq(input$YTD[1]- years(2), by = "month", length.out = n)
                          data <- data[data$Month >= (input$YTD[1]- years(2)),]
                          
                          data0 <- data[data$Month >= input$YTD[1],]
                          if(nrow(data0)==12){data0 <- data0
                                        }else{m <- nrow(data0)
                                              data0[(m+1):12,] <- NA
                                              data0$Month <- seq(input$YTD[1], by = "month", length.out = 12)
                                              data0$complaint <- c(data0[1:m,]$complaint, rep(0, 12-m) )
                                              data0$CPM <- c(data0[1:m,]$CPM, rep(0, 12-m) )
                          }
                          data_bar <- rbind(data[data$Month < input$YTD[1],], data0)
                          data_bar$Month.Plot <- rep(1:12, 3)
                          data_bar$Year <- c(rep(pastyear2, 12), rep(pastyear1, 12), rep(currentyear, 12) )
                          data_bar$Month <- rep(c("Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"), 3)
                          
                          data_line <- data
                          data_line$Month.Plot <- rep(1:12, 3)[1:n]
                          data_line$Year <- c(rep(pastyear2, 12), rep(pastyear1, 12), rep(currentyear, 12) )[1:n]
                          data_line$Month <- rep(c("Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"), 3)[1:n]
                          
                          bar = ggplot(data = data_bar) +
                                ggtitle(paste0(pastyear2, " - ", currentyear, " NIC Complaints Received"))+
                                labs(x = "", y = "") +
                                theme(plot.title = element_text(size = 14, hjust = 0.5)) +
                                geom_bar(aes(x = Month.Plot, y = complaint, color = Year, fill = Year), 
                                         stat = "identity", position = "dodge", width = 0.5) +
                                scale_color_manual(values = c("2019" = "dodgerblue3", "2020" = "darkorange", "2021" = "darkolivegreen4")) +
                                scale_fill_manual(values = c("2019" = "dodgerblue3", "2020" = "darkorange", "2021" = "darkolivegreen4")) +
                                scale_x_discrete(limits = 1:12,
                                                 labels = c("1" = "Jan", "2" = "Feb", "3" = "Mar", "4" = "Apr", "5" = "May", "6" = "Jun",
                                                            "7" = "Jul", "8" = "Aug", "9" = "Sep", "10" = "Oct", "11" = "Nov", "12" = "Dec") ) +
                                scale_y_continuous(limits = c(0, 2000)) +     
                                theme(panel.background = element_rect(fill = "white", colour = NA),
                                      panel.border = element_blank(), 
                                      panel.grid.major = element_line(colour = "grey95"), 
                                      panel.grid.minor = element_line(colour = "grey95", size = 0.25)) 
                          
                          line = ggplot(data = data_line) +
                                 ggtitle(paste0(pastyear2, " - ", currentyear," NIC CPM") ) +
                                 labs(x = "", y = "") +
                                 geom_point(aes(x = Month.Plot, y = CPM, color = Year, fill = Year)) +
                                 geom_line(aes(x = Month.Plot, y = CPM, color = Year)) +
                                 geom_label_repel(aes(x = Month.Plot, y = CPM, group = Year, color = Year, label = CPM), 
                                                  size = 2.5, label.padding = 0.15, show.legend = F)+
                                 scale_x_discrete(limits = 1:12,
                                                  labels = c("1" = "Jan", "2" = "Feb", "3" = "Mar", "4" = "Apr", "5" = "May", "6" = "Jun",
                                                             "7" = "Jul", "8" = "Aug", "9" = "Sep", "10" = "Oct", "11" = "Nov", "12" = "Dec")) +
                                 scale_color_manual(values = c("2019" = "dodgerblue3", "2020" = "darkorange", "2021" = "darkolivegreen4")) +
                                 theme(plot.title = element_text(size = 14, hjust = 0.5),
                                       panel.background = element_rect(fill = "white", colour = NA),
                                       panel.border = element_blank(), 
                                       panel.grid.major = element_line(colour = "grey95"), 
                                       panel.grid.minor = element_line(colour = "grey95", size = 0.25)) 
                          
                          grid.newpage() 
                          grid.draw(rbind(ggplotGrob(line), ggplotGrob(bar), size = "last"))
    
  })
  
  Trend_Chart_GEM <- function()({if(is.null(CPM_region_monthly()))return(NULL)
                                 CPM_monthly <- CPM_region_monthly()
                                 target <- target_region()
                                 target <- as.numeric(target[which(target$Region == "GEM"),]$Target)
                                 month_n <- input$plot_range
                                 comp_CPM.plt <- CPM_monthly[(nrow(CPM_monthly) - (month_n - 1)):nrow(CPM_monthly),]
                                 comp_CPM.plt$Month.Plot <- c(1:nrow(comp_CPM.plt))
                                 fit_GEM <- lm(GEM.CPM ~ Month.Plot, data = comp_CPM.plt)
                                 m_GEM <- summary(fit_GEM)
    
                                 plot <- plot(comp_CPM.plt$Month.Plot, 
                                              comp_CPM.plt$GEM.CPM,
                                              type = "l",
                                              main = paste0("CPM of GEM for the Past ", month_n, " Months"),
                                              col.main = "dodgerblue4",
                                              ylim = c(25, 60),
                                              col = "dodgerblue4",
                                              lwd = 3,
                                              cex.main = 2,
                                              axes = F,
                                              ylab = expression(italic("CPM")),
                                              xlab = "",
                                              cex.lab = 2,
                                              yaxs = "i")
                                 #x-axis
                                  labellist <- comp_CPM.plt$Month
                                  axis(side = 1, at = seq(1, month_n, by = 1), labels = FALSE)
                                  text(seq(1, month_n, by = 1), par("usr")[3], srt = 45, label = labellist, pos = 1, xpd = TRUE, cex = 0.7)
    
                                 #y-axis
                                  axis(side = 2, at = c(25, 30, 35, 40, 45, 50, 55, 60), 
                                  labels = c("25", "30", "35", "40", "45", "50", "55", "60"), cex.axis = 0.7)
                                  abline(h = target, col = "red", lty = 1)
                                  text(x = 2, y = target - 2, paste0("Target CPM = ", target), col = "red")
                                  text(c(1:month_n), comp_CPM.plt$GEM.CPM[1:month_n], labels = comp_CPM.plt$GEM.CPM[1:month_n], col = "dodgerblue4", cex = 0.7, pos = 3)
                                  points(comp_CPM.plt$Month.Plot, comp_CPM.plt$GEM.CPM, col = "dodgerblue4", cex = 1, pch = 1)
                                  lines(comp_CPM.plt$Month.Plot, fitted(fit_GEM), col = "orange", lwd = 2, lty = 2)
                                  text(x = 2, y = 56.6, paste0("R square: ", round(m_GEM$r.squared, 2)), col = "orange")
                                  text(x = 2, y = 58, paste0("P-value: ", format.pval(round(pf(m_GEM$fstatistic[1], # F-statistic
                                                                                            m_GEM$fstatistic[2], # df
                                                                                            m_GEM$fstatistic[3], # df
                                                                                            lower.tail = FALSE),2))), col = "orange")
                                  text(x = 2, y = 59, paste0("Slope: ", round(m_GEM$coefficients[2],2)), col = "orange")
                                  return(plot)
  })
  
  Trend_Chart_EUR <- function()({if(is.null(CPM_region_monthly()))return(NULL)
                                 CPM_monthly <- CPM_region_monthly()
                                 target <- target_region()
                                 target <- as.numeric(target[which(target$Region == "EUR"),]$Target)
                                 month_n <- input$plot_range
                                 comp_CPM.plt <- CPM_monthly[(nrow(CPM_monthly) - (month_n - 1)):nrow(CPM_monthly),]
                                 comp_CPM.plt$Month.Plot <- c(1:nrow(comp_CPM.plt))
                                 fit_EUR <- lm(EUR.CPM ~ Month.Plot, data = comp_CPM.plt)
                                 m_EUR <- summary(fit_EUR)
    
                                 plot <- plot(comp_CPM.plt$Month.Plot, 
                                              comp_CPM.plt$EUR.CPM,
                                              type = "l",
                                              main = paste0("CPM of EUR for the Past ", month_n, " Months"),
                                              col.main = "darkorange1",
                                              ylim = c(2, 12),
                                              col = "darkorange1",
                                              lwd = 3,
                                              cex.main = 2,
                                              axes = F,
                                              ylab = expression(italic("CPM")),
                                              xlab = "",
                                              cex.lab = 2,
                                              yaxs = "i")
                                  #x-axis
                                  labellist <- comp_CPM.plt$Month
                                  axis(side = 1, at = seq(1, month_n, by = 1), labels = FALSE)
                                  text(seq(1, month_n, by = 1), par("usr")[3], srt = 45, label = labellist, pos = 1, xpd = TRUE, cex = 0.7)
    
                                  #y-axis
                                  axis(side = 2, at = c(2, 7, 12), 
                                  labels = c("2", "7", "12"), cex.axis = 0.7)
                                  abline(h = target, col = "red", lty = 1)
                                  text(x = 2, y = target - 1.8, paste0("Target CPM = ", target), col = "red")
                                  text(c(1:month_n), comp_CPM.plt$EUR.CPM[1:month_n], labels = comp_CPM.plt$EUR.CPM[1:month_n], col = "darkorange1", cex = 0.7, pos = 3)
                                  points(comp_CPM.plt$Month.Plot, comp_CPM.plt$EUR.CPM, col = "darkorange1", cex = 1, pch = 1)
                                  lines(comp_CPM.plt$Month.Plot, fitted(fit_EUR), col = "orange", lwd = 2, lty = 2)
                                  text(x = 2, y = 10, paste0("R square: ", round(m_EUR$r.squared, 2)), col = "orange")
                                  text(x = 2, y = 10.5, paste0("P-value: ", format.pval(round(pf(m_EUR$fstatistic[1], # F-statistic
                                                                                               m_EUR$fstatistic[2], # df
                                                                                               m_EUR$fstatistic[3], # df
                                                                                               lower.tail = FALSE),2))),col = "orange")
                                       
                                  text(x = 2, y = 11, paste0("Slope: ", round(m_EUR$coefficients[2],2)), col = "orange")
                                  return(plot)
  })
  
  Trend_Chart_NAM <- function()({if(is.null(CPM_region_monthly()))return(NULL)
                                 CPM_monthly <- CPM_region_monthly()
                                 target <- target_region()
                                 target <- as.numeric(target[which(target$Region == "NAM"),]$Target)
                                 month_n <- input$plot_range
                                 comp_CPM.plt <- CPM_monthly[(nrow(CPM_monthly) - (month_n - 1)):nrow(CPM_monthly),]
                                 comp_CPM.plt$Month.Plot <- c(1:nrow(comp_CPM.plt))
                                 fit_NAM <- lm(NAM.CPM ~ Month.Plot, data = comp_CPM.plt)
                                 m_NAM <- summary(fit_NAM)
    
                                 plot <- plot(comp_CPM.plt$Month.Plot, 
                                              comp_CPM.plt$NAM.CPM,
                                              type = "l",
                                 main = paste0("CPM of North America for the Past ", month_n, " Months"),
                                 col.main = "aquamarine4",
                                 ylim = c(60, 100),
                                 col = "aquamarine4",
                                 lwd = 3,
                                 cex.main = 2,
                                 axes = F,
                                 ylab = expression(italic("CPM")),
                                 xlab = "",
                                 cex.lab = 2,
                                 yaxs = "i")
                                 #x-axis
                                 labellist <- comp_CPM.plt$Month
                                 axis(side = 1, at = seq(1, month_n, by = 1), labels = FALSE)
                                 text(seq(1, month_n, by = 1), par("usr")[3], srt = 45, label = labellist, pos = 1, xpd = TRUE, cex = 0.7)
    
                                 #y-axis
                                  axis(side = 2, at = c(60, 70, 80, 90, 100), 
                                  labels = c("60", "70", "80", "90", "100"), cex.axis = 0.7)
                                  abline(h = target, col = "red", lty = 1)
                                  text(x = 2, y = target - 4, paste0("Target CPM = ", target), col = "red")
                                  text(c(1:month_n), comp_CPM.plt$NAM.CPM[1:month_n], labels = comp_CPM.plt$NAM.CPM[1:month_n], col = "aquamarine4", cex = 0.7, pos = 3)
                                  points(comp_CPM.plt$Month.Plot, comp_CPM.plt$NAM.CPM, col = "aquamarine4", cex = 1, pch = 1)
                                  lines(comp_CPM.plt$Month.Plot, fitted(fit_NAM), col = "orange", lwd = 2, lty = 2)
                                  text(x = 2, y = 85, paste0("R square: ", round(m_NAM$r.squared, 2)), col = "orange")
                                  text(x = 2, y = 86.5, paste0("P-value: ", format.pval(round(pf(m_NAM$fstatistic[1], # F-statistic
                                                                                               m_NAM$fstatistic[2], # df
                                                                                               m_NAM$fstatistic[3], # df
                                                                                               lower.tail = FALSE),2))),
                                       col = "orange")
                                  text(x = 2, y = 88, paste0("Slope: ", round(m_NAM$coefficients[2],2)), col = "orange")
                                  return(plot)
  })
  
  Trend_Chart_AP <- function()({if(is.null(CPM_region_monthly2()))return(NULL)
                                CPM_monthly <- CPM_region_monthly2()
                                target <- target_region()
                                target <- as.numeric(target[which(target$Region == "APAC"),]$Target)
                                month_n <- input$plot_range
                                comp_CPM.plt <- CPM_monthly[(nrow(CPM_monthly) - (month_n - 1)):nrow(CPM_monthly),]
                                comp_CPM.plt$Month.Plot <- c(1:nrow(comp_CPM.plt))
                                fit_AP <- lm(AP.CPM ~ Month.Plot, data = comp_CPM.plt)
                                m_AP <- summary(fit_AP)
    
                                plot <- plot(comp_CPM.plt$Month.Plot, 
                                             comp_CPM.plt$AP.CPM,
                                             type = "l",
                                             main = paste0("CPM of APAC for the Past ", month_n, " Months"),
                                             col.main = "cornflowerblue",
                                             ylim = c(40, 90),
                                             col = "cornflowerblue",
                                             lwd = 3,
                                             cex.main = 2,
                                             axes = F,
                                             ylab = expression(italic("CPM")),
                                             xlab = "",
                                             cex.lab = 2,
                                             yaxs = "i")
                                             #x-axis
                                             labellist <- comp_CPM.plt$Month
                                             axis(side = 1, at = seq(1, month_n, by = 1), labels = FALSE)
                                             text(seq(1, month_n, by = 1), par("usr")[3], srt = 45, label = labellist, pos = 1, xpd = TRUE, cex = 0.7)
    
                                             #y-axis
                                             axis(side = 2, at = c(40, 50, 60, 70, 80, 90), 
                                             labels = c("40", "50", "60", "70", "80", "90"), cex.axis = 0.7)
                                             abline(h = target, col = "red", lty = 1)
                                             text(x = 2, y = target - 2, paste0("Target CPM = ", target), col = "red")
                                             text(c(1:month_n), comp_CPM.plt$AP.CPM[1:month_n], labels = comp_CPM.plt$AP.CPM[1:month_n], col = "cornflowerblue", cex = 0.7, pos = 3)
                                             points(comp_CPM.plt$Month.Plot, comp_CPM.plt$AP.CPM, col = "cornflowerblue", cex = 1, pch = 1)
                                             lines(comp_CPM.plt$Month.Plot, fitted(fit_AP), col = "orange", lwd = 2, lty = 2)
                                             text(x = 2, y = 76.6, paste0("R square: ", round(m_AP$r.squared, 2)), col = "orange")
                                             text(x = 2, y = 78, paste0("P-value: ", format.pval(round(pf(m_AP$fstatistic[1], # F-statistic
                                                                                                          m_AP$fstatistic[2], # df
                                                                                                          m_AP$fstatistic[3], # df
                                                                                                          lower.tail = FALSE),2))),
                                                  col = "orange")
                                             text(x = 2, y = 79, paste0("Slope: ", round(m_AP$coefficients[2],2)), col = "orange")
                                             return(plot)
  })
  
  Trend_Chart_LA <- function()({if(is.null(CPM_region_monthly2()))return(NULL)
                                CPM_monthly <- CPM_region_monthly2()
                                target <- target_region()
                                target <- as.numeric(target[which(target$Region == "LATAM"),]$Target)
                                month_n <- input$plot_range
                                comp_CPM.plt <- CPM_monthly[(nrow(CPM_monthly) - (month_n - 1)):nrow(CPM_monthly),]
                                comp_CPM.plt$Month.Plot <- c(1:nrow(comp_CPM.plt))
                                fit_LA <- lm(LA.CPM ~ Month.Plot, data = comp_CPM.plt)
                                m_LA <- summary(fit_LA)
    
                                plot <- plot(comp_CPM.plt$Month.Plot, 
                                             comp_CPM.plt$LA.CPM,
                                             type = "l",
                                             main = paste0("CPM of LATAM for the Past ", month_n, " Months"),
                                             col.main = "cornflowerblue",
                                             ylim = c(0, 20),
                                             col = "cornflowerblue",
                                             lwd = 3,
                                             cex.main = 2,
                                             axes = F,
                                             ylab = expression(italic("CPM")),
                                             xlab = "",
                                             cex.lab = 2,
                                             yaxs = "i")
                                #x-axis
                                labellist <- comp_CPM.plt$Month
                                axis(side = 1, at = seq(1, month_n, by = 1), labels = FALSE)
                                text(seq(1, month_n, by = 1), par("usr")[3], srt = 45, label = labellist, pos = 1, xpd = TRUE, cex = 0.7)
    
                                #y-axis
                                axis(side = 2, at = c(0, 5, 10, 15, 20), 
                                    labels = c("0", "5", "10", "15", "20"), cex.axis = 0.7)
                                abline(h = target, col = "red", lty = 1)
                                text(x = 2, y = target - 1, paste0("Target CPM = ", target), col = "red")
                                text(c(1:month_n), comp_CPM.plt$LA.CPM[1:month_n], labels = comp_CPM.plt$LA.CPM[1:month_n], col = "cornflowerblue", cex = 0.7, pos = 3)
                                points(comp_CPM.plt$Month.Plot, comp_CPM.plt$LA.CPM, col = "cornflowerblue", cex = 1, pch = 1)
                                lines(comp_CPM.plt$Month.Plot, fitted(fit_LA), col = "orange", lwd = 2, lty = 2)
                                text(x = 2, y = 17.7, paste0("R square: ", round(m_LA$r.squared, 2)), col = "orange")
                                text(x = 2, y = 18.5, paste0("P-value: ", format.pval(round(pf(m_LA$fstatistic[1], # F-statistic
                                                                                               m_LA$fstatistic[2], # df
                                                                                               m_LA$fstatistic[3], # df
                                                                                               lower.tail = FALSE),2))),
                                                                                      col = "orange")
                                text(x = 2, y = 19, paste0("Slope: ", round(m_LA$coefficients[2],2)), col = "orange")
                                return(plot)
  })
  
  Trend_Chart_MEA <- function()({if(is.null(CPM_region_monthly2()))return(NULL)
                                 CPM_monthly <- CPM_region_monthly2()
                                 target <- target_region()
                                 target <- as.numeric(target[which(target$Region == "MEA"),]$Target)
                                 month_n <- input$plot_range
                                 comp_CPM.plt <- CPM_monthly[(nrow(CPM_monthly) - (month_n - 1)):nrow(CPM_monthly),]
                                 comp_CPM.plt$Month.Plot <- c(1:nrow(comp_CPM.plt))
                                 fit_MEA <- lm(MEA.CPM ~ Month.Plot, data = comp_CPM.plt)
                                 m_MEA <- summary(fit_MEA)
    
                                 plot <- plot(comp_CPM.plt$Month.Plot, 
                                              comp_CPM.plt$MEA.CPM,
                                              type = "l",
                                              main = paste0("CPM of MEA for the Past ", month_n, " Months"),
                                              col.main = "cornflowerblue",
                                              ylim = c(0, 8),
                                              col = "cornflowerblue",
                                              lwd = 3,
                                              cex.main = 2,
                                              axes = F,
                                              ylab = expression(italic("CPM")),
                                              xlab = "",
                                              cex.lab = 2,
                                              yaxs = "i")
                                  #x-axis
                                   labellist <- comp_CPM.plt$Month
                                   axis(side = 1, at = seq(1, month_n, by = 1), labels = FALSE)
                                   text(seq(1, month_n, by = 1), par("usr")[3], srt = 45, label = labellist, pos = 1, xpd = TRUE, cex = 0.7)
    
                                  #y-axis
                                   axis(side = 2, at = c(0, 2, 4, 6, 8), 
                                   labels = c("0", "2", "4", "6", "8"), cex.axis = 0.7)
                                   abline(h = target, col = "red", lty = 1)
                                   text(x = 2, y = target - 0.3, paste0("Target CPM = ", target), col = "red")
                                   text(c(1:month_n), comp_CPM.plt$MEA.CPM[1:month_n], labels = comp_CPM.plt$MEA.CPM[1:month_n], col = "cornflowerblue", cex = 0.7, pos = 3)
                                   points(comp_CPM.plt$Month.Plot, comp_CPM.plt$MEA.CPM, col = "cornflowerblue", cex = 1, pch = 1)
                                   lines(comp_CPM.plt$Month.Plot, fitted(fit_MEA), col = "orange", lwd = 2, lty = 2)
                                   text(x = 2, y = 6.8, paste0("R square: ", round(m_MEA$r.squared, 2)), col = "orange")
                                   text(x = 2, y = 7.2, paste0("P-value: ", format.pval(round(pf(m_MEA$fstatistic[1], # F-statistic
                                                                                                 m_MEA$fstatistic[2], # df
                                                                                                 m_MEA$fstatistic[3], # df
                                                                                                 lower.tail = FALSE),2))),
                                                                                                 col = "orange")
                                   text(x = 2, y = 7.5, paste0("Slope: ", round(m_MEA$coefficients[2],2)), col = "orange")
                                   return(plot)
  })
  

  most_recent_m <- reactive({if(is.null(CPM_monthly())|is.null(CPM_monthly_NIC)|is.null(CPM_monthly_all))return(NULL)
                             CPM_monthly <- CPM_monthly()
                             CPM_monthly_NIC <- CPM_monthly_NIC()
                             CPM_monthly_all <- CPM_monthly_all()
                             rows <- (nrow(CPM_monthly)-11):nrow(CPM_monthly)
                             most_recent_m <- data.frame("Month" = as.character(CPM_monthly[rows,]$Month),
                                                         "All BU" = CPM_monthly_all[rows,]$CPM,
                                                         "Non-IC" = CPM_monthly_NIC[rows,]$CPM,
                                                         "IC" = CPM_monthly[rows,]$IC.CPM,
                                                         "AWC" = CPM_monthly[rows,]$AWC.CPM,
                                                         "CC" = CPM_monthly[rows,]$CC.CPM,
                                                         "C" = CPM_monthly[rows,]$C.CPM,
                                                         "OC" = CPM_monthly[rows,]$OC.CPM
                                             )
                             return(most_recent_m) 
  })
  
  most_recent_m_region <- reactive({if(is.null(CPM_region_monthly()))return(NULL)
                                    CPM_monthly <- CPM_region_monthly()
                                    rows <- (nrow(CPM_monthly)-11):nrow(CPM_monthly)
                                    most_recent_m <- data.frame("Month" = as.character(CPM_monthly[rows,]$Month),
                                                                "GEM" = CPM_monthly[rows,]$GEM.CPM,
                                                                "EUR" = CPM_monthly[rows,]$EUR.CPM,
                                                                "NAM" = CPM_monthly[rows,]$NAM.CPM
                                                                 )
                                     return(most_recent_m) 
  })
  
  most_recent_m_region2 <- reactive({if(is.null(CPM_region_monthly()))return(NULL)
                                     CPM_monthly2 <- CPM_region_monthly2()
                                     rows <- (nrow(CPM_monthly2)-11):nrow(CPM_monthly2)
                                     most_recent_m <- data.frame("Month" = as.character(CPM_monthly2[rows,]$Month),
                                                                 "APAC" = CPM_monthly2[rows,]$AP.CPM,
                                                                 "LATAM" = CPM_monthly2[rows,]$LA.CPM,
                                                                 "MEA" = CPM_monthly2[rows,]$MEA.CPM
                                     )
                                     return(most_recent_m) 
  })
  
  most_recent_y <- reactive({if(is.null(CPM_yearly())|is.null(CPM_yearly_NIC)|is.null(CPM_yearly_all))return(NULL)
                             CPM_yearly <- CPM_yearly()
                             CPM_yearly_NIC <- CPM_yearly_NIC()
                             CPM_yearly_all <- CPM_yearly_all()
                             #pastyear2 <- format((input$YTD[1] - years(2)), "%Y")
                             pastyear1 <- format((input$YTD[1] - years(1)), "%Y")
                             currentyear <- format(input$YTD[1], "%Y")
                             
                             most_recent_y <- data.frame("YTD" = CPM_yearly[which(CPM_yearly$Year==paste0(format((input$YTD[1]), "%Y")," YTD")),]$Year,
                                                         "All BU" = CPM_yearly_all[which(CPM_yearly_all$Year==paste0(format((input$YTD[1]), "%Y")," YTD")),]$CPM,
                                                         "NIC" = CPM_yearly_NIC[which(CPM_yearly_NIC$Year==paste0(format((input$YTD[1]), "%Y")," YTD")),]$CPM,
                                                         "IC" = CPM_yearly[which(CPM_yearly$Year==paste0(format((input$YTD[1]), "%Y")," YTD")),]$IC.CPM,
                                                         "AWC" = CPM_yearly[which(CPM_yearly$Year==paste0(format((input$YTD[1]), "%Y")," YTD")),]$AWC.CPM,
                                                         "CC" = CPM_yearly[which(CPM_yearly$Year==paste0(format((input$YTD[1]), "%Y")," YTD")),]$CC.CPM,
                                                         "C" = CPM_yearly[which(CPM_yearly$Year==paste0(format((input$YTD[1]), "%Y")," YTD")),]$C.CPM,
                                                         "OC" = CPM_yearly[which(CPM_yearly$Year==paste0(format((input$YTD[1]), "%Y")," YTD")),]$OC.CPM
                             )
                             return(most_recent_y) 
  })
  
  most_recent_y_region <- reactive({if(is.null(CPM_region_yearly()))return(NULL)
                                    CPM_yearly <- CPM_region_yearly()
                                    pastyear1 <- format((input$YTD[1] - years(1)), "%Y")
                                    currentyear <- format(input$YTD[1], "%Y")
    
                                    most_recent_y <- data.frame("YTD" = CPM_yearly[which(CPM_yearly$Year==paste0(format((input$YTD[1]), "%Y")," YTD")),]$Year,
                                                                "GEM" = CPM_yearly[which(CPM_yearly$Year==paste0(format((input$YTD[1]), "%Y")," YTD")),]$GEM.CPM,
                                                                "EUR" = CPM_yearly[which(CPM_yearly$Year==paste0(format((input$YTD[1]), "%Y")," YTD")),]$EUR.CPM,
                                                                "NAM" = CPM_yearly[which(CPM_yearly$Year==paste0(format((input$YTD[1]), "%Y")," YTD")),]$NAM.CPM
                                                                 )
                                     return(most_recent_y) 
  })
  
  most_recent_y_region2 <- reactive({if(is.null(CPM_region_yearly2()))return(NULL)
                                     CPM_yearly2 <- CPM_region_yearly2()
                                     pastyear1 <- format((input$YTD[1] - years(1)), "%Y")
                                     currentyear <- format(input$YTD[1], "%Y")
    
                                     most_recent_y <- data.frame("YTD" = CPM_yearly2[which(CPM_yearly2$Year==paste0(format((input$YTD[1]), "%Y")," YTD")),]$Year,
                                                                 "APAC" = CPM_yearly2[which(CPM_yearly2$Year==paste0(format((input$YTD[1]), "%Y")," YTD")),]$AP.CPM,
                                                                 "LATAM" = CPM_yearly2[which(CPM_yearly2$Year==paste0(format((input$YTD[1]), "%Y")," YTD")),]$LA.CPM,
                                                                 "MEA" = CPM_yearly2[which(CPM_yearly2$Year==paste0(format((input$YTD[1]), "%Y")," YTD")),]$MEA.CPM
                                                                 )
                                 return(most_recent_y) 
  })
  
  YTD_html <- reactive({if(is.null(most_recent_y()))return(NULL)
                        most_recent_y <- most_recent_y()
                        tb <- paste0('<table class = "tf" id = "t01">
                                      <tr>
                                      <th> Business Unit </th>
                                      <td> All BU </td>
                                      <td> Non-IC </td>
                                      <td> IC </td>
                                      <td> AWC </td>
                                      <td> CC </td>
                                      <td> C </td>
                                      <td> OC </td>
                                      </tr>
                                      <tr>
                                      <th>', most_recent_y$"YTD", '</th> 
                                      <td>', most_recent_y[,2], '</td>
                                      <td>', most_recent_y$"NIC", ' </td>
                                      <td>', most_recent_y$"IC", ' </td>
                                      <td>', most_recent_y$"AWC", ' </td>
                                      <td>', most_recent_y$"CC", ' </td>
                                      <td>', most_recent_y$"C", ' </td>
                                      <td>', most_recent_y$"OC", ' </td>
                                      </tr>
                                      </table>')
    return(tb)
  })
  
  YTD_region_html <- reactive({if(is.null(most_recent_y_region()))return(NULL)
                               most_recent_y <- most_recent_y_region()
                               tb <- paste0('<table class = "tf" id = "t01">
                                             <tr>
                                             <th> Region </th>
                                             <td> GEM </td>
                                             <td> EUR </td>
                                             <td> North America </td>
                                             </tr>
                                             <tr>
                                             <th>', most_recent_y$"YTD", '</th>
                                             <td>', most_recent_y$"GEM", ' </td>
                                             <td>', most_recent_y$"EUR", ' </td>
                                             <td>', most_recent_y$"NAM", ' </td>\
                                             </tr>
                                             </table>')
                                             return(tb)
  })
  
  YTD_region_html2 <- reactive({if(is.null(most_recent_y_region2()))return(NULL)
                               most_recent_y <- most_recent_y_region2()
                               tb <- paste0('<table class = "tf" id = "t01">
                                             <tr>
                                             <th> Region </th>
                                             <td> APAC </td>
                                             <td> LATAM </td>
                                             <td> MEA </td>
                                             </tr>
                                             <tr>
                                             <th>', most_recent_y$"YTD", '</th>
                                             <td>', most_recent_y$"APAC", ' </td>
                                             <td>', most_recent_y$"LATAM", ' </td>
                                             <td>', most_recent_y$"MEA", ' </td>
                                             </tr>
                                             </table>')
                                              return(tb)
  })
  
  
  output$target_ui <- renderUI({
                                HTML(sprintf(
                                            '<!DOCTYPE html>
                                             <html>

                                             <style>
                                             table {width:70%%;}
                                             table, th, td {border: 1px solid black; border-collapse: collapse;}
                                             th, td {padding: 12px; text-align: left;}
                                             table#t01 tr:nth-child(even) {background-color: #fff;}
                                             table#t01 tr:nth-child(odd) {background-color: #eee;}
                                             table#t01 th {background-color: grey; color: white;}
                                             </style>
    
                                             <body>
                                             <h3> Target CPM for Current Year </h3>
                                             %s
                                             </body>
                                             </html>', 
                                             target_html())
                                              )
  })
  
  output$target_ui_r <- renderUI({
                            HTML(sprintf(
                                         '<!DOCTYPE html>
                                          <html>
      
                                          <style>
                                          table {width:70%%;}
                                          table, th, td {border: 1px solid black; border-collapse: collapse;}
                                          th, td {padding: 12px; text-align: left;}
                                          table#t01 tr:nth-child(even) {background-color: #fff;}
                                          table#t01 tr:nth-child(odd) {background-color: #eee;}
                                          table#t01 th {background-color: grey; color: white;}
                                          </style>
      
                                          <body>
                                          <h3> Target CPM for Current Year </h3>
                                          %s
                                          </body>
                                          </html>', 
                                          target_region_html())
    )
  })
  
  output$target_ui_r2 <- renderUI({
                              HTML(sprintf(
                                           '<!DOCTYPE html>
                                            <html>
      
                                            <style>
                                            table {width:70%%;}
                                            table, th, td {border: 1px solid black; border-collapse: collapse;}
                                            th, td {padding: 12px; text-align: left;}
                                            table#t01 tr:nth-child(even) {background-color: #fff;}
                                            table#t01 tr:nth-child(odd) {background-color: #eee;}
                                            table#t01 th {background-color: grey; color: white;}
                                            </style>
      
                                            <body>
                                            <h3> Target CPM for Current Year </h3>
                                              %s
                                            </body>
                                            </html>', 
                                             target_region_html2())
    )
  })
  
  output$YTD_ui <- renderUI({
                             HTML(sprintf(
                                         '<!DOCTYPE html>
                                          <html>
      
                                          <style>
                                          table {width:50%%;}
                                          table, th, td {border: 1px solid black; border-collapse: collapse;}
                                          th, td {padding: 12px; text-align: left;}
                                          table#t01 tr:nth-child(even) {background-color: #fff;}
                                          table#t01 tr:nth-child(odd) {background-color: #eee;}
                                          table#t01 th {background-color: grey; color: white;}
                                          </style>
      
                                          <body>
                                          <h3> YTD CPM by Business Unit </h3>
                                           %s
                                          </body>
                                          </html>', 
                                          YTD_html())
                                   )
  })
  
  output$YTD_ui_r <- renderUI({HTML(sprintf(
                                            '<!DOCTYPE html>
                                             <html>
      
                                             <style>
                                             table {width:50%%;}
                                             table, th, td {border: 1px solid black; border-collapse: collapse;}
                                             th, td {padding: 12px; text-align: left;}
                                             table#t01 tr:nth-child(even) {background-color: #fff;}
                                             table#t01 tr:nth-child(odd) {background-color: #eee;}
                                             table#t01 th {background-color: grey; color: white;}
                                             </style>
      
                                             <body>
                                             <h3> YTD CPM by Region </h3>
                                              %s
                                             </body>
                                             </html>', 
                                             YTD_region_html())
                                               )
  })
  
  output$YTD_ui_r2 <- renderUI({HTML(sprintf(
                                             '<!DOCTYPE html>
                                              <html>
    
                                              <style>
                                              table {width:50%%;}
                                              table, th, td {border: 1px solid black; border-collapse: collapse;}
                                              th, td {padding: 12px; text-align: left;}
                                              table#t01 tr:nth-child(even) {background-color: #fff;}
                                              table#t01 tr:nth-child(odd) {background-color: #eee;}
                                              table#t01 th {background-color: grey; color: white;}
                                              </style>
    
                                              <body>
                                              <h3> YTD CPM by Region </h3>
                                               %s
                                              </body>
                                              </html>', 
                                              YTD_region_html2())
                                               )
  })
  

  
  output$plot <- renderPlot(switch(input$plot_option,
                                   NIC = Trend_Chart_NIC(),
                                   all_com = Trend_Chart_all_com(),
                                   all = Trend_Chart_all(), # all BU
                                   all_bu = Trend_Chart_ALL(), #all BU combined
                                   IC = Trend_Chart_IC(),
                                   OC = Trend_Chart_OC(),
                                   CC = Trend_Chart_CC(),
                                   C = Trend_Chart_C(),
                                   AWC = Trend_Chart_AWC(),
                                   Eurotec = Trend_Chart_Eurotec(),
                                   Bar_Line = Bar_Line()),
                            height = 745, width = 970, res = 105)
  
  output$plot_r <- renderPlot(switch(input$plot_option_r,
                                     GEM = Trend_Chart_GEM(),
                                     EUR = Trend_Chart_EUR(),
                                     NAM = Trend_Chart_NAM(),
                                     AP = Trend_Chart_AP(),
                                     LA = Trend_Chart_LA(),
                                     MEA = Trend_Chart_MEA()),
                              height = 745, width = 970, res = 105)
  

  output$Monthly_tb <- renderTable({most_recent_m()})
  output$Monthly_tb_r <- renderTable({most_recent_m_region()})
  output$Monthly_tb_r2 <- renderTable({most_recent_m_region2()})
  output$slides <- downloadHandler(filename = function() {paste(format(input$YTD[2],"%Y"), "_", format(input$YTD[2],"%B"), "_CPM_Digest", ".pptx", sep = "")},
                                   content = function(file) {
                                     
                                   titlepageImage <- paste(input$images$datapath[input$images$name=="title_page.PNG"])
                                   leftCornerImage <- paste(input$images$datapath[input$images$name=="left_corner.PNG"])
                                   logoImage <- paste(input$images$datapath[input$images$name=="logo.PNG"])
                                   
                                   OC <- OC()
                                   AWC <- AWC()
                                   CC <- CC()
                                   C <- C()
                                   IC <- IC()
                                   Eurotec <- Eurotec()
                                   NIC <- NIC()
                                   ALL <- ALL()
                                   
                                   BSC <- BSC()
                                   BSC_region <- BSC_region()
                                   GMSC <- GMSC()
                                   CPM_COM_NIC <- CPM_COM_NIC()
                                   target <- target()
                                   target.IC <- as.numeric(target[which(target$BU == "IC"),]$Target)
                                   target.OC <- as.numeric(target[which(target$BU == "OC"),]$Target)
                                   target.AWC <- as.numeric(target[which(target$BU == "AWC"),]$Target)
                                   target.CC <- as.numeric(target[which(target$BU == "CC"),]$Target)
                                   target.C <- as.numeric(target[which(target$BU == "C"),]$Target)
                                   target.Eurotec <- as.numeric(target[which(target$BU == "(Eurotec)"),]$Target)
                                   
                                   std_border = fp_border(color="black") 
                                   pastyear2 <- format((input$YTD[1] - years(2)), "%Y")
                                   pastyear1 <- format((input$YTD[1] - years(1)), "%Y")
                                   currentyear <- format(input$YTD[1], "%Y")
                                   
                                   GEM <- GEM()
                                   EUR <- EUR()
                                   NAM <- NAM()
                                   AP <- AP()
                                   LA <- LA()
                                   MEA <- MEA()
                                   
                                   
                                   title_BU <- block_list(
                                     fpar(ftext("CPM Digest", shortcuts$fp_bold(font.size = 36, color = "white")))
                                   )
                                
                                   title2 <- block_list(
                                     fpar(ftext(paste("Data Coverage: through", input$YTD[2]), shortcuts$fp_bold(font.size = 18, color = "white"))),
                                     fpar(ftext(paste("Report Date:", input$report), shortcuts$fp_bold(font.size = 18, color = "white"))),
                                     fpar(ftext(paste(input$xx, "on behalf of Metrics Central"), shortcuts$fp_bold(font.size = 18, color = "white")))
                                   )
                                   
                                   caption_GMSC <- block_list(
                                     fpar(ftext("2021 CPM Reports per Business Unit", shortcuts$fp_bold(font.size = 22, color = "darkred")))
                                   )
                                   
                                   caption_BSC <- block_list(
                                     fpar(ftext("2021 CPM - Balanced Score Card", shortcuts$fp_bold(font.size = 22, color = "darkred")))
                                   )
                                   caption_BSC_region <- block_list(
                                     fpar(ftext("2021 CPM (Region) - Balanced Score Card", shortcuts$fp_bold(font.size = 22, color = "darkred")))
                                   )
                                   
                                   caption_OC <- block_list(
                                     fpar(ftext("CPM - Ostomy Care (OC)", shortcuts$fp_bold(font.size = 22, color = "aquamarine4")))
                                   )
                                   
                                   caption_WC <- block_list(
                                     fpar(ftext("CPM - Wound Care (WC)", shortcuts$fp_bold(font.size = 22, color = "dodgerblue4")))
                                   )
                                   
                                   caption_CC <- block_list(
                                     fpar(ftext("CPM - Continence Care (CC)", shortcuts$fp_bold(font.size = 22, color = "darkgoldenrod1")))
                                   )
                                   
                                   caption_C <- block_list(
                                     fpar(ftext("CPM - Critical Care (C)", shortcuts$fp_bold(font.size = 22, color = "cornsilk4")))
                                   )
                                   
                                   
                                   caption_IC <- block_list(
                                     fpar(ftext("CPM - Infusion Care (IC)", shortcuts$fp_bold(font.size = 22, color = "darkorange1")))
                                   )
                                   
                                   caption_Eurotec <- block_list(
                                     fpar(ftext("CPM - (Eurotec)", shortcuts$fp_bold(font.size = 22, color = "aquamarine3")))
                                   )
                                   
                                   caption_NIC <- block_list(
                                     fpar(ftext("CPM - non-Infusion Care BU combined (Non-IC)", shortcuts$fp_bold(font.size = 22, color = "darkcyan")))
                                   )
                                   
                                   caption_all <- block_list(
                                     fpar(ftext("CPM - All Business Units", shortcuts$fp_bold(font.size = 22, color = "coral4")))
                                   )
                                   
                                   caption_combined <- block_list(
                                     fpar(ftext("CPM by Business Unit", shortcuts$fp_bold(font.size = 22, color = "black")))
                                   )
                                   
                                   note_BU <- block_list(
                                     fpar(ftext(paste0("*Target: Assuming ", (1-as.numeric(reduction_ALL()))*100, "% complaints reduction from previous year."), shortcuts$fp_italic(color="black"))) 
                                   )
                                   
                                   note_NIC <- block_list(
                                     fpar(ftext(paste0("*Target: Assuming ", (1-as.numeric(reduction_NIC()))*100, "% complaints reduction from previous year."), shortcuts$fp_italic(color="black"))) 
                                   )
                                   
                                   note_IC <- block_list(
                                     fpar(ftext(paste0("*Target: Assuming ", (1-as.numeric(input$reduction_IC))*100, "% complaints reduction from previous year."), shortcuts$fp_italic(color="black"))) 
                                   )
                                   
                                   note_Eurotec <- block_list(
                                     fpar(ftext(paste0("*Target: Assuming ", (1-as.numeric(input$reduction_Eurotec))*100, "% complaints reduction from previous year."), shortcuts$fp_italic(color="black"))) 
                                   )
                                   
                                   note_AWC <- block_list(
                                     fpar(ftext(paste0("*Target: Assuming ", (1-as.numeric(input$reduction_AWC))*100, "% complaints reduction from previous year."), shortcuts$fp_italic(color="black"))) 
                                   )
                                   
                                   note_CC <- block_list(
                                     fpar(ftext(paste0("*Target: Assuming ", (1-as.numeric(input$reduction_CC))*100, "% complaints reduction from previous year."), shortcuts$fp_italic(color="black"))) 
                                   )
                                   
                                   note_C <- block_list(
                                     fpar(ftext(paste0("*Target: Assuming ", (1-as.numeric(input$reduction_C))*100, "% complaints reduction from previous year."), shortcuts$fp_italic(color="black"))) 
                                   )
                                  
                                   note_OC <- block_list(
                                     fpar(ftext(paste0("*Target: Assuming ", (1-as.numeric(input$reduction_OC))*100, "% complaints reduction from previous year."), shortcuts$fp_italic(color="black"))) 
                                   )
                                   
                                   note_SKC <- block_list(
                                     fpar(ftext(paste0("*Complaints for SKC (Skin Care) have been regrouped to under corresponding business units for OC and WC."), shortcuts$fp_italic(color="red"))) 
                                   )
                                   
                                   note_GMSC1 <- block_list(
                                     fpar(ftext(paste0("*Currently targeted on 10% complaints reduction to ALL business unites; subject to change."), 
                                                shortcuts$fp_italic(font.size = 9, color="red"))) 
                                   )
                                   
                                   note_GMSC2 <- block_list(
                                     fpar(ftext(paste0("*Performance for previous CC&C business unit breaks down to Continence and Critical Care correspondingly."), 
                                                shortcuts$fp_italic(font.size = 9, color="red"))) 
                                   )
                                   
                                   note_GMSC3 <- block_list(
                                     fpar(ftext(paste0("*Complaints and sales eaches for previous Skin Care business unit are regrouped back to OC and WC respectively based on TW Part # and product ICC code ."), 
                                                shortcuts$fp_italic(font.size = 9, color="red"))) 
                                   )
                                   
                                   note_GMSC4 <- block_list(
                                     fpar(ftext(paste0("*Both Eurotec Complaints and Sales are included in the Ostomy CPM calculation."), 
                                                shortcuts$fp_italic(font.size = 9, color="red"))) 
                                   )
                                   
                                   note_BSC1 <- block_list(
                                     fpar(ftext(paste0("*Currently targeted on 10% complaints reduction to ALL business unites; subject to change by GQOR/QARA LT."), 
                                                shortcuts$fp_italic(font.size = 9, color="red"))) 
                                   )
                                   
                                   note_BSC2 <- block_list(
                                     fpar(ftext(paste0("*CPM calculation by region only reflects performance from non-Infusion Care business units."), 
                                                shortcuts$fp_italic(font.size = 9, color="red"))) 
                                   )
                                   
                                   note_BSC3 <- block_list(
                                     fpar(ftext(paste0("*GEM: Global Emerging Market includes country performances from APAC, LATAM, MEA and TURKEY."), 
                                                shortcuts$fp_italic(font.size = 9, color="black"))) 
                                   )
                                   
                                   note_BSC4 <- block_list(
                                     fpar(ftext(paste0("*NA: North American includes country performances from USA and CANADA."), 
                                                shortcuts$fp_italic(font.size = 9, color="black"))) 
                                   )
                                   
                                   note_BSC5 <- block_list(
                                     fpar(ftext(paste0("*EUR: European includes all European country excluding TURKEY."), 
                                                shortcuts$fp_italic(font.size = 9, color="black"))) 
                                   )
                                   
                                   note_BSC6 <- block_list(
                                     fpar(ftext(paste0("*Exclude data from Infusion Care, SKC-Medline, Co-barnding, and Eurotec."), 
                                                shortcuts$fp_italic(font.size = 9, color="red"))) 
                                   )
                                  
                                   target <- block_list(
                                     fpar(ftext(paste0("Reduction Target:", " IC - ", target.IC, "  OC - ", target.OC, "  WC - ", target.AWC, "  CC - ", target.CC, "  C - ", target.C), 
                                                shortcuts$fp_italic(color="black"))) 
                                   )
                                   
                                   
                                   title_Region <- block_list(
                                     fpar(ftext("CPM by Region", shortcuts$fp_bold(font.size = 36, color = "white")))
                                   )
                                   
                                   caption_GEM <- block_list(
                                     fpar(ftext("CPM - GEM", shortcuts$fp_bold(font.size = 22, color = "dodgerblue4")))
                                   )
                                   
                                   caption_EUR <- block_list(
                                     fpar(ftext("CPM - EUR", shortcuts$fp_bold(font.size = 22, color = "darkorange1")))
                                   )
                                   
                                   caption_NAM <- block_list(
                                     fpar(ftext("CPM - NAM", shortcuts$fp_bold(font.size = 22, color = "aquamarine4")))
                                   )
                                   
                                   caption_AP <- block_list(
                                     fpar(ftext("CPM - GEM Breakdown (APAC)", shortcuts$fp_bold(font.size = 22, color = "cornflowerblue")))
                                   )
                                   
                                   caption_LA <- block_list(
                                     fpar(ftext("CPM - GEM Breakdown (LATAM)", shortcuts$fp_bold(font.size = 22, color = "cornflowerblue")))
                                   )
                                   
                                   caption_MEA <- block_list(
                                     fpar(ftext("CPM - GEM Breakdown (MEA)", shortcuts$fp_bold(font.size = 22, color = "cornflowerblue")))
                                   )
                                   
                                   note_Region <- block_list(
                                     fpar(ftext(paste0("*Note: Exclude data from Infusion Care, SKC-Medline, Co-barnding, and Eurotec."), shortcuts$fp_italic(font.size = 10, color="red"))) 
                                   )
                                   
                                   note_GEM <- block_list(
                                     fpar(ftext(paste0("GEM - Global Emerging Market (APAC+LATAM+MEA+TURKEY)"), shortcuts$fp_italic(font.size = 10, color="black"))) 
                                   )
                                   
                                   note_EUR <- block_list(
                                     fpar(ftext(paste0("EUR: European Countries (excluding Turkey)"), shortcuts$fp_italic(font.size = 10, color="black"))) 
                                   )
                                   
                                   note_NAM <- block_list(
                                     fpar(ftext(paste0("NAM: North America (US +PUERTO RICO+ CANADA)"), shortcuts$fp_italic(font.size = 10, color="black"))) 
                                   )
                                   
                                   reduction_GEM <- block_list(
                                     fpar(ftext(paste0("*Target: Assuming ", (1-as.numeric(reduction_region_GEM()))*100, "% complaints reduction from previous year."), shortcuts$fp_italic(color="black"))) 
                                   )
                                   
                                   reduction_EUR <- block_list(
                                     fpar(ftext(paste0("*Target: Assuming ", (1-as.numeric(input$reduction_EUR))*100, "% complaints reduction from previous year."), shortcuts$fp_italic(color="black"))) 
                                   )
                                   
                                   reduction_NAM <- block_list(
                                     fpar(ftext(paste0("*Target: Assuming ", (1-as.numeric(input$reduction_NAM))*100, "% complaints reduction from previous year."), shortcuts$fp_italic(color="black"))) 
                                   )
                                   
                                   reduction_AP <- block_list(
                                     fpar(ftext(paste0("*Target: Assuming ", (1-as.numeric(input$reduction_AP))*100, "% complaints reduction from previous year."), shortcuts$fp_italic(color="black"))) 
                                   )
                                   
                                   reduction_LA <- block_list(
                                     fpar(ftext(paste0("*Target: Assuming ", (1-as.numeric(input$reduction_LA))*100, "% complaints reduction from previous year."), shortcuts$fp_italic(color="black"))) 
                                   )
                                   
                                   reduction_MEA <- block_list(
                                     fpar(ftext(paste0("*Target: Assuming ", (1-as.numeric(input$reduction_MEA))*100, "% complaints reduction from previous year."), shortcuts$fp_italic(color="black"))) 
                                   )
                              
                                   
                                   tb_BSC <- flextable(BSC) %>%
                                     merge_at(i = 1, j = 5:6, part = "body") %>%
                                     merge_at(i = 9, j = 5:6, part = "body") %>%
                                     merge_at(i = 1:2, j = 1:2, part = "body") %>%
                                     merge_at(i = 1:2, j = 3, part = "body") %>%
                                     merge_at(i = 1:2, j = 4, part = "body") %>%
                                     merge_at(i = 9:10, j = 1:2, part = "body") %>%
                                     merge_at(i = 9:10, j = 3, part = "body") %>%
                                     merge_at(i = 9:10, j = 4, part = "body") %>%
                                     merge_h_range(i = 1:13, j1 = 1, j2 = 2, part = "body") %>%
                                     bold(j = 1, part = "body") %>%
                                     bold(i = c(1, 2, 9, 10),  part = "body") %>%
                                     bg(i = c(1, 2, 9, 10), bg = "steelblue", part = "body") %>%
                                     bg(i = c(3, 5, 7, 11, 13), bg = "grey90", part = "body") %>%
                                     bg(i = c(4, 6, 8, 12), bg = "grey95", part = "body") %>%
                                     color(i = c(1, 2, 9, 10), color = "white", part = "body") %>%
                                     align(align = "center", part = "all")%>%
                                     delete_part(part = "header") %>%
                                     border(i = 1, border.top = fp_border(color = "black", style = "solid", width = 1.5), part = "body") %>%
                                     border(i = 2, border.top = fp_border(color = "black", style = "solid", width = 1.5), part = "body") %>%
                                     height(height = 0.38) %>%
                                     width(width = 1.55)
                                   
                                   tb_BSC_region <- flextable(BSC_region) %>%
                                     merge_at(i = 1, j = 6:7, part = "body") %>%
                                     merge_at(i = 1:2, j = 1:3, part = "body") %>%
                                     merge_at(i = 1:2, j = 4, part = "body") %>%
                                     merge_at(i = 1:2, j = 5, part = "body") %>%
                                     merge_at(i = 3:6, j = 1:2, part = "body") %>%
                                     merge_h_range(i = c(1, 2), j1 = 1, j2 = 3, part = "body") %>%
                                     merge_h_range(i = c(3:8), j1 = 1, j2 = 2, part = "body") %>%
                                     bold(j = 1, part = "body") %>%
                                     bold(i = c(1, 2),  part = "body") %>%
                                     bg(i = c(1, 2), bg = "steelblue", part = "body") %>%
                                     bg(i = c(3:5), j = 3:7, bg = "lightskyblue", part = "body") %>%
                                     bg(i = 6, j = 3:7, bg = "steelblue", part = "body") %>%
                                     bg(i = 7, j = 3:7, bg = "aquamarine4", part = "body") %>%
                                     bg(i = 8, j = 3:7, bg = "darkorange1", part = "body") %>%
                                     color(i = c(1, 2), color = "white", part = "body") %>%
                                     align(align = "center", part = "all")%>%
                                     delete_part(part = "header") %>%
                                     border(i = 1, border.top = fp_border(color = "black", style = "solid", width = 1.5), part = "body") %>%
                                     #border(i = 2, j = 1:7, border.bottom  = fp_border(color = "black", style = "solid", width = 1.5), part = "body") %>%
                                     border(i = 3:7, border.bottom  = fp_border(color = "black", style = "solid", width = 1), part = "body") %>%
                                     height(height = 0.4) %>%
                                     width(width = 1.3)
                                   
                                   tb_GMSC <- flextable(GMSC) %>% 
                                     set_header_labels(BU = "Business Unit", na1 = "Business Unit", na2 = "Business Unit",
                                                       OC = "Ostomy Care", na3 = "Ostomy Care",
                                                       AWC = "Wound Care", na4 = "Wound Care",
                                                       CC = "Continence Care", na5 = "Continence Care",
                                                       C = "Critical Care", na6 = "Critical Care",
                                                       NIC = "All non-IC BUs combined", na7 = "All non-IC BUs combined",
                                                       IC = "Infusion Care", na8 = "Infusion Care",
                                                       ALL = "All BUs Combined", na9 = "All BUs Combined") %>%
                                     merge_h(part = "header") %>%
                                     merge_h_range(i = 1:17, j1 = 1, j2 = 3, part = "body") %>%
                                     merge_h_range(i = 1:17, j1 = 4, j2 = 5, part = "body") %>%
                                     merge_h_range(i = 1:17, j1 = 6, j2 = 7, part = "body") %>%
                                     merge_h_range(i = 1:17, j1 = 8, j2 = 9, part = "body") %>%
                                     merge_h_range(i = 1:17, j1 = 10, j2 = 11, part = "body") %>%
                                     merge_h_range(i = 1:17, j1 = 12, j2 = 13, part = "body") %>%
                                     merge_h_range(i = 1:17, j1 = 14, j2 = 15, part = "body") %>%
                                     merge_h_range(i = 1:17, j1 = 16, j2 = 17, part = "body") %>%
                                     bold(part = "header") %>%
                                     bold(i = c(1:3, 16:17), j = 1, part = "body") %>%
                                     fontsize(size = 11, part = "all") %>% 
                                     bg(bg = "steelblue", part = "header") %>%
                                     color(color = "white", part = "header") %>%
                                     bg(i = c(1, 3, 5, 7, 9, 11, 13, 15, 17), bg = "grey90", part = "body") %>%
                                     bg(i = c(2, 4, 6, 8, 10, 12, 14, 16), bg = "grey95", part = "body") %>%
                                     align(align = "center", part = "all")%>%
                                     #border_outer(border = std_border, part = "all") %>%
                                     #border_inner(border = std_border, part = "all")  %>%
                                     hline(i = c(3, 15, 17), border = std_border, part = "body")%>%
                                     height(height = 0.28) %>%
                                     width(width = 0.56)
                                   
                                   CPM_COM_NIC <- flextable(CPM_COM_NIC) %>% 
                                     merge_at(i = 1:2, j = 1, part = "body") %>%
                                     merge_at(i = 3:4, j = 1, part = "body") %>%
                                     merge_at(i = 5:6, j = 1, part = "body") %>%
                                     bg(i = 1, j = 1, bg = "dodgerblue3", part = "body") %>%
                                     bg(i = 3, j = 1, bg = "darkorange", part = "body") %>%
                                     bg(i = 5, j = 1, bg = "darkolivegreen4", part = "body") %>%
                                     bg(i = c(1, 3, 5), j = 2:14, bg = "grey90", part = "body") %>%
                                     bg(i = c(2, 4, 6), j = 2:14, bg = "grey95", part = "body") %>%
                                     bold(part = "header") %>%
                                     hline(i = c(2, 4, 6), border = std_border, part = "body")%>%
                                     hline(j = 1, border = std_border, part = "body")%>%
                                     align(align = "center", part = "all")%>%
                                     #height(height = 0.18) %>%
                                     width(j = 1, width = 0.5) %>%
                                     width(j = c(3:14), width = 0.62) %>%
                                     width(j = 2, width = 0.95)
                                   
                                   tb_ALL <- flextable(ALL) %>% 
                                     set_header_labels(na = "Month",
                                                       na1 = "Month",
                                                       #pastyear2 = paste0(pastyear2, "-Overall"), 
                                                       pastyear1 = paste0(pastyear1, "-Overall"), 
                                                       currentyear_target = paste0(currentyear, "-Target*"), 
                                                       currentyear_actual = paste0(currentyear, "-Actual")) %>%
                                     
                                     merge_h(part = "header") %>%
                                     merge_h_range(i = 1:13, j1 = 1, j2 = 2, part = "body") %>%
                                     merge_at(i = 1:13, j = 3, part = "body") %>%
                                     merge_at(i = 1:13, j = 4, part = "body") %>%
                                     
                                     bold(part = "header") %>%
                                     bold(i = 13, part = "body") %>%
                                     fontsize(size = 14, part = "all") %>% 
                                     fontsize(i = 13, size = 15, part = "body") %>% 
                                     bg(bg = "coral4", part = "header") %>%
                                     color(color = "white", part = "header") %>%
                                     width(width = 1.1) %>% 
                                     height(height = 0.45, part = "header") %>%
                                     height(height = 0.40, part = "body") %>%
                                     align(align = "center", part = "all")%>%
                                     border_outer(border = std_border, part = "all") %>%
                                     border_inner(border = std_border, part = "all") 
                                   
                                   tb_OC <- flextable(OC) %>% 
                                     set_header_labels(na = "Month",
                                                       na1 = "Month",
                                                       #pastyear2 = paste0(pastyear2, "-Overall"), 
                                                       pastyear1 = paste0(pastyear1, "-Overall"), 
                                                       currentyear_target = paste0(currentyear, "-Target*"), 
                                                       currentyear_actual = paste0(currentyear, "-Actual")) %>%
                                     
                                     merge_h(part = "header") %>%
                                     merge_h_range(i = 1:13, j1 = 1, j2 = 2, part = "body") %>% 
                                     merge_at(i = 1:13, j = 3, part = "body") %>%
                                     merge_at(i = 1:13, j = 4, part = "body") %>%
                                     
                                     bold(part = "header") %>%
                                     bold(i = 13, part = "body") %>%
                                     fontsize(size = 14, part = "all") %>% 
                                     fontsize(i = 13, size = 15, part = "body") %>% 
                                     bg(bg = "aquamarine4", part = "header") %>%
                                     color(color = "white", part = "header") %>%
                                     width(width = 1.1) %>% 
                                     height(height = 0.45, part = "header") %>%
                                     height(height = 0.40, part = "body") %>%
                                     align(align = "center", part = "all")%>%
                                     border_outer(border = std_border, part = "all") %>%
                                     border_inner(border = std_border, part = "all") 
                                   
                                   
                                   tb_AWC <- flextable(AWC) %>% 
                                     set_header_labels(na = "Month",
                                                       na1 = "Month",
                                                       #pastyear2 = paste0(pastyear2, "-Overall"), 
                                                       pastyear1 = paste0(pastyear1, "-Overall"), 
                                                       currentyear_target = paste0(currentyear, "-Target*"), 
                                                       currentyear_actual = paste0(currentyear, "-Actual")) %>%
                                     
                                     merge_h(part = "header") %>%
                                     merge_h_range(i = 1:13, j1 = 1, j2 = 2, part = "body") %>%
                                     merge_at(i = 1:13, j = 3, part = "body") %>%
                                     merge_at(i = 1:13, j = 4, part = "body") %>%
                                     #merge_at(i = 1:13, j = 5, part = "body") %>%
                                     
                                     bold(part = "header") %>%
                                     bold(i = 13, part = "body") %>%
                                     fontsize(size = 14, part = "all") %>% 
                                     fontsize(i = 13, size = 15, part = "body") %>% 
                                     bg(bg = "dodgerblue4", part = "header") %>%
                                     color(color = "white", part = "header") %>%
                                     width(width = 1.1) %>% 
                                     height(height = 0.45, part = "header") %>%
                                     height(height = 0.40, part = "body") %>%
                                     align(align = "center", part = "all")%>%
                                     border_outer(border = std_border, part = "all") %>%
                                     border_inner(border = std_border, part = "all") 
                                   
                                   tb_CC <- flextable(CC) %>% 
                                     set_header_labels(na = "Month",
                                                       na1 = "Month",
                                                       #pastyear2 = paste0(pastyear2, "-Overall"), 
                                                       pastyear1 = paste0(pastyear1, "-Overall"), 
                                                       currentyear_target = paste0(currentyear, "-Target*"), 
                                                       currentyear_actual = paste0(currentyear, "-Actual")) %>%
                                     
                                     merge_h(part = "header") %>%
                                     merge_h_range(i = 1:13, j1 = 1, j2 = 2, part = "body") %>% 
                                     merge_at(i = 1:13, j = 3, part = "body") %>%
                                     merge_at(i = 1:13, j = 4, part = "body") %>%
                                     #merge_at(i = 1:13, j = 5, part = "body") %>%
                                     
                                     bold(part = "header") %>%
                                     bold(i = 13, part = "body") %>%
                                     fontsize(size = 14, part = "all") %>% 
                                     fontsize(i = 13, size = 15, part = "body") %>% 
                                     bg(bg = "darkgoldenrod1", part = "header") %>%
                                     color(color = "white", part = "header") %>%
                                     width(width = 1.1) %>% 
                                     height(height = 0.45, part = "header") %>%
                                     height(height = 0.40, part = "body") %>%
                                     align(align = "center", part = "all")%>%
                                     border_outer(border = std_border, part = "all") %>%
                                     border_inner(border = std_border, part = "all") 
                                   
                                   tb_C <- flextable(C) %>% 
                                     set_header_labels(na = "Month",
                                                       na1 = "Month",
                                                       #pastyear2 = paste0(pastyear2, "-Overall"), 
                                                       pastyear1 = paste0(pastyear1, "-Overall"), 
                                                       currentyear_target = paste0(currentyear, "-Target*"), 
                                                       currentyear_actual = paste0(currentyear, "-Actual")) %>%
                                     
                                     merge_h(part = "header") %>%
                                     merge_h_range(i = 1:13, j1 = 1, j2 = 2, part = "body") %>% 
                                     merge_at(i = 1:13, j = 3, part = "body") %>%
                                     merge_at(i = 1:13, j = 4, part = "body") %>%
                                     #merge_at(i = 1:13, j = 5, part = "body") %>%
                                     
                                     bold(part = "header") %>%
                                     bold(i = 13, part = "body") %>%
                                     fontsize(size = 14, part = "all") %>% 
                                     fontsize(i = 13, size = 15, part = "body") %>% 
                                     bg(bg = "cornsilk4", part = "header") %>%
                                     color(color = "white", part = "header") %>%
                                     width(width = 1.1) %>% 
                                     height(height = 0.45, part = "header") %>%
                                     height(height = 0.40, part = "body") %>%
                                     align(align = "center", part = "all")%>%
                                     border_outer(border = std_border, part = "all") %>%
                                     border_inner(border = std_border, part = "all") 
                                   
                                   tb_IC <- flextable(IC) %>% 
                                     set_header_labels(na = "Month",
                                                       na1 = "Month",
                                                       #pastyear2 = paste0(pastyear2, "-Overall"), 
                                                       pastyear1 = paste0(pastyear1, "-Overall"), 
                                                       currentyear_target = paste0(currentyear, "-Target*"), 
                                                       currentyear_actual = paste0(currentyear, "-Actual")) %>%
                                     
                                     merge_h(part = "header") %>%
                                     merge_h_range(i = 1:13, j1 = 1, j2 = 2, part = "body") %>%
                                     merge_at(i = 1:13, j = 3, part = "body") %>%
                                     merge_at(i = 1:13, j = 4, part = "body") %>%
                                     #merge_at(i = 1:13, j = 5, part = "body") %>%
                                     
                                     bold(part = "header") %>%
                                     bold(i = 13, part = "body") %>%
                                     fontsize(size = 14, part = "all") %>% 
                                     fontsize(i = 13, size = 15, part = "body") %>% 
                                     bg(bg = "darkorange1", part = "header") %>%
                                     color(color = "white", part = "header") %>%
                                     width(width = 1.1) %>% 
                                     height(height = 0.45, part = "header") %>%
                                     height(height = 0.40, part = "body") %>%
                                     align(align = "center", part = "all")%>%
                                     border_outer(border = std_border, part = "all") %>%
                                     border_inner(border = std_border, part = "all")  
                                   
                                   tb_Eurotec <- flextable(Eurotec) %>% 
                                     set_header_labels(na = "Month",
                                                       na1 = "Month",
                                                       #pastyear2 = paste0(pastyear2, "-Overall"), 
                                                       pastyear1 = paste0(pastyear1, "-Overall"), 
                                                       currentyear_target = paste0(currentyear, "-Target*"), 
                                                       currentyear_actual = paste0(currentyear, "-Actual")) %>%
                                     
                                     merge_h(part = "header") %>%
                                     merge_h_range(i = 1:13, j1 = 1, j2 = 2, part = "body") %>%
                                     merge_at(i = 1:13, j = 3, part = "body") %>%
                                     merge_at(i = 1:13, j = 4, part = "body") %>%
                                     #merge_at(i = 1:13, j = 5, part = "body") %>%
                                     
                                     bold(part = "header") %>%
                                     bold(i = 13, part = "body") %>%
                                     fontsize(size = 14, part = "all") %>% 
                                     fontsize(i = 13, size = 15, part = "body") %>% 
                                     bg(bg = "aquamarine3", part = "header") %>%
                                     color(color = "white", part = "header") %>%
                                     width(width = 1.1) %>% 
                                     height(height = 0.45, part = "header") %>%
                                     height(height = 0.40, part = "body") %>%
                                     align(align = "center", part = "all")%>%
                                     border_outer(border = std_border, part = "all") %>%
                                     border_inner(border = std_border, part = "all")  
                                   
                                   tb_NIC <- flextable(NIC) %>% 
                                     set_header_labels(na = "Month",
                                                       na1 = "Month",
                                                       #pastyear2 = paste0(pastyear2, "-Overall"), 
                                                       pastyear1 = paste0(pastyear1, "-Overall"), 
                                                       currentyear_target = paste0(currentyear, "-Target*"), 
                                                       currentyear_actual = paste0(currentyear, "-Actual")) %>%
                                     
                                     merge_h(part = "header") %>%
                                     merge_h_range(i = 1:13, j1 = 1, j2 = 2, part = "body") %>% 
                                     merge_at(i = 1:13, j = 3, part = "body") %>%
                                     merge_at(i = 1:13, j = 4, part = "body") %>%
                                     #merge_at(i = 1:13, j = 5, part = "body") %>%
                                     
                                     bold(part = "header") %>%
                                     bold(i = 13, part = "body") %>%
                                     fontsize(size = 14, part = "all") %>% 
                                     fontsize(i = 13, size = 15, part = "body") %>% 
                                     bg(bg = "darkcyan", part = "header") %>%
                                     color(color = "white", part = "header") %>%
                                     width(width = 1.1) %>% 
                                     height(height = 0.45, part = "header") %>%
                                     height(height = 0.40, part = "body") %>%
                                     align(align = "center", part = "all")%>%
                                     border_outer(border = std_border, part = "all") %>%
                                     border_inner(border = std_border, part = "all")
                                   
                                   tb_GEM <- flextable(GEM) %>% 
                                     set_header_labels(na = "Month",
                                                       na1 = "Month",
                                                       #pastyear2 = paste0(pastyear2, "-Overall"), 
                                                       pastyear1 = paste0(pastyear1, "-Overall"), 
                                                       currentyear_target = paste0(currentyear, "-Target*"), 
                                                       currentyear_actual = paste0(currentyear, "-Actual")) %>%
                                     
                                     merge_h(part = "header") %>%
                                     merge_h_range(i = 1:13, j1 = 1, j2 = 2, part = "body") %>% 
                                     merge_at(i = 1:13, j = 3, part = "body") %>%
                                     merge_at(i = 1:13, j = 4, part = "body") %>%
                                     
                                     bold(part = "header") %>%
                                     bold(i = 13, part = "body") %>%
                                     fontsize(size = 14, part = "all") %>% 
                                     fontsize(i = 13, size = 15, part = "body") %>% 
                                     bg(bg = "dodgerblue4", part = "header") %>%
                                     color(color = "white", part = "header") %>%
                                     width(width = 1.1) %>% 
                                     height(height = 0.45, part = "header") %>%
                                     height(height = 0.40, part = "body") %>%
                                     align(align = "center", part = "all")%>%
                                     border_outer(border = std_border, part = "all") %>%
                                     border_inner(border = std_border, part = "all") 
                                   
                                   
                                   tb_EUR <- flextable(EUR) %>% 
                                     set_header_labels(na = "Month",
                                                       na1 = "Month",
                                                       #pastyear2 = paste0(pastyear2, "-Overall"), 
                                                       pastyear1 = paste0(pastyear1, "-Overall"), 
                                                       currentyear_target = paste0(currentyear, "-Target*"), 
                                                       currentyear_actual = paste0(currentyear, "-Actual")) %>%
                                     
                                     merge_h(part = "header") %>%
                                     merge_h_range(i = 1:13, j1 = 1, j2 = 2, part = "body") %>%
                                     merge_at(i = 1:13, j = 3, part = "body") %>%
                                     merge_at(i = 1:13, j = 4, part = "body") %>%
                                     #merge_at(i = 1:13, j = 5, part = "body") %>%
                                     
                                     bold(part = "header") %>%
                                     bold(i = 13, part = "body") %>%
                                     fontsize(size = 14, part = "all") %>% 
                                     fontsize(i = 13, size = 15, part = "body") %>% 
                                     bg(bg = "darkorange1", part = "header") %>%
                                     color(color = "white", part = "header") %>%
                                     width(width = 1.1) %>% 
                                     height(height = 0.45, part = "header") %>%
                                     height(height = 0.40, part = "body") %>%
                                     align(align = "center", part = "all")%>%
                                     border_outer(border = std_border, part = "all") %>%
                                     border_inner(border = std_border, part = "all") 
                                   
                                   tb_NAM <- flextable(NAM) %>% 
                                     set_header_labels(na = "Month",
                                                       na1 = "Month",
                                                       #pastyear2 = paste0(pastyear2, "-Overall"), 
                                                       pastyear1 = paste0(pastyear1, "-Overall"), 
                                                       currentyear_target = paste0(currentyear, "-Target*"), 
                                                       currentyear_actual = paste0(currentyear, "-Actual")) %>%
                                     
                                     merge_h(part = "header") %>%
                                     merge_h_range(i = 1:13, j1 = 1, j2 = 2, part = "body") %>%
                                     merge_at(i = 1:13, j = 3, part = "body") %>%
                                     merge_at(i = 1:13, j = 4, part = "body") %>%
                                     #merge_at(i = 1:13, j = 5, part = "body") %>%
                                     
                                     bold(part = "header") %>%
                                     bold(i = 13, part = "body") %>%
                                     fontsize(size = 14, part = "all") %>% 
                                     fontsize(i = 13, size = 15, part = "body") %>% 
                                     bg(bg = "aquamarine4", part = "header") %>%
                                     color(color = "white", part = "header") %>%
                                     width(width = 1.1) %>% 
                                     height(height = 0.45, part = "header") %>%
                                     height(height = 0.40, part = "body") %>%
                                     align(align = "center", part = "all")%>%
                                     border_outer(border = std_border, part = "all") %>%
                                     border_inner(border = std_border, part = "all") 
                                   
                                   tb_AP <- flextable(AP) %>% 
                                     set_header_labels(na = "Month",
                                                       na1 = "Month",
                                                       #pastyear2 = paste0(pastyear2, "-Overall"), 
                                                       pastyear1 = paste0(pastyear1, "-Overall"), 
                                                       currentyear_target = paste0(currentyear, "-Target*"), 
                                                       currentyear_actual = paste0(currentyear, "-Actual")) %>%
                                     
                                     merge_h(part = "header") %>%
                                     merge_h_range(i = 1:13, j1 = 1, j2 = 2, part = "body") %>%
                                     merge_at(i = 1:13, j = 3, part = "body") %>%
                                     merge_at(i = 1:13, j = 4, part = "body") %>%
                                     #merge_at(i = 1:13, j = 5, part = "body") %>%
                                     
                                     bold(part = "header") %>%
                                     bold(i = 13, part = "body") %>%
                                     fontsize(size = 14, part = "all") %>% 
                                     fontsize(i = 13, size = 15, part = "body") %>% 
                                     bg(bg = "cornflowerblue", part = "header") %>%
                                     color(color = "white", part = "header") %>%
                                     width(width = 1.1) %>% 
                                     height(height = 0.45, part = "header") %>%
                                     height(height = 0.40, part = "body") %>%
                                     align(align = "center", part = "all")%>%
                                     border_outer(border = std_border, part = "all") %>%
                                     border_inner(border = std_border, part = "all") 
                                   
                                   tb_LA <- flextable(LA) %>% 
                                     set_header_labels(na = "Month",
                                                       na1 = "Month",
                                                       #pastyear2 = paste0(pastyear2, "-Overall"), 
                                                       pastyear1 = paste0(pastyear1, "-Overall"), 
                                                       currentyear_target = paste0(currentyear, "-Target*"), 
                                                       currentyear_actual = paste0(currentyear, "-Actual")) %>%
                                     
                                     merge_h(part = "header") %>%
                                     merge_h_range(i = 1:13, j1 = 1, j2 = 2, part = "body") %>%
                                     merge_at(i = 1:13, j = 3, part = "body") %>%
                                     merge_at(i = 1:13, j = 4, part = "body") %>%
                                     #merge_at(i = 1:13, j = 5, part = "body") %>%
                                     
                                     bold(part = "header") %>%
                                     bold(i = 13, part = "body") %>%
                                     fontsize(size = 14, part = "all") %>% 
                                     fontsize(i = 13, size = 15, part = "body") %>% 
                                     bg(bg = "cornflowerblue", part = "header") %>%
                                     color(color = "white", part = "header") %>%
                                     width(width = 1.1) %>% 
                                     height(height = 0.45, part = "header") %>%
                                     height(height = 0.40, part = "body") %>%
                                     align(align = "center", part = "all")%>%
                                     border_outer(border = std_border, part = "all") %>%
                                     border_inner(border = std_border, part = "all") 
                                   
                                   tb_MEA <- flextable(MEA) %>% 
                                     set_header_labels(na = "Month",
                                                       na1 = "Month",
                                                       #pastyear2 = paste0(pastyear2, "-Overall"), 
                                                       pastyear1 = paste0(pastyear1, "-Overall"), 
                                                       currentyear_target = paste0(currentyear, "-Target*"), 
                                                       currentyear_actual = paste0(currentyear, "-Actual")) %>%
                                     
                                     merge_h(part = "header") %>%
                                     merge_h_range(i = 1:13, j1 = 1, j2 = 2, part = "body") %>% 
                                     merge_at(i = 1:13, j = 3, part = "body") %>%
                                     merge_at(i = 1:13, j = 4, part = "body") %>%
                                     
                                     bold(part = "header") %>%
                                     bold(i = 13, part = "body") %>%
                                     fontsize(size = 14, part = "all") %>% 
                                     fontsize(i = 13, size = 15, part = "body") %>% 
                                     bg(bg = "cornflowerblue", part = "header") %>%
                                     color(color = "white", part = "header") %>%
                                     width(width = 1.1) %>% 
                                     height(height = 0.45, part = "header") %>%
                                     height(height = 0.40, part = "body") %>%
                                     align(align = "center", part = "all")%>%
                                     border_outer(border = std_border, part = "all") %>%
                                     border_inner(border = std_border, part = "all") 
                                   
                                   example_pp <- read_pptx() %>% 
                                     
                                     add_slide(layout = "Title Slide", master = "Office Theme") %>% 
                                               ph_with(external_img(src = titlepageImage), 
                                                       location = ph_location(left = 0, top = 0, width = 10, height = 7.5) )%>% 
                                               ph_with(value = title_BU,
                                                       location = ph_location(top = 0.5, left = 0.2, width = 10, height = 1) ) %>%
                                               ph_with(value = title2, 
                                                       location = ph_location(top = 1.5, left = 0.2, width = 10, height = 1) ) %>%
                                     
                                     #add_slide(layout = "Title and Content", master = "Office Theme") %>% 
                                      #         ph_with(external_img(src = leftCornerImage), 
                                       #                location = ph_location(left = 0, top = 0, width = 2.2, height = 1.5) )%>%
                                        #       ph_with(external_img(src = logoImage), 
                                         #              location = ph_location(left = 7.8, top = 6, width = 2.2, height = 1.5) ) %>%
                                          #     ph_with(location = ph_location(top = 0.1, left = 2.6, width = 10, height = 1),
                                           #            value = caption_BSC) %>% 
                                            #   ph_with_flextable_at(value = tb_BSC, left = 0.35, top = 1.3) %>%
                                             #  ph_with(value = note_BSC1,
                                              #         location = ph_location(left = 0.3, top = 6, width = 9, height = 1, label = "note_BSC1")) %>%
                                               #ph_with(value = note_BSC2,
                                                #       location = ph_location(left = 0.3, top = 6.2, width = 9, height = 1, label = "note_BSC2")) %>%
                                              # ph_with(value = note_BSC3,
                                               #        location = ph_location(left = 0.3, top = 6.4, width = 9, height = 1, label = "note_BSC3")) %>%
                                               #ph_with(value = note_BSC4,
                                              #         location = ph_location(left = 0.3, top = 6.6, width = 9, height = 1, label = "note_BSC4")) %>%
                                               #ph_with(value = note_BSC5,
                                                #       location = ph_location(left = 0.3, top = 6.8, width = 9, height = 1, label = "note_BSC5")) %>%
                                   
                                     
                                     add_slide(layout = "Title and Content", master = "Office Theme") %>%
                                               ph_with(external_img(src = leftCornerImage),
                                                       location = ph_location(left = 0, top = 0, width = 2.2, height = 1.5) )%>%
                                               ph_with(external_img(src = logoImage),
                                                       location = ph_location(left = 7.8, top = 6, width = 2.2, height = 1.5) ) %>%
                                               ph_with(value = caption_GMSC,
                                                       location = ph_location(top = 0.1, left = 2.6, width = 10, height = 1) ) %>%
                                               ph_with(value = tb_GMSC,
                                                       location = ph_location(left = 0.2, top = 1.2)) %>%
                                               ph_with(value = note_GMSC1,
                                                       location = ph_location(left = 0.3, top = 6.5, width = 10, height = 0.5, label = "note_GMSC1")) %>%
                                               ph_with(value = note_GMSC2,
                                                       location = ph_location(left = 0.3, top = 6.7, width = 10, height = 0.5, label = "note_GMSC2")) %>%
                                               ph_with(value = note_GMSC3,
                                                       location = ph_location(left = 0.3, top = 6.9, width = 10, height = 0.5, label = "note_GMSC3")) %>%
                                               ph_with(value = note_GMSC4,
                                                       location = ph_location(left = 0.3, top = 7.1, width = 10, height = 0.5, label = "note_GMSC4")) %>%

                                     add_slide(layout = "Title and Content", master = "Office Theme") %>%
                                               ph_with(external_img(src = leftCornerImage),
                                                       location = ph_location(left = 0, top = 0, width = 2.2, height = 1.5) ) %>%
                                               ph_with(external_img(src = logoImage),
                                                       location = ph_location(left = 7.8, top = 6, width = 2.2, height = 1.5) ) %>%
                                               ph_with(location = ph_location(top = 0, left = 3, width = 10, height = 1),
                                                       value = caption_combined) %>%
                                               ph_with(value = dml(Trend_Chart6()),
                                                       location = ph_location(left = 0.2, top = 0.7, width = 9.5, height = 7)) %>%


                                     add_slide(layout = "Title and Content", master = "Office Theme") %>%
                                               ph_with(external_img(src = leftCornerImage),
                                                       location = ph_location(left = 0, top = 0, width = 2.2, height = 1.5) )%>%
                                               ph_with(external_img(src = logoImage),
                                                       location = ph_location(left = 7.8, top = 6, width = 2.2, height = 1.5) ) %>%
                                               ph_with(value = caption_BSC_region,
                                                       location = ph_location(top = 1, left = 2.3, width = 10, height = 1) ) %>%
                                               ph_with(value = tb_BSC_region,
                                                       location = ph_location(left = 0.35, top = 2.4) )%>%
                                               ph_with(value = note_BSC3,
                                                       location = ph_location(left = 0.3, top = 5.4, width = 9, height = 1, label = "note_BSC3")) %>%
                                               ph_with(value = note_BSC4,
                                                       location = ph_location(left = 0.3, top = 5.6, width = 9, height = 1, label = "note_BSC4")) %>%
                                               ph_with(value = note_BSC5,
                                                       location = ph_location(left = 0.3, top = 5.8, width = 9, height = 1, label = "note_BSC5")) %>%
                                               ph_with(value = note_BSC2,
                                                       location = ph_location(left = 0.3, top = 6, width = 9, height = 1, label = "note_BSC2")) %>%
                                               ph_with(value = note_BSC6,
                                                       location = ph_location(left = 0.3, top = 6.2, width = 9, height = 1, label = "note_BSC6")) %>%

                                     add_slide(layout = "Title and Content", master = "Office Theme") %>%
                                               ph_with(external_img(src = leftCornerImage),
                                                       location = ph_location(left = 0, top = 0, width = 2.2, height = 1.5) ) %>%
                                               ph_with(external_img(src = logoImage),
                                                       location = ph_location(left = 7.8, top = 6, width = 2.2, height = 1.5) ) %>%
                                               ph_with(value = dml(Bar_Line()),
                                                       location = ph_location(left = 0.35, top = 0.2, width = 9.5, height = 5.3) )%>%
                                               ph_with(value = CPM_COM_NIC,
                                                       location = ph_location(left = 0.38, top = 5.6)) %>%

                                     add_slide(layout = "Title and Content", master = "Office Theme") %>%
                                               ph_with(external_img(src = leftCornerImage),
                                                       location = ph_location(left = 0, top = 0, width = 2.2, height = 1.5) )%>%
                                               ph_with(external_img(src = logoImage),
                                                       location = ph_location(left = 7.8, top = 6, width = 2.2, height = 1.5) ) %>%
                                               ph_with(value = caption_all,
                                                       location = ph_location(top = 0.1, left = 3.3, width = 10, height = 1) ) %>%
                                               ph_with(value = tb_ALL,
                                                       location = ph_location(left = 2.2, top = 1.1)) %>%
                                               ph_with(value = note_BU,
                                                       location = ph_location(left = 2.2, top = 6.4, width = 6, height = 1, label = "note_BU")) %>%
                                               ph_with(value = note_SKC,
                                                       location = ph_location(left = 2.2, top = 6.65, width = 6, height = 1, label = "note_SKC")) %>%

                                     add_slide(layout = "Title and Content", master = "Office Theme") %>%
                                               ph_with(external_img(src = leftCornerImage),
                                                       location = ph_location(left = 0, top = 0, width = 2.2, height = 1.5) )%>%
                                               ph_with(external_img(src = logoImage),
                                                       location = ph_location(left = 7.8, top = 6, width = 2.2, height = 1.5) ) %>%
                                               ph_with(value = caption_NIC,
                                                       location = ph_location(top = 0.1, left = 1.6, width = 10, height = 1)) %>%
                                               ph_with(value = tb_NIC,
                                                       location = ph_location(left = 2.2, top = 1.1) )%>%
                                               ph_with(value = note_NIC,
                                                       location = ph_location(left = 2.2, top = 6.5, width = 6, height = 1, label = "note_NIC")) %>%
                                     

                                     add_slide(layout = "Title and Content", master = "Office Theme") %>%
                                               ph_with(external_img(src = leftCornerImage),
                                                       location = ph_location(left = 0, top = 0, width = 2.2, height = 1.5) ) %>%
                                               ph_with(external_img(src = logoImage),
                                                       location = ph_location(left = 7.8, top = 6, width = 2.2, height = 1.5) ) %>%
                                               ph_with(location = ph_location(top = 0.1, left = 2.9, width = 10, height = 1),
                                                       value = caption_IC) %>%
                                               ph_with(value = tb_IC,
                                                       location = ph_location(left = 2.2, top = 1.1) )%>%
                                               ph_with(value = note_IC,
                                                       location = ph_location(left = 2.2, top = 6.5, width = 6, height = 1, label = "note_IC")) %>%

                                     
                                     add_slide(layout = "Title and Content", master = "Office Theme") %>%
                                               ph_with(external_img(src = leftCornerImage),
                                                       location = ph_location(left = 0, top = 0, width = 2.2, height = 1.5) )%>%
                                               ph_with(external_img(src = logoImage),
                                                       location = ph_location(left = 7.8, top = 6, width = 2.2, height = 1.5) ) %>%
                                               ph_with(value = caption_WC,
                                                       location = ph_location(top = 0.1, left = 3.1, width = 10, height = 1) ) %>%
                                               ph_with(value = tb_AWC,
                                                       location = ph_location(left = 2.2, top = 1.1) )%>%
                                               ph_with(value = note_AWC,
                                                       location = ph_location(left = 2.2, top = 6.5, width = 6, height = 1, label = "note_AWC")) %>%

                                     add_slide(layout = "Title and Content", master = "Office Theme") %>%
                                               ph_with(external_img(src = leftCornerImage),
                                                       location = ph_location(left = 0, top = 0, width = 2.2, height = 1.5) )%>%
                                               ph_with(external_img(src = logoImage),
                                                       location = ph_location(left = 7.8, top = 6, width = 2.2, height = 1.5) ) %>%
                                               ph_with(value = caption_CC,
                                                      location = ph_location(top = 0.1, left = 2.7, width = 10, height = 1) ) %>%
                                               ph_with(value = tb_CC,
                                                       location = ph_location(left = 2.2, top = 1.1) )%>%
                                               ph_with(value = note_CC,
                                                       location = ph_location(left = 2.2, top = 6.5, width = 6, height = 1, label = "note_CC")) %>%

                                     add_slide(layout = "Title and Content", master = "Office Theme") %>%
                                               ph_with(external_img(src = leftCornerImage),
                                                       location = ph_location(left = 0, top = 0, width = 2.2, height = 1.5) )%>%
                                               ph_with(external_img(src = logoImage),
                                                       location = ph_location(left = 7.8, top = 6, width = 2.2, height = 1.5) ) %>%
                                               ph_with(value = caption_C,
                                                       location = ph_location(top = 0.1, left = 3.1, width = 10, height = 1) ) %>%
                                               ph_with(value = tb_C,
                                                       location = ph_location(left = 2.2, top = 1.1) )%>%
                                               ph_with(value = note_C,
                                                       location = ph_location(left = 2.2, top = 6.5, width = 6, height = 1, label = "note_C")) %>%
                                     
                                     add_slide(layout = "Title and Content", master = "Office Theme") %>%
                                               ph_with(external_img(src = leftCornerImage),
                                                       location = ph_location(left = 0, top = 0, width = 2.2, height = 1.5) )%>%
                                               ph_with(external_img(src = logoImage),
                                                       location = ph_location(left = 7.8, top = 6, width = 2.2, height = 1.5) ) %>%
                                               ph_with(value = caption_OC,
                                                       location = ph_location(top = 0.1, left = 3.1, width = 10, height = 1)) %>%
                                               ph_with(value = tb_OC,
                                                       location = ph_location(left = 2.2, top = 1.1)) %>%
                                               ph_with(value = note_OC,
                                                       location = ph_location(left = 2.2, top = 6.5, width = 6, height = 1, label = "note_OC")) %>%
                                               ph_with(value = note_GMSC4,
                                                       location = ph_location(left = 2.2, top = 6.7, width = 6, height = 1, label = "note_GMSC4")) %>%
                                     
                                     add_slide(layout = "Title and Content", master = "Office Theme") %>%
                                               ph_with(external_img(src = leftCornerImage),
                                                       location = ph_location(left = 0, top = 0, width = 2.2, height = 1.5) )%>%
                                               ph_with(external_img(src = logoImage),
                                                       location = ph_location(left = 7.8, top = 6, width = 2.2, height = 1.5) ) %>%
                                               ph_with(value = caption_Eurotec,
                                                       location = ph_location(top = 0.1, left = 3.1, width = 10, height = 1) ) %>%
                                               ph_with(value = tb_Eurotec,
                                                       location = ph_location(left = 2.2, top = 1.1) )%>%
                                               ph_with(value = note_Eurotec,
                                                       location = ph_location(left = 2.2, top = 6.5, width = 6, height = 1, label = "note_Eurotec")) %>%

                                     add_slide(layout = "Title and Content", master = "Office Theme") %>%
                                               ph_with(external_img(src = leftCornerImage),
                                                       location = ph_location(left = 0, top = 0, width = 2.2, height = 1.5) ) %>%
                                               ph_with(external_img(src = logoImage),
                                                       location = ph_location(left = 7.8, top = 6, width = 2.2, height = 1.5) ) %>%
                                               ph_with(value = dml(Trend_Chart_all_com()),
                                                       location = ph_location(left = 0.2, top = 0.2, width = 9.5, height = 7) )%>%

                                     add_slide(layout = "Title and Content", master = "Office Theme") %>%
                                               ph_with(external_img(src = leftCornerImage),
                                                       location = ph_location(left = 0, top = 0, width = 2.2, height = 1.5) ) %>%
                                               ph_with(external_img(src = logoImage),
                                                       location = ph_location(left = 7.8, top = 6, width = 2.2, height = 1.5) ) %>%
                                               ph_with(value = dml(Trend_Chart_all()),
                                                       location = ph_location(left = 0.2, top = 0.2, width = 9.5, height = 7) )%>%
                                               ph_with(value = target,
                                                       location = ph_location(left = 2.3, top = 0.45, width = 8, height = 1, label = "target")) %>%

                                     add_slide(layout = "Title and Content", master = "Office Theme") %>%
                                               ph_with(external_img(src = leftCornerImage),
                                                       location = ph_location(left = 0, top = 0, width = 2.2, height = 1.5) ) %>%
                                               ph_with(external_img(src = logoImage),
                                                       location = ph_location(left = 7.8, top = 6, width = 2.2, height = 1.5) ) %>%
                                               ph_with(value = dml(Trend_Chart_ALL()),
                                                       location = ph_location(left = 0.2, top = 0.2, width = 9.5, height = 7) )%>%

                                     add_slide(layout = "Title and Content", master = "Office Theme") %>%
                                               ph_with(external_img(src = leftCornerImage),
                                                       location = ph_location(left = 0, top = 0, width = 2.2, height = 1.5) ) %>%
                                               ph_with(external_img(src = logoImage),
                                                       location = ph_location(left = 7.8, top = 6, width = 2.2, height = 1.5) ) %>%
                                               ph_with(value = dml(Trend_Chart_NIC()),
                                                       location = ph_location(left = 0.23, top = 0.23, width = 9.5, height = 7)) %>%

                                     add_slide(layout = "Title and Content", master = "Office Theme") %>%
                                               ph_with(external_img(src = leftCornerImage),
                                                       location = ph_location(left = 0, top = 0, width = 2.2, height = 1.5) )%>%
                                               ph_with(external_img(src = logoImage),
                                                       location = ph_location(left = 7.8, top = 6, width = 2.2, height = 1.5) )%>%
                                               ph_with(value = dml(Trend_Chart_IC()),
                                                       location = ph_location(left = 0.23, top = 0.23, width = 9.5, height = 7) )%>%

                                     add_slide(layout = "Title and Content", master = "Office Theme") %>%
                                               ph_with(external_img(src = leftCornerImage),
                                                       location = ph_location(left = 0, top = 0, width = 2.2, height = 1.5) )%>%
                                               ph_with(external_img(src = logoImage),
                                                       location = ph_location(left = 7.8, top = 6, width = 2.2, height = 1.5) )%>%
                                               ph_with(value = dml(Trend_Chart_AWC()),
                                                       location = ph_location(left = 0.23, top = 0.23, width = 9.5, height = 7) )%>%

                                     add_slide(layout = "Title and Content", master = "Office Theme") %>%
                                               ph_with(external_img(src = leftCornerImage),
                                                       location = ph_location(left = 0, top = 0, width = 2.2, height = 1.5) )%>%
                                               ph_with(external_img(src = logoImage),
                                                       location = ph_location(left = 7.8, top = 6, width = 2.2, height = 1.5) )%>%
                                               ph_with(value = dml(Trend_Chart_CC()),
                                                       location = ph_location(left = 0.23, top = 0.23, width = 9.5, height = 7) )%>%

                                     add_slide(layout = "Title and Content", master = "Office Theme") %>%
                                               ph_with(external_img(src = leftCornerImage),
                                                       location = ph_location(left = 0, top = 0, width = 2.2, height = 1.5) )%>%
                                               ph_with(external_img(src = logoImage),
                                                       location = ph_location(left = 7.8, top = 6, width = 2.2, height = 1.5) )%>%
                                               ph_with(value = dml(Trend_Chart_C()),
                                                       location = ph_location(left = 0.23, top = 0.23, width = 9.5, height = 7)) %>%
                                     
                                     add_slide(layout = "Title and Content", master = "Office Theme") %>%
                                               ph_with(external_img(src = leftCornerImage),
                                                       location = ph_location(left = 0, top = 0, width = 2.2, height = 1.5) ) %>%
                                               ph_with(external_img(src = logoImage),
                                                       location = ph_location(left = 7.8, top = 6, width = 2.2, height = 1.5) ) %>%
                                               ph_with(value = dml(Trend_Chart_OC()),
                                                       location = ph_location(left = 0.23, top = 0.23, width = 9.5, height = 7)) %>%
                                     
                                     add_slide(layout = "Title and Content", master = "Office Theme") %>%
                                               ph_with(external_img(src = leftCornerImage),
                                                       location = ph_location(left = 0, top = 0, width = 2.2, height = 1.5) ) %>%
                                               ph_with(external_img(src = logoImage),
                                                       location = ph_location(left = 7.8, top = 6, width = 2.2, height = 1.5) ) %>%
                                               ph_with(value = dml(Trend_Chart_Eurotec()),
                                                       location = ph_location(left = 0.23, top = 0.23, width = 9.5, height = 7)) %>%

                                     add_slide(layout = "Title Slide", master = "Office Theme") %>%
                                              ph_with(external_img(src = titlepageImage),
                                                      location = ph_location(left = 0, top = 0, width = 10, height = 7.5) )%>%
                                              ph_with(value = title_Region,
                                                      location = ph_location(top = 0.5, left = 0.2, width = 10, height = 1) ) %>%
                                              ph_with(value = title2,
                                                      location = ph_location(top = 1.5, left = 0.2, width = 10, height = 1) ) %>%

                                     add_slide(layout = "Title and Content", master = "Office Theme") %>%
                                               ph_with(external_img(src = leftCornerImage),
                                                       location = ph_location(left = 0, top = 0, width = 2.2, height = 1.5) )%>%
                                               ph_with(external_img(src = logoImage),
                                                       location = ph_location(left = 7.8, top = 6, width = 2.2, height = 1.5) ) %>%
                                               ph_with(value = caption_EUR,
                                                       location = ph_location(top = 0.1, left = 3.8, width = 10, height = 1)) %>%
                                               ph_with(value = tb_EUR,
                                                       location = ph_location(left = 2.2, top = 1.3) )%>%
                                               ph_with(value = note_Region,
                                                       location = ph_location(left = 2, top = 6.6, width = 6, height = 1, label = "note_Region")) %>%
                                               ph_with(value = note_EUR,
                                                       location = ph_location(left = 3.1, top = 0.5, width = 6, height = 1, label = "note_EUR")) %>%
                                               ph_with(value = reduction_EUR,
                                                       location = ph_location(left = 2, top = 6.8, width = 6, height = 1, label = "reduction_EUR")) %>%

                                     add_slide(layout = "Title and Content", master = "Office Theme") %>%
                                               ph_with(external_img(src = leftCornerImage),
                                                       location = ph_location(left = 0, top = 0, width = 2.2, height = 1.5) ) %>%
                                               ph_with(external_img(src = logoImage),
                                                       location = ph_location(left = 7.8, top = 6, width = 2.2, height = 1.5) ) %>%
                                               ph_with(value = caption_NAM,
                                                       location = ph_location(top = 0.1, left = 3.8, width = 10, height = 1)) %>%
                                               ph_with(value = tb_NAM,
                                                       location = ph_location(left = 2.2, top = 1.3)) %>%
                                               ph_with(value = note_Region,
                                                       location = ph_location(left = 2, top = 6.6, width = 6, height = 1, label = "note_Region")) %>%
                                               ph_with(value = note_NAM,
                                                       location = ph_location(left = 3, top = 0.5, width = 6, height = 1, label = "note_NAM")) %>%
                                               ph_with(value = reduction_NAM,
                                                       location = ph_location(left = 2, top = 6.8, width = 6, height = 1, label = "reduction_NAM")) %>%

                                     add_slide(layout = "Title and Content", master = "Office Theme") %>%
                                               ph_with(external_img(src = leftCornerImage),
                                                       location = ph_location(left = 0, top = 0, width = 2.2, height = 1.5) )%>%
                                               ph_with(external_img(src = logoImage),
                                                       location = ph_location(left = 7.8, top = 6, width = 2.2, height = 1.5) ) %>%
                                               ph_with(value = caption_GEM,
                                                       location = ph_location(top = 0.1, left = 3.7, width = 10, height = 1)) %>%
                                               ph_with(value = tb_GEM,
                                                       location = ph_location(left = 2.2, top = 1.3)) %>%
                                               ph_with(value = note_Region,
                                                       location = ph_location(left = 2, top = 6.6, width = 6, height = 1, label = "note_Region")) %>%
                                               ph_with(value = note_GEM,
                                                       location = ph_location(left = 2.6, top = 0.5, width = 6, height = 1, label = "note_GEM")) %>%
                                               ph_with(value = reduction_GEM,
                                                       location = ph_location(left = 2, top = 6.8, width = 6, height = 1, label = "reduction_GEM")) %>%

                                     add_slide(layout = "Title and Content", master = "Office Theme") %>%
                                               ph_with(external_img(src = leftCornerImage),
                                                       location = ph_location(left = 0, top = 0, width = 2.2, height = 1.5) )%>%
                                               ph_with(external_img(src = logoImage),
                                                       location = ph_location(left = 7.8, top = 6, width = 2.2, height = 1.5) ) %>%
                                               ph_with(value = caption_AP,
                                                       location = ph_location(top = 0.1, left = 3, width = 10, height = 1)) %>%
                                               ph_with(value = tb_AP,
                                                       location = ph_location(left = 2.2, top = 1.3) )%>%
                                               ph_with(value = note_Region,
                                                       location = ph_location(left = 2, top = 6.6, width = 6, height = 1, label = "note_Region")) %>%
                                               ph_with(value = reduction_AP,
                                                       location = ph_location(left = 2, top = 6.8, width = 6, height = 1, label = "reduction_AP")) %>%

                                     add_slide(layout = "Title and Content", master = "Office Theme") %>%
                                               ph_with(external_img(src = leftCornerImage),
                                                       location = ph_location(left = 0, top = 0, width = 2.2, height = 1.5) )%>%
                                               ph_with(external_img(src = logoImage),
                                                       location = ph_location(left = 7.8, top = 6, width = 2.2, height = 1.5) ) %>%
                                               ph_with(value = caption_LA,
                                                       location = ph_location(top = 0.1, left = 3, width = 10, height = 1)) %>%
                                               ph_with(value = tb_LA,
                                                       location = ph_location(left = 2.2, top = 1.3)) %>%
                                               ph_with(value = note_Region,
                                                       location = ph_location(left = 2, top = 6.6, width = 6, height = 1, label = "note_Region")) %>%
                                               ph_with(value = reduction_LA,
                                                       location = ph_location(left = 2, top = 6.8, width = 6, height = 1, label = "reduction_LA")) %>%

                                     add_slide(layout = "Title and Content", master = "Office Theme") %>%
                                               ph_with(external_img(src = leftCornerImage),
                                                       location = ph_location(left = 0, top = 0, width = 2.2, height = 1.5) )%>%
                                               ph_with(external_img(src = logoImage),
                                                       location = ph_location(left = 7.8, top = 6, width = 2.2, height = 1.5) ) %>%
                                               ph_with(value = caption_MEA,
                                                       location = ph_location(top = 0.1, left = 3, width = 10, height = 1)) %>%
                                               ph_with(value = tb_MEA,
                                                       location = ph_location(left = 2.2, top = 1.3) )%>%
                                               ph_with(value = note_Region,
                                                       location = ph_location(left = 2, top = 6.6, width = 6, height = 1, label = "note_Region")) %>%
                                               ph_with(value = reduction_MEA,
                                                       location = ph_location(left = 2, top = 6.8, width = 6, height = 1, label = "reduction_MEA")) %>%

                                     add_slide(layout = "Title and Content", master = "Office Theme") %>%
                                               ph_with(external_img(src = leftCornerImage),
                                                       location = ph_location(left = 0, top = 0, width = 2.2, height = 1.5) )%>%
                                               ph_with(external_img(src = logoImage),
                                                       location = ph_location(left = 7.8, top = 6, width = 2.2, height = 1.5) )%>%
                                               ph_with(dml(Trend_Chart_EUR()),
                                                       location = ph_location(left = 0.23, top = 0.23, width = 9.5, height = 7) )%>%

                                     add_slide(layout = "Title and Content", master = "Office Theme") %>%
                                               ph_with(external_img(src = leftCornerImage),
                                                       location = ph_location(left = 0, top = 0, width = 2.2, height = 1.5) ) %>%
                                               ph_with(external_img(src = logoImage),
                                                       location = ph_location(left = 7.8, top = 6, width = 2.2, height = 1.5) ) %>%
                                               ph_with(dml(Trend_Chart_NAM()),
                                                       location = ph_location(left = 0.23, top = 0.23, width = 9.5, height = 7)) %>%

                                     add_slide(layout = "Title and Content", master = "Office Theme") %>%
                                               ph_with(external_img(src = leftCornerImage),
                                                       location = ph_location(left = 0, top = 0, width = 2.2, height = 1.5) ) %>%
                                               ph_with(external_img(src = logoImage),
                                                       location = ph_location(left = 7.8, top = 6, width = 2.2, height = 1.5) ) %>%
                                               ph_with(dml(Trend_Chart_GEM()),
                                                       location = ph_location(left = 0.23, top = 0.23, width = 9.5, height = 7)) %>%

                                     add_slide(layout = "Title and Content", master = "Office Theme") %>%
                                               ph_with(external_img(src = leftCornerImage),
                                                       location = ph_location(left = 0, top = 0, width = 2.2, height = 1.5) )%>%
                                               ph_with(external_img(src = logoImage),
                                                       location = ph_location(left = 7.8, top = 6, width = 2.2, height = 1.5) )%>%
                                               ph_with(dml(Trend_Chart_AP()),
                                                       location = ph_location(left = 0.23, top = 0.23, width = 9.5, height = 7) )%>%

                                     add_slide(layout = "Title and Content", master = "Office Theme") %>%
                                               ph_with(external_img(src = leftCornerImage),
                                                       location = ph_location(left = 0, top = 0, width = 2.2, height = 1.5) ) %>%
                                               ph_with(external_img(src = logoImage),
                                                       location = ph_location(left = 7.8, top = 6, width = 2.2, height = 1.5) ) %>%
                                               ph_with(dml(Trend_Chart_LA()),
                                                       location = ph_location(left = 0.23, top = 0.23, width = 9.5, height = 7)) %>%

                                     add_slide(layout = "Title and Content", master = "Office Theme") %>%
                                               ph_with(external_img(src = leftCornerImage),
                                                       location = ph_location(left = 0, top = 0, width = 2.2, height = 1.5) ) %>%
                                               ph_with(external_img(src = logoImage),
                                                       location = ph_location(left = 7.8, top = 6, width = 2.2, height = 1.5) ) %>%
                                               ph_with(dml(Trend_Chart_MEA()),
                                                       location = ph_location(left = 0.23, top = 0.23, width = 9.5, height = 7)) %>%

                                     
                                   print(example_pp, target = file)
                                   
                                   })

  output$slides_r <- downloadHandler(filename = function() {paste("CPM by Region Report ", format(input$YTD[2],"%B %Y"), ".pptx", sep = "")},
                                     content = function(file) {
                                     titlepageImage <- paste(input$images$datapath[input$images$name=="title_page.PNG"])
                                     leftCornerImage <- paste(input$images$datapath[input$images$name=="left_corner.PNG"])
                                     logoImage <- paste(input$images$datapath[input$images$name=="logo.PNG"])
                                     GEM <- GEM()
                                     EUR <- EUR()
                                     NAM <- NAM()
                                     AP <- AP()
                                     LA <- LA()
                                     MEA <- MEA()

                                     std_border = fp_border(color="black")
                                     pastyear2 <- format((input$YTD[1] - years(2)), "%Y")
                                     pastyear1 <- format((input$YTD[1] - years(1)), "%Y")
                                     currentyear <- format(input$YTD[1], "%Y")


                                     title1 <- block_list(
                                       fpar(ftext("CPM by Region", shortcuts$fp_bold(font.size = 36, color = "white")))
                                     )

                                     title2 <- block_list(
                                       fpar(ftext(paste("Report Date:", input$report), shortcuts$fp_bold(font.size = 18, color = "white"))),
                                       fpar(ftext(paste(input$xx, "on behalf of Metrics Central"), shortcuts$fp_bold(font.size = 18, color = "white")))
                                     )

                                     caption1 <- block_list(
                                       fpar(ftext("CPM - GEM", shortcuts$fp_bold(font.size = 22, color = "dodgerblue4")))
                                     )

                                     caption2 <- block_list(
                                       fpar(ftext("CPM - EUR", shortcuts$fp_bold(font.size = 22, color = "darkorange1")))
                                     )

                                     caption3 <- block_list(
                                       fpar(ftext("CPM - NAM", shortcuts$fp_bold(font.size = 22, color = "aquamarine4")))
                                     )

                                     caption4 <- block_list(
                                       fpar(ftext("CPM - GEM Breakdown (APAC)", shortcuts$fp_bold(font.size = 22, color = "cornflowerblue")))
                                     )

                                     caption5 <- block_list(
                                       fpar(ftext("CPM - GEM Breakdown (LATAM)", shortcuts$fp_bold(font.size = 22, color = "cornflowerblue")))
                                     )

                                     caption6 <- block_list(
                                       fpar(ftext("CPM - GEM Breakdown (MEA)", shortcuts$fp_bold(font.size = 22, color = "cornflowerblue")))
                                     )

                                     note <- block_list(
                                       fpar(ftext(paste0("*Note: Exclude Infusion Care and Eurotec"), shortcuts$fp_italic(font.size = 10, color="red")))
                                     )

                                     note_GEM <- block_list(
                                       fpar(ftext(paste0("GEM - Global Emerging Market (APAC+LATAM+MEA+TURKEY)"), shortcuts$fp_italic(font.size = 10, color="black")))
                                     )

                                     note_EUR <- block_list(
                                       fpar(ftext(paste0("EUR: European Countries (excluding Turkey)"), shortcuts$fp_italic(font.size = 10, color="black")))
                                     )

                                     note_NAM <- block_list(
                                       fpar(ftext(paste0("NAM: North America (US +PUERTO RICO+ CANADA)"), shortcuts$fp_italic(font.size = 10, color="black")))
                                     )

                                     reduction_GEM <- block_list(
                                       fpar(ftext(paste0("*Target: Assuming ", (1-as.numeric(reduction_region_GEM()))*100, "% complaints reduction from previous year."), shortcuts$fp_italic(color="black")))
                                     )

                                     reduction_EUR <- block_list(
                                       fpar(ftext(paste0("*Target: Assuming ", (1-as.numeric(input$reduction_EUR))*100, "% complaints reduction from previous year."), shortcuts$fp_italic(color="black")))
                                     )

                                     reduction_NAM <- block_list(
                                       fpar(ftext(paste0("*Target: Assuming ", (1-as.numeric(input$reduction_NAM))*100, "% complaints reduction from previous year."), shortcuts$fp_italic(color="black")))
                                     )

                                     reduction_AP <- block_list(
                                       fpar(ftext(paste0("*Target: Assuming ", (1-as.numeric(input$reduction_AP))*100, "% complaints reduction from previous year."), shortcuts$fp_italic(color="black")))
                                     )

                                     reduction_LA <- block_list(
                                       fpar(ftext(paste0("*Target: Assuming ", (1-as.numeric(input$reduction_LA))*100, "% complaints reduction from previous year."), shortcuts$fp_italic(color="black")))
                                     )

                                     reduction_MEA <- block_list(
                                       fpar(ftext(paste0("*Target: Assuming ", (1-as.numeric(input$reduction_MEA))*100, "% complaints reduction from previous year."), shortcuts$fp_italic(color="black")))
                                     )



                                     tb_GEM <- flextable(GEM) %>%
                                       set_header_labels(na = "Month",
                                                         na1 = "Month",
                                                         #pastyear2 = paste0(pastyear2, "-Overall"),
                                                         pastyear1 = paste0(pastyear1, "-Overall"),
                                                         currentyear_target = paste0(currentyear, "-Target*"),
                                                         currentyear_actual = paste0(currentyear, "-Actual")) %>%

                                       merge_h(part = "header") %>%
                                       merge_h_range(i = 1:13, j1 = 1, j2 = 2, part = "body") %>%
                                       merge_at(i = 1:13, j = 3, part = "body") %>%
                                       merge_at(i = 1:13, j = 4, part = "body") %>%

                                       bold(part = "header") %>%
                                       bold(i = 13, part = "body") %>%
                                       fontsize(size = 14, part = "all") %>%
                                       fontsize(i = 13, size = 15, part = "body") %>%
                                       bg(bg = "dodgerblue4", part = "header") %>%
                                       color(color = "white", part = "header") %>%
                                       width(width = 1.1) %>%
                                       height(height = 0.45, part = "header") %>%
                                       height(height = 0.40, part = "body") %>%
                                       align(align = "center", part = "all")%>%
                                       border_outer(border = std_border, part = "all") %>%
                                       border_inner(border = std_border, part = "all")


                                     tb_EUR <- flextable(EUR) %>%
                                       set_header_labels(na = "Month",
                                                         na1 = "Month",
                                                         #pastyear2 = paste0(pastyear2, "-Overall"),
                                                         pastyear1 = paste0(pastyear1, "-Overall"),
                                                         currentyear_target = paste0(currentyear, "-Target*"),
                                                         currentyear_actual = paste0(currentyear, "-Actual")) %>%

                                       merge_h(part = "header") %>%
                                       merge_h_range(i = 1:13, j1 = 1, j2 = 2, part = "body") %>%
                                       merge_at(i = 1:13, j = 3, part = "body") %>%
                                       merge_at(i = 1:13, j = 4, part = "body") %>%
                                       #merge_at(i = 1:13, j = 5, part = "body") %>%

                                       bold(part = "header") %>%
                                       bold(i = 13, part = "body") %>%
                                       fontsize(size = 14, part = "all") %>%
                                       fontsize(i = 13, size = 15, part = "body") %>%
                                       bg(bg = "darkorange1", part = "header") %>%
                                       color(color = "white", part = "header") %>%
                                       width(width = 1.1) %>%
                                       height(height = 0.45, part = "header") %>%
                                       height(height = 0.40, part = "body") %>%
                                       align(align = "center", part = "all")%>%
                                       border_outer(border = std_border, part = "all") %>%
                                       border_inner(border = std_border, part = "all")

                                     tb_NAM <- flextable(NAM) %>%
                                       set_header_labels(na = "Month",
                                                         na1 = "Month",
                                                         #pastyear2 = paste0(pastyear2, "-Overall"),
                                                         pastyear1 = paste0(pastyear1, "-Overall"),
                                                         currentyear_target = paste0(currentyear, "-Target*"),
                                                         currentyear_actual = paste0(currentyear, "-Actual")) %>%

                                       merge_h(part = "header") %>%
                                       merge_h_range(i = 1:13, j1 = 1, j2 = 2, part = "body") %>%
                                       merge_at(i = 1:13, j = 3, part = "body") %>%
                                       merge_at(i = 1:13, j = 4, part = "body") %>%
                                       #merge_at(i = 1:13, j = 5, part = "body") %>%

                                       bold(part = "header") %>%
                                       bold(i = 13, part = "body") %>%
                                       fontsize(size = 14, part = "all") %>%
                                       fontsize(i = 13, size = 15, part = "body") %>%
                                       bg(bg = "aquamarine4", part = "header") %>%
                                       color(color = "white", part = "header") %>%
                                       width(width = 1.1) %>%
                                       height(height = 0.45, part = "header") %>%
                                       height(height = 0.40, part = "body") %>%
                                       align(align = "center", part = "all")%>%
                                       border_outer(border = std_border, part = "all") %>%
                                       border_inner(border = std_border, part = "all")

                                     tb_AP <- flextable(AP) %>%
                                       set_header_labels(na = "Month",
                                                         na1 = "Month",
                                                         #pastyear2 = paste0(pastyear2, "-Overall"),
                                                         pastyear1 = paste0(pastyear1, "-Overall"),
                                                         currentyear_target = paste0(currentyear, "-Target*"),
                                                         currentyear_actual = paste0(currentyear, "-Actual")) %>%

                                       merge_h(part = "header") %>%
                                       merge_h_range(i = 1:13, j1 = 1, j2 = 2, part = "body") %>%
                                       merge_at(i = 1:13, j = 3, part = "body") %>%
                                       merge_at(i = 1:13, j = 4, part = "body") %>%
                                       #merge_at(i = 1:13, j = 5, part = "body") %>%

                                       bold(part = "header") %>%
                                       bold(i = 13, part = "body") %>%
                                       fontsize(size = 14, part = "all") %>%
                                       fontsize(i = 13, size = 15, part = "body") %>%
                                       bg(bg = "cornflowerblue", part = "header") %>%
                                       color(color = "white", part = "header") %>%
                                       width(width = 1.1) %>%
                                       height(height = 0.45, part = "header") %>%
                                       height(height = 0.40, part = "body") %>%
                                       align(align = "center", part = "all")%>%
                                       border_outer(border = std_border, part = "all") %>%
                                       border_inner(border = std_border, part = "all")

                                     tb_LA <- flextable(LA) %>%
                                       set_header_labels(na = "Month",
                                                         na1 = "Month",
                                                         #pastyear2 = paste0(pastyear2, "-Overall"),
                                                         pastyear1 = paste0(pastyear1, "-Overall"),
                                                         currentyear_target = paste0(currentyear, "-Target*"),
                                                         currentyear_actual = paste0(currentyear, "-Actual")) %>%

                                       merge_h(part = "header") %>%
                                       merge_h_range(i = 1:13, j1 = 1, j2 = 2, part = "body") %>%
                                       merge_at(i = 1:13, j = 3, part = "body") %>%
                                       merge_at(i = 1:13, j = 4, part = "body") %>%
                                       #merge_at(i = 1:13, j = 5, part = "body") %>%

                                       bold(part = "header") %>%
                                       bold(i = 13, part = "body") %>%
                                       fontsize(size = 14, part = "all") %>%
                                       fontsize(i = 13, size = 15, part = "body") %>%
                                       bg(bg = "cornflowerblue", part = "header") %>%
                                       color(color = "white", part = "header") %>%
                                       width(width = 1.1) %>%
                                       height(height = 0.45, part = "header") %>%
                                       height(height = 0.40, part = "body") %>%
                                       align(align = "center", part = "all")%>%
                                       border_outer(border = std_border, part = "all") %>%
                                       border_inner(border = std_border, part = "all")

                                     tb_MEA <- flextable(MEA) %>%
                                       set_header_labels(na = "Month",
                                                         na1 = "Month",
                                                         #pastyear2 = paste0(pastyear2, "-Overall"),
                                                         pastyear1 = paste0(pastyear1, "-Overall"),
                                                         currentyear_target = paste0(currentyear, "-Target*"),
                                                         currentyear_actual = paste0(currentyear, "-Actual")) %>%

                                       merge_h(part = "header") %>%
                                       merge_h_range(i = 1:13, j1 = 1, j2 = 2, part = "body") %>%
                                       merge_at(i = 1:13, j = 3, part = "body") %>%
                                       merge_at(i = 1:13, j = 4, part = "body") %>%

                                       bold(part = "header") %>%
                                       bold(i = 13, part = "body") %>%
                                       fontsize(size = 14, part = "all") %>%
                                       fontsize(i = 13, size = 15, part = "body") %>%
                                       bg(bg = "cornflowerblue", part = "header") %>%
                                       color(color = "white", part = "header") %>%
                                       width(width = 1.1) %>%
                                       height(height = 0.45, part = "header") %>%
                                       height(height = 0.40, part = "body") %>%
                                       align(align = "center", part = "all")%>%
                                       border_outer(border = std_border, part = "all") %>%
                                       border_inner(border = std_border, part = "all")


                                     example_pp <- read_pptx() %>%

                                       add_slide(layout = "Title Slide", master = "Office Theme") %>%
                                       ph_with(external_img(src = titlepageImage),
                                               location = ph_location(left = 0, top = 0, width = 10, height = 7.5) )%>%
                                       ph_with(value = title1,
                                               location = ph_location(top = 0.5, left = 0.2, width = 10, height = 1)) %>%
                                       ph_with(value = title2,
                                               location = ph_location(top = 1.5, left = 0.2, width = 10, height = 1) ) %>%


                                       add_slide(layout = "Title and Content", master = "Office Theme") %>%
                                       ph_with(external_img(src = leftCornerImage),
                                               location = ph_location(left = 0, top = 0, width = 2.2, height = 1.5) )%>%
                                       ph_with(external_img(src = logoImage),
                                               location = ph_location(left = 7.8, top = 6, width = 2.2, height = 1.5) ) %>%
                                       ph_with(value = caption2,
                                               location = ph_location(top = 0.1, left = 3.8, width = 10, height = 1)) %>%
                                       ph_with(value = tb_EUR,
                                               location = ph_location(left = 2.2, top = 1.3) )%>%
                                       ph_with(value = note,
                                               location = ph_location(left = 2, top = 6.6, width = 6, height = 1, label = "note")) %>%
                                       ph_with(value = note_EUR,
                                               location = ph_location(left = 3.1, top = 0.5, width = 6, height = 1, label = "note_EUR")) %>%
                                       ph_with(value = reduction_EUR,
                                               location = ph_location(left = 2, top = 6.8, width = 6, height = 1, label = "reduction_EUR")) %>%

                                       add_slide(layout = "Title and Content", master = "Office Theme") %>%
                                       ph_with(external_img(src = leftCornerImage),
                                               location = ph_location(left = 0, top = 0, width = 2.2, height = 1.5) ) %>%
                                       ph_with(external_img(src = logoImage),
                                               location = ph_location(left = 7.8, top = 6, width = 2.2, height = 1.5) ) %>%
                                       ph_with(value = caption3,
                                               location = ph_location(top = 0.1, left = 3.8, width = 10, height = 1)) %>%
                                       ph_with(value = tb_NAM,
                                               location = ph_location(left = 2.2, top = 1.3)) %>%
                                       ph_with(value = note,
                                               location = ph_location(left = 2, top = 6.6, width = 6, height = 1, label = "note")) %>%
                                       ph_with(value = note_NAM,
                                               location = ph_location(left = 3, top = 0.5, width = 6, height = 1, label = "note_NAM")) %>%
                                       ph_with(value = reduction_NAM,
                                               location = ph_location(left = 2, top = 6.8, width = 6, height = 1, label = "reduction_NAM")) %>%

                                       add_slide(layout = "Title and Content", master = "Office Theme") %>%
                                       ph_with(external_img(src = leftCornerImage),
                                               location = ph_location(left = 0, top = 0, width = 2.2, height = 1.5) )%>%
                                       ph_with(external_img(src = logoImage),
                                               location = ph_location(left = 7.8, top = 6, width = 2.2, height = 1.5) ) %>%
                                       ph_with(value = caption1,
                                               location = ph_location(top = 0.1, left = 3.7, width = 10, height = 1)) %>%
                                       ph_with(value = tb_GEM,
                                               location = ph_location(left = 2.2, top = 1.3)) %>%
                                       ph_with(value = note,
                                               location = ph_location(left = 2, top = 6.6, width = 6, height = 1, label = "note")) %>%
                                       ph_with(value = note_GEM,
                                               location = ph_location(left = 2.6, top = 0.5, width = 6, height = 1, label = "note_GEM")) %>%
                                       ph_with(value = reduction_GEM,
                                               location = ph_location(left = 2, top = 6.8, width = 6, height = 1, label = "reduction_GEM")) %>%

                                       add_slide(layout = "Title and Content", master = "Office Theme") %>%
                                       ph_with(external_img(src = leftCornerImage),
                                               location = ph_location(left = 0, top = 0, width = 2.2, height = 1.5) )%>%
                                       ph_with(external_img(src = logoImage),
                                               location = ph_location(left = 7.8, top = 6, width = 2.2, height = 1.5) ) %>%
                                       ph_with(value = caption4,
                                               location = ph_location(top = 0.1, left = 3, width = 10, height = 1)) %>%
                                       ph_with(value = tb_AP,
                                               location = ph_location(left = 2.2, top = 1.3) )%>%
                                       ph_with(value = note,
                                               location = ph_location(left = 2, top = 6.6, width = 6, height = 1, label = "note")) %>%
                                       ph_with(value = reduction_AP,
                                               location = ph_location(left = 2, top = 6.8, width = 6, height = 1, label = "reduction_AP")) %>%

                                       add_slide(layout = "Title and Content", master = "Office Theme") %>%
                                       ph_with(external_img(src = leftCornerImage),
                                               location = ph_location(left = 0, top = 0, width = 2.2, height = 1.5) )%>%
                                       ph_with(external_img(src = logoImage),
                                               location = ph_location(left = 7.8, top = 6, width = 2.2, height = 1.5) ) %>%
                                       ph_with(value = caption5,
                                               location = ph_location(top = 0.1, left = 3, width = 10, height = 1)) %>%
                                       ph_with(value = tb_LA,
                                               location = ph_location(left = 2.2, top = 1.3)) %>%
                                       ph_with(value = note,
                                               location = ph_location(left = 2, top = 6.6, width = 6, height = 1, label = "note")) %>%
                                       ph_with(value = reduction_LA,
                                               location = ph_location(left = 2, top = 6.8, width = 6, height = 1, label = "reduction_LA")) %>%

                                       add_slide(layout = "Title and Content", master = "Office Theme") %>%
                                       ph_with(external_img(src = leftCornerImage),
                                               location = ph_location(left = 0, top = 0, width = 2.2, height = 1.5) )%>%
                                       ph_with(external_img(src = logoImage),
                                               location = ph_location(left = 7.8, top = 6, width = 2.2, height = 1.5) ) %>%
                                       ph_with(value = caption6,
                                               location = ph_location(top = 0.1, left = 3, width = 10, height = 1)) %>%
                                       ph_with(value = tb_MEA,
                                               location = ph_location(left = 2.2, top = 1.3) )%>%
                                       ph_with(value = note,
                                               location = ph_location(left = 2, top = 6.6, width = 6, height = 1, label = "note")) %>%
                                       ph_with(value = reduction_MEA,
                                               location = ph_location(left = 2, top = 6.8, width = 6, height = 1, label = "reduction_MEA")) %>%

                                       add_slide(layout = "Title and Content", master = "Office Theme") %>%
                                       ph_with(external_img(src = leftCornerImage),
                                               location = ph_location(left = 0, top = 0, width = 2.2, height = 1.5) )%>%
                                       ph_with(external_img(src = logoImage),
                                               location = ph_location(left = 7.8, top = 6, width = 2.2, height = 1.5) )%>%
                                       ph_with(dml(Trend_Chart_EUR()),
                                               location = ph_location(left = 0.23, top = 0.23, width = 9.5, height = 7) )%>%

                                       add_slide(layout = "Title and Content", master = "Office Theme") %>%
                                       ph_with(external_img(src = leftCornerImage),
                                               location = ph_location(left = 0, top = 0, width = 2.2, height = 1.5) ) %>%
                                       ph_with(external_img(src = logoImage),
                                               location = ph_location(left = 7.8, top = 6, width = 2.2, height = 1.5) ) %>%
                                       ph_with(dml(Trend_Chart_NAM()),
                                               location = ph_location(left = 0.23, top = 0.23, width = 9.5, height = 7)) %>%

                                       add_slide(layout = "Title and Content", master = "Office Theme") %>%
                                       ph_with(external_img(src = leftCornerImage),
                                               location = ph_location(left = 0, top = 0, width = 2.2, height = 1.5) ) %>%
                                       ph_with(external_img(src = logoImage),
                                               location = ph_location(left = 7.8, top = 6, width = 2.2, height = 1.5) ) %>%
                                       ph_with(dml(Trend_Chart_GEM()),
                                                location = ph_location(left = 0.23, top = 0.23, width = 9.5, height = 7)) %>%

                                       add_slide(layout = "Title and Content", master = "Office Theme") %>%
                                       ph_with(external_img(src = leftCornerImage),
                                               location = ph_location(left = 0, top = 0, width = 2.2, height = 1.5) )%>%
                                       ph_with(external_img(src = logoImage),
                                               location = ph_location(left = 7.8, top = 6, width = 2.2, height = 1.5) )%>%
                                       ph_with(dml(Trend_Chart_AP()),
                                               location = ph_location(left = 0.23, top = 0.23, width = 9.5, height = 7) )%>%

                                       add_slide(layout = "Title and Content", master = "Office Theme") %>%
                                       ph_with(external_img(src = leftCornerImage),
                                               location = ph_location(left = 0, top = 0, width = 2.2, height = 1.5) ) %>%
                                       ph_with(external_img(src = logoImage),
                                               location = ph_location(left = 7.8, top = 6, width = 2.2, height = 1.5) ) %>%
                                       ph_with(dml(Trend_Chart_LA()),
                                               location = ph_location(left = 0.23, top = 0.23, width = 9.5, height = 7)) %>%

                                       add_slide(layout = "Title and Content", master = "Office Theme") %>%
                                       ph_with(external_img(src = leftCornerImage),
                                               location = ph_location(left = 0, top = 0, width = 2.2, height = 1.5) ) %>%
                                       ph_with(external_img(src = logoImage),
                                               location = ph_location(left = 7.8, top = 6, width = 2.2, height = 1.5) ) %>%
                                       ph_with(dml(Trend_Chart_MEA()),
                                               location = ph_location(left = 0.23, top = 0.23, width = 9.5, height = 7)) %>%

                                       print(example_pp, target = file)

                                   })


  output$file <- downloadHandler(filename = function() {paste("Complaints and CPM by BU ", format(input$YTD[2],"%B %Y"), ".xlsx",sep = "")},
                                    content = function(file) {wb <- createWorkbook()
                                    addWorksheet(wb = wb, sheet = "CPM_monthly_by_BU" )
                                    writeData(wb = wb, sheet = 1, x =  CPM_monthly())
                                    addWorksheet(wb = wb, sheet = "Sales_monthly_by_BU" )
                                    writeData(wb = wb, sheet = 2, x =  sales_download())
                                    addWorksheet(wb = wb, sheet = "CPM_yearly_by_BU" )
                                    writeData(wb = wb, sheet = 3, x =  CPM_yearly())
                                    addWorksheet(wb = wb, sheet = "CPM_monthly_NIC" )
                                    writeData(wb = wb, sheet = 4, x =  CPM_monthly_NIC())
                                    addWorksheet(wb = wb, sheet = "CPM_yearly_NIC" )
                                    writeData(wb = wb, sheet = 5, x =  CPM_yearly_NIC())
                                    addWorksheet(wb = wb, sheet = "CPM_monthly_ALL_BU" )
                                    writeData(wb = wb, sheet = 6, x =  CPM_monthly_all())
                                    addWorksheet(wb = wb, sheet = "CPM_yearly_ALL_BU" )
                                    writeData(wb = wb, sheet = 7, x =  CPM_yearly_all())
                                    withProgress(message = paste0("Downloading Complaints & CPM File"),
                                                 value = 0, 
                                                 {for (i in 1:10) {incProgress(1/10, detail = paste(i*10, "%"))
                                                   Sys.sleep(0.1)}
                                                 saveWorkbook(wb, file, overwrite = T)
                                                 })
                                    })
  
  output$file_r <- downloadHandler(filename = function() {paste("Complaints and CPM by Region ", format(input$YTD[2],"%B %Y"), ".xlsx",sep = "")},
                                   content = function(file) {wb <- createWorkbook()
                                   addWorksheet(wb = wb, sheet = "CPM_Monthly_by_Region" )
                                   writeData(wb = wb, sheet = 1, x =  CPM_region_monthly())
                                   addWorksheet(wb = wb, sheet = "Sales_Monthly_by_Region" )
                                   writeData(wb = wb, sheet = 2, x =  sales_region_download())
                                   addWorksheet(wb = wb, sheet = "CPM_Yearly_by_Region" )
                                   writeData(wb = wb, sheet = 3, x =  CPM_region_yearly())
                                   addWorksheet(wb = wb, sheet = "CPM_Monthly_by_GEM" )
                                   writeData(wb = wb, sheet = 4, x =  CPM_region_monthly2())
                                   addWorksheet(wb = wb, sheet = "Sales_Monthly_by_GEM" )
                                   writeData(wb = wb, sheet = 5, x =  sales_region2_download())
                                   addWorksheet(wb = wb, sheet = "CPM_Yearly_by_GEM" )
                                   writeData(wb = wb, sheet = 6, x =  CPM_region_yearly2())
                                   withProgress(message = paste0("Downloading Complaints & CPM File"),
                                                value = 0, 
                                                {for (i in 1:10) {incProgress(1/10, detail = paste(i*10, "%"))
                                                  Sys.sleep(0.1)}
                                                
                                                  saveWorkbook(wb, file, overwrite = T)
                                              })
                                 })
  
  end_date <-  reactive(paste(format(input$YTD[2],'%m'), format(input$YTD[2],'%d'), format(input$YTD[2],'%y'), sep=""))

  
  output$TW_complaint <- downloadHandler(filename = function() {paste("TW8.7 January 2017 - ", format(input$YTD[2],"%B %Y"), ".xlsx",sep = "")},
                                         content = function(file) {
                                     
                                         withProgress(message = paste0("Downloading Complaints & CPM File"),
                                                      value = 0, 
                                                     {for (i in 1:10) {incProgress(1/10, detail = paste(i*10, "%"))
                                                      Sys.sleep(0.1)}
                                                      })
                                         write.xlsx(complaint0(),file)
                                         })
                                      
}



# Run the application 

shinyApp(ui = ui, server = server)