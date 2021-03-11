#===========================================================================================================
#
#                                 MRL Dashboard
#                                 
#                                 Last Update:  July 2020
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
#install.packages("zoo") #function 'as.yearmon'
#install.packages("mschart") #enable to create native office charts that can be used with officer.
#install.packages("ggplot2")
#install.packages("ggrepel")
#install.packages("gridExtra") 
#install.packages("rvg") #providing editable vector graphics that can be used with officer
#install.packages("flextable")
#install.packages("officer")
#install.packages("writexl")
#install.packages("readxl")
#install.packages("dplyr")
#install.packages("formattable")
#install.packages("openxlsx")

library(shiny)
library(shinythemes)
library(officer)
library(flextable)
library(dplyr)
library(zoo)
library(base)
library(ggplot2)
library(ggrepel)
library(grid)
library(rvg)
library(writexl)
library(readxl)
library(stringi)
library(formattable) # function  "percent"
library(openxlsx)

#-----------------------------------------------------------------------------------------------------------
loc_path <- "C:/Users/QI15403/OneDrive - ConvaTec Limited/Desktop/combined MRL reporting/R Shiny/www/"

#-----------------------------------------------------------------------------------------------------------
#
#                                         Define UI 
#
#-----------------------------------------------------------------------------------------------------------

ui <- navbarPage(
  
  theme = shinytheme("cerulean"),
  
  "MRL Reporting", # App title
  
  tabsetPanel(
    
    
    tabPanel("MRL Reporting", #tab title
             
             sidebarPanel(width = 3,
                          dateRangeInput('dateRange', label = 'Date range input: yyyy-mm-dd',
                                         start = Sys.Date(), end = Sys.Date()),
                          #fileInput("base","Choose MRL Base File (.csv)",multiple = FALSE,
                          #          accept = c("text/csv","text/comma-separated-values,text/plain",".csv")),
                          fileInput("base","Choose MRL Base File (.csv)", multiple = F, accept = c(".xlsx")),
                          fileInput("current","Choose Six Files from Current Month (.xlsx)",multiple = T, accept = c(".xlsx")),
                          fileInput("old","Choose Historical Data File (.xlsx)",multiple = T, accept = c(".xlsx")),
                          textInput("received", "Received Dates"),
                          #textInput("report", "Report Date"),
                          textInput("slide2", "Trend Chart Comments"),
                          textInput("slide3", "Market Working Load Table Comments"),
                          textInput("slide4", "Cycle Time Table Comments"),
                          textInput("slide5", "Regulatory Break Down Table Comments"),
                          textInput("slide6", "Outlier Table Comments"),
                          
                          
                          downloadButton("report", "Download PPT Report"),
                          downloadButton("base", "Download Base File"),
                          downloadButton("outlier", "Download Outliers File"),
                          downloadButton("historical", "Download Historical Data")
                          
             ),
             
             mainPanel(
               p("=================================================================================================================================="),
               
               fluidRow(
                 column(6, h1("MRL Reporting"),
                        p("Copyright belong to ConvaTec Metrics Central team"),
                        p("Updated Date: Dec 2019")),
                 column(4, img(src = "ConvaTec.png",height = 120,width = 400))
               ),
               
               p("=================================================================================================================================="),
               tableOutput("table"),
               plotOutput("Trend_Chart")
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
  
  base <- reactive({if(is.null(input$base))return(NULL)
                    MRL_chart_base <- read.csv(input$base$datapath, header = TRUE, stringsAsFactors = FALSE)
                    n <- nrow(MRL_chart_base)
                    MRL_chart_base[(n+1),] <- NA
                    MRL_chart_base[(n+1), 1] <- format(input$dateRange[1],"%b.%Y")
                    MRL_chart_base$Month <- as.character(MRL_chart_base$Month)
                    return(MRL_chart_base)})
  
  
# Save combined "Start Work Flow" file and combined "Review" file for next month
  medswf <- reactive({if(is.null(input$current)|is.null(input$old))return(NULL)
                      med_st_old <- read_excel(input$old$datapath, sheet = 1)
                      med_st_new <- read_excel(input$current$datapath[input$current$name==paste0("Start_Workflow_Medical_", format(input$dateRange[1],"%B"), ".xlsx")], sheet = 1)
                      med_st_new$time_stamp <- strptime(med_st_new$time_stamp, format = "%m/%d/%Y %H:%M:%S")
                      med_st_new$start_date <- med_st_new$time_stamp
                      med_st_old$start_date <- med_st_old$time_stamp
                      medswf <- rbind(med_st_old, med_st_new)
                      return(medswf)})
  
  medreview <- reactive({if(is.null(input$current)|is.null(input$old))return(NULL)
                         med_end_old <- read_excel(input$old$datapath, sheet = 2)
                         med_end_new <- read_excel(input$current$datapath[input$current$name==paste0("Signoff_Medical_", format(input$dateRange[1],"%B"), ".xlsx")], sheet = 1)
                         med_end_new$time_stamp <- strptime(med_end_new$time_stamp, format = "%m/%d/%Y %H:%M:%S")
                         med_end_new$end_date <- med_end_new$time_stamp
                         med_end_old$end_date <- med_end_old$time_stamp
                         medreview <- rbind(med_end_old, med_end_new)
                         medreview})
  
  regswf <- reactive({if(is.null(input$current)|is.null(input$old))return(NULL)
                      reg_st_old <- read_excel(input$old$datapath, sheet = 3)
                      reg_st_new <- read_excel(input$current$datapath[input$current$name==paste0("Start_Workflow_Regulatory_", format(input$dateRange[1],"%B"), ".xlsx")], sheet = 1)
                      reg_st_new$time_stamp <- strptime(reg_st_new$time_stamp, format = "%m/%d/%Y %H:%M:%S")
                      reg_st_new$start_date <- reg_st_new$time_stamp
                      reg_st_old$start_date <- reg_st_old$time_stamp
                      regswf <- rbind(reg_st_old, reg_st_new)
                      regswf})
  
  regreview <- reactive({if(is.null(input$current)|is.null(input$old))return(NULL)
                         reg_end_old <- read_excel(input$old$datapath, sheet = 4)
                         reg_end_new <- read_excel(input$current$datapath[input$current$name==paste0("Signoff_Regulatory_", format(input$dateRange[1],"%B"), ".xlsx")], sheet = 1)
                         reg_end_new$time_stamp <- strptime(reg_end_new$time_stamp, format = "%m/%d/%Y %H:%M:%S")
                         reg_end_new$end_date <- reg_end_new$time_stamp
                         reg_end_old$end_date <- reg_end_old$time_stamp
                         regreview <- rbind(reg_end_old, reg_end_new)
                         regreview}) 
  
  legswf <- reactive({if(is.null(input$current)|is.null(input$old))return(NULL)
                      leg_st_old <- read_excel(input$old$datapath, sheet = 5)
                      leg_st_new <- read_excel(input$current$datapath[input$current$name==paste0("Start_Workflow_Legal_", format(input$dateRange[1],"%B"), ".xlsx")], sheet = 1)
                      leg_st_new$time_stamp <- strptime(leg_st_new$time_stamp, format = "%m/%d/%Y %H:%M:%S")
                      leg_st_new$start_date <- leg_st_new$time_stamp
                      leg_st_old$start_date <- leg_st_old$time_stamp
                      legswf <- rbind(leg_st_old, leg_st_new)
                      legswf})
  
  legreview <- reactive({if(is.null(input$current)|is.null(input$old))return(NULL)
                         leg_end_old <- read_excel(input$old$datapath, sheet = 6)
                         leg_end_new <- read_excel(input$current$datapath[input$current$name==paste0("Signoff_Legal_", format(input$dateRange[1],"%B"), ".xlsx")], sheet = 1)
                         leg_end_new$time_stamp <- strptime(leg_end_new$time_stamp, format = "%m/%d/%Y %H:%M:%S")
                         leg_end_new$end_date <- leg_end_new$time_stamp
                         leg_end_old$end_date <- leg_end_old$time_stamp
                         legreview <- rbind(leg_end_old, leg_end_new)
                         legreview}) 
  
  
#--------------------------------------- Medical -------------------------------------------
  
# Read new "End" file, assign "doc.type" and "T_market" based on "object_name, and delete duplicates
  med_end_new <- reactive({if(is.null(input$current))return(NULL)
                           med_end_new <- read_excel(input$current$datapath[input$current$name==paste0("Signoff_Medical_", format(input$dateRange[1],"%B"), ".xlsx")], sheet = 1)
                           med_end_new$doc.type <- substr(med_end_new$object_name,1,2)
                           med_end_new$T_market <- stri_sub(med_end_new$object_name,-2,-1)
                           med_end_new$T_market <- ifelse(med_end_new$T_market=="12","MM", med_end_new$T_market)
                           #observe(print(table(med_end_new$doc.type)))
                           #observe(print(table(med_end_new$T_market))) #Summary by 'doc.type' and 'T_market'. Keep here just in case we need them in the future
                           #substring targeted market: AP: advertisement & Promotional; FC: field Communication; SC: Scientific Communication
                           #validate(need(length(unique(med_end_new$audited_obj_id))==dim(med_end_new)[1], "Warning Message: Please check duplicates in current Medical review file."))
                           
                           med_end_new$end_date <- strptime(med_end_new$time_stamp, format = "%m/%d/%Y %H:%M:%S")
                           med_end_new$end_date <- as.Date(med_end_new$end_date, "%m/%d/%Y")
                           med_end_new <- med_end_new[rev(order(med_end_new$end_date)),]
                           med_end_new <- med_end_new[!duplicated(med_end_new$audited_obj_id),] 
                           med_end_new})   
  
# Merge the combined "Start" file and the new "End" file by "audited_obj_id", and calculate time difference (unit: days)
  med_month <- reactive({if(is.null(medswf())|is.null(med_end_new()))return(NULL)
                         medswf <- medswf()
                         med_end_new <- med_end_new()
                         medswf <- medswf[!duplicated(medswf$audited_obj_id),] 
                         medswf$start_date <- as.Date(medswf$start_date, "%m/%d/%Y")
                         med_month <- merge(x = medswf, y = med_end_new, by.x = "audited_obj_id", by.y = "audited_obj_id", all.y = T)
                         med_month <- med_month[rev(order(med_month$end_date)),]
                         med_month <- med_month[!duplicated(med_month$audited_obj_id),] 
                         med_month$Time.diff <- med_month$end_date - med_month$start_date
                         med_month$Day <- ceiling(round(as.numeric(med_month$Time.diff, units = "days"), 10)) + 1 # round to next nearest integer
                         med_month}) 
  
# Subset outliers: > 10 days       
  MedOutliers <- reactive({if(is.null(med_month()))return(NULL)
                           med.ol.s <- subset(med_month(), Day > 10, select = c("object_name.x", "version_label.x", "user_name.x", "user_name.y", "string_1.x","start_date","end_date","Day"))
                           colnames(med.ol.s) <- c("Object", "Version", "Start_username", "Signoff_username", "View Type","Start_date","End_date","Aging")
                           med.ol.s <- med.ol.s[order(med.ol.s$Aging,decreasing = T),]
                           med.ol.s})  
  
  
#--------------------------------------- Regulatory -------------------------------------------
  
# Read new "End" file, assign "doc.type" and "T_market" based on "object_name, and delete duplicates
  reg_end_new <- reactive({if(is.null(input$current))return(NULL)
                           reg_end_new <- read_excel(input$current$datapath[input$current$name==paste0("Signoff_Regulatory_", format(input$dateRange[1],"%B"), ".xlsx")], sheet = 1)
                           reg_end_new$doc.type <- substr(reg_end_new$object_name,1,2)
                           reg_end_new$T_market <- stri_sub(reg_end_new$object_name,-2,-1)
                           reg_end_new$T_market <- ifelse(reg_end_new$T_market=="12","MM", reg_end_new$T_market)
                           #observe(print(table(reg_end_new$doc.type)))
                           #observe(print(table(reg_end_new$T_market))) #Summary by 'doc.type' and 'T_market'. Keep here just in case we need them in the future
                           #substring targeted market: AP: advertisement & Promotional; FC: field Communication; SC: Scientific Communication
                           #validate(need(length(unique(reg_end_new$audited_obj_id))==dim(reg_end_new)[1], "Warning Message: Please check duplicates in current Regulatory review file."))
    
                           reg_end_new$end_date <- strptime(reg_end_new$time_stamp, format = "%m/%d/%Y %H:%M:%S")
                           reg_end_new$end_date <- as.Date(reg_end_new$end_date, "%m/%d/%Y")
                           reg_end_new <- reg_end_new[rev(order(reg_end_new$end_date)),]
                           reg_end_new <- reg_end_new[!duplicated(reg_end_new$audited_obj_id),] 
                           reg_end_new})   
  
# Merge Regulatory file and Medical Review file by "audited_obj_id", delete duplicates, and calculate time difference (unit: days)      
  med_reg <- reactive({if(is.null(regswf())|is.null(reg_end_new())|is.null(medreview()))return(NULL)
                       regswf <- regswf()
                       reg_end_new <- reg_end_new()
                       medreview <- medreview()
                       regswf <- regswf[!duplicated(regswf$audited_obj_id),] 
                       regswf$start_date <- as.Date(regswf$start_date, "%m/%d/%Y")
                       medreview$end_date <- as.Date(medreview$end_date, "%m/%d/%Y")
                       reg.new <- merge(regswf, reg_end_new, by.x = "audited_obj_id", by.y = "audited_obj_id", all.y = T)  
                       med_reg <- merge(medreview, reg.new, by.x = "audited_obj_id", by.y = "audited_obj_id", all.y = T )
                       med_reg <- med_reg[rev(order(med_reg$end_date.y, med_reg$end_date.x)),]
                       med_reg <- med_reg[!duplicated(med_reg$audited_obj_id),] 
                       med_reg$reg_start <- med_reg$end_date.x

                       for (i in 1: nrow(med_reg)){
                            if (med_reg$reg_start[i] %in% NA & med_reg$string_1.x[i] %in% "RL  Review"){
                                                       med_reg$reg_start[i] <- med_reg$start_date[i]
                            }
                       }  # insert start time for "RL review"
                       
                       if (any(is.na(med_reg$reg_start))){
                             med_reg <- med_reg[-c(which(is.na(med_reg$reg_start))),]
                       } 
    
                       med_reg$Time.diff <- med_reg$end_date.y - med_reg$reg_start
                       med_reg$Day <- ceiling(round(as.numeric(med_reg$Time.diff, units = "days"), 10)) + 1
                       med_reg})
  
# Subset outliers: > 10 days       
  RegOutliers <- reactive({if(is.null(med_reg()))return(NULL)
    reg.ol.s <- subset(med_reg(), Day > 10 )
                           reg.ol.s <- subset(med_reg(), Day > 10, select = c("object_name.x", "version_label.x", "user_name.x", "user_name.y", "string_1.x", "reg_start", "end_date.y", "Day"))
                           colnames(reg.ol.s) <- c("Object", "Version", "Start_username", "Signoff_username", "View Type","Start_date","End_date","Aging")
                           reg.ol.s <- reg.ol.s[order(reg.ol.s$Aging,decreasing = T),]
                           reg.ol.s})   
  
#--------------------------------------- Legal -------------------------------------------
  
# Read new "End" file, assign "doc.type" and "T_market" based on "object_name, and delete duplicates
  leg_end_new <- reactive({if(is.null(input$current))return(NULL)
                           leg_end_new <- read_excel(input$current$datapath[input$current$name==paste0("Signoff_Legal_", format(input$dateRange[1],"%B"), ".xlsx")], sheet = 1)
                           leg_end_new$doc.type <- substr(leg_end_new$object_name,1,2)
                           leg_end_new$T_market <- stri_sub(leg_end_new$object_name,-2,-1)
                           leg_end_new$T_market <- ifelse(leg_end_new$T_market=="12","MM", leg_end_new$T_market)
                           #observe(print(table(leg_end_new$doc.type)))
                           #observe(print(table(leg_end_new$T_market))) #Summary by 'doc.type' and 'T_market'. Keep here just in case we need them in the future
                           #substring targeted market: AP: advertisement & Promotional; FC: field Communication; SC: Scientific Communication
                           #validate(need(length(unique(leg_end_new$audited_obj_id))==dim(leg_end_new)[1], "Warning Message: Please check duplicates in current legulatory review file."))
    
                           leg_end_new$end_date <- strptime(leg_end_new$time_stamp, format = "%m/%d/%Y %H:%M:%S")
                           leg_end_new$end_date <- as.Date(leg_end_new$end_date, "%m/%d/%Y")
                           leg_end_new <- leg_end_new[rev(order(leg_end_new$end_date)),]
                           leg_end_new <- leg_end_new[!duplicated(leg_end_new$audited_obj_id),] 
                           leg_end_new}) 
  
# Merge Legal file and Regulatory Review file, delete duplicates, and calculate time difference (unit: days)      
  reg_leg <- reactive({if(is.null(legswf())|is.null(leg_end_new())|is.null(regreview()))return(NULL)
                       legswf <- legswf()
                       leg_end_new <- leg_end_new()
                       regreview <- regreview()
                       legswf <- legswf[!duplicated(legswf$audited_obj_id),] 
                       legswf$start_date <- as.Date(legswf$start_date, "%m/%d/%Y")
                       regreview$end_date <- as.Date(regreview$end_date, "%m/%d/%Y")
                       leg.new <- merge(legswf,leg_end_new, by.x = "audited_obj_id", by.y = "audited_obj_id", all.y = T)  
                       reg_leg <- merge(regreview,leg.new, by.x = "audited_obj_id", by.y = "audited_obj_id", all.y = T )
                       reg_leg <- reg_leg[rev(order(reg_leg$end_date.y,reg_leg$end_date.x)),]
                       reg_leg <- reg_leg[!duplicated(reg_leg$audited_obj_id),] 
                       reg_leg$leg_start <- reg_leg$end_date.x
    
                       for (i in 1: length(reg_leg$leg_start)){
                            if (reg_leg$leg_start[i]%in%NA & reg_leg$string_1.x[i]%in%"Legal  Only  Review"){
                                                        reg_leg$leg_start[i] <- reg_leg$start_date[i]
                            }
                       }         # insert start time for "Legal  Only  Review"
    
                       reg_leg$Time.diff <- reg_leg$end_date.y - reg_leg$leg_start
                       reg_leg$Day <- ceiling(round(as.numeric(reg_leg$Time.diff, units = "days"), 10)) +1
                       reg_leg})  

# Subset outliers: > 10 days       
  LegOutliers <- reactive({if(is.null(reg_leg()))return(NULL)
                           leg.ol.s <- subset(reg_leg(), Day > 10, select = c("object_name.x", "version_label.x", "user_name.x", "user_name.y", "string_1.x", "leg_start", "end_date.y", "Day"))
                           colnames(leg.ol.s) <- c("Object", "Version", "Start_username", "Signoff_username", "View Type", "Start_date", "End_date", "Aging")
                           leg.ol.s <- leg.ol.s[order(leg.ol.s$Aging,decreasing = T),]
                           leg.ol.s}) 
  
#---------------------------------- Summarize Data ------------------------------------------------
  
# TABLE 1: Total Documents Processed by Process and Region   
  
  Summary.by.Process.Region <- reactive({if(is.null(med_month())|is.null(med_reg())|is.null(reg_leg()))return(NULL)
                                         med.Mrt.table <- as.data.frame(table(med_month()$T_market))
                                         sub.mrt <- subset(med.Mrt.table, Var1=="GB" | Var1=="MM" | Var1=="US")
                                         sub.mrt$Var1 <- as.character(sub.mrt$Var1)
                                         med.Mrt.table <- rbind(sub.mrt, c("OT", sum(subset(med.Mrt.table,Var1!="GB" & Var1!="MM" & Var1!="US")[,2])))[,2]  
                                         #sum(as.numeric(med.Mrt.table$Freq)) 
    
                                         reg.Mrt.table <- as.data.frame(table(med_reg()$T_market))
                                         sub.mrt <- subset(reg.Mrt.table, Var1=="GB" | Var1=="MM" | Var1=="US")
                                         sub.mrt$Var1 <- as.character(sub.mrt$Var1)
                                         reg.Mrt.table <- rbind(sub.mrt, c("OT",sum(subset(reg.Mrt.table, Var1!="GB" & Var1!="MM" & Var1!="US")[,2])))[,2] 
                                         #sum(as.numeric(reg.Mrt.table$Freq))
    
                                         leg.Mrt.table <- as.data.frame(table(reg_leg()$T_market))
                                         sub.mrt <- subset(leg.Mrt.table, Var1=="GB" | Var1=="MM" | Var1=="US")
                                         sub.mrt$Var1 <- as.character(sub.mrt$Var1)
                                         leg.Mrt.table <- rbind(sub.mrt, c("OT",sum(subset(leg.Mrt.table, Var1!="GB" & Var1!="MM" & Var1!="US")[,2])))[,2] 
                                         #sum(as.numeric(leg.Mrt.table$Freq))
    
                                         table1 <- data.frame()
                                         
                                         table1[1:5,1] <- c("GB", "MulM", "USA", "OTHER", "Total (all region)")
                                         table1[1:5,2] <- c("", "", "", "", "")
                                         table1[1:4,3] <- med.Mrt.table
                                         table1[5,3] <- sum(as.numeric(table1[1:4,3]))
                                         table1[1:4,4] <- reg.Mrt.table
                                         table1[5,4] <- sum(as.numeric(table1[1:4,4]))
                                         table1[1:4,5] <- leg.Mrt.table
                                         table1[5,5] <- sum(as.numeric(table1[1:4,5]))
                                         table1[1,6] <- format(sum(as.numeric(table1[1,3:5])), digits = 0)
                                         table1[2,6] <- format(sum(as.numeric(table1[2,3:5])), digits = 0)
                                         table1[3,6] <- format(sum(as.numeric(table1[3,3:5])), digits = 0)
                                         table1[4,6] <- format(sum(as.numeric(table1[4,3:5])), digits = 0)
                                         table1[5,6] <- format(sum(as.numeric(table1[5,3:5])), digits = 0)
                                         names(table1) <- c("Na", "Na1", "med", "reg", "leg", "tot")
                                         table1}) 
  
  
# TABLE 2: Calculate total number of reviewed, average review time and % target achieved by process     
  Summary.by.Process <- reactive({if(is.null(med_month())|is.null(med_reg())|is.null(reg_leg()))return(NULL)
                                  med.viewed <- length(med_month()$Day[!is.na(med_month()$Day)]) 
                                  med.avg <- round(mean(med_month()$Day, na.rm = T), 1)    
                                  num.med.day <- as.numeric(med_month()$Day<=10)
                                  med.target <- as.character(percent(sum(num.med.day,na.rm = T)/med.viewed, digits = 0))
    
                                  reg.viewed <- length(med_reg()$Day[!is.na(med_reg()$Day)])
                                  reg.avg <- round(mean(med_reg()$Day, na.rm = T), 1)
                                  num.reg.day <- as.numeric(med_reg()$Day<=10)
                                  reg.target <- as.character(percent(sum(num.reg.day,na.rm = T)/reg.viewed, digits = 0))
    
                                  leg.viewed <- length(reg_leg()$Day[!is.na(reg_leg()$Day)])
                                  leg.avg <- round(mean(reg_leg()$Day, na.rm = T), 1)
                                  num.leg.day <- as.numeric(reg_leg()$Day<=10)
                                  leg.target <- as.character(percent(sum(num.leg.day,na.rm = T)/leg.viewed, digits = 0))
    
                                  table2 <- data.frame()
                                  
                                  table2[1:3,1] <- c("Medical (<= 10 days)", "Regulatory (<= 10 days)", "Legal (<= 10 days)")
                                  table2[1:3,2] <- c("", "", "")
                                  table2[1:3,3] <- c("", "", "")
                                  table2[1,4:6] <- c(med.viewed, med.avg, med.target)
                                  table2[2,4:6] <- c(reg.viewed, reg.avg, reg.target)
                                  table2[3,4:6] <- c(leg.viewed, leg.avg, leg.target)
                                  names(table2) <- c("Na", "Na1", "Na2", "tot", "avg", "percent")
                                  table2})
  
# TABLE 3: Regulatory Summary by Region 
  Reg.Summary <- reactive({if(is.null(med_reg())|is.null(Summary.by.Process()))return(NULL)
                           table3 <- data.frame(matrix(vector(), 5, 6), stringsAsFactors = FALSE)
                           table3[1:5,1] <- c("GB (<= 10 days)", "MulM (<= 10 days)", "USA (<= 10 days)", "Other (<= 10 days)", "Total")
                           table3[1:5,2] <- c("", "", "", "", "")
                           table3[1:5,3] <- c("", "", "", "", "")
                           table3[ ,4] <- as.numeric(Summary.by.Process.Region()[ ,4])
                           table3[1,5] <- as.character(round(mean(med_reg()$Day[med_reg()$T_market=="GB"]),1))
                           table3[2,5] <- as.character(round(mean(med_reg()$Day[med_reg()$T_market=="MM"]),1))
                           table3[3,5] <- as.character(round(mean(med_reg()$Day[med_reg()$T_market=="US"]),1))
                           table3[4,5] <- as.character(round(mean(med_reg()$Day[med_reg()$T_market!="GB" & med_reg()$T_market!="MM" & 
                                                             med_reg()$T_market!="US" ]),1))
                           table3[5,5] <- Summary.by.Process()[2,5]
                           table3[1,6] <- as.character(percent(sum(as.numeric(med_reg()$Day[med_reg()$T_market=="GB"]<=10),na.rm = T)/table3[1,4], digits = 0))
                           table3[2,6] <- as.character(percent(sum(as.numeric(med_reg()$Day[med_reg()$T_market=="MM"]<=10),na.rm = T)/table3[2,4], digits = 0))
                           table3[3,6] <- as.character(percent(sum(as.numeric(med_reg()$Day[med_reg()$T_market=="US"]<=10),na.rm = T)/table3[3,4], digits = 0))
                           table3[4,6] <- as.character(percent(sum(as.numeric(med_reg()$Day[med_reg()$T_market!="GB" & med_reg()$T_market!="MM" & 
                                                                      med_reg()$T_market!="US" ]<=10),na.rm = T)/table3[4,4], digits = 0))
                           table3[5,6] <- Summary.by.Process()[2,6]
                           table3[ ,4] <- format(table3[ ,4], digits = 0)
                           names(table3) <- c("Na", "Na1", "Na2", "tot", "avg", "percent")
                           table3})
  
# Outliers
  
  outliers <- reactive({if(is.null(MedOutliers())|is.null(RegOutliers())|is.null(LegOutliers()))return(NULL)
                        med <- MedOutliers()
                        reg <- RegOutliers()
                        leg <- LegOutliers()
                        outliers <- data.frame()
                        outliers[1:3,1] <- c("Medical", "Regulatory", "Legal")
                        outliers[1:3,2] <- c("", "", "")
                        outliers[1:3,3] <- c(nrow(med), nrow(reg), nrow(leg))
                        names(outliers) <- c("Process", "Na", "# of Outliers")
                        return(outliers)
                        })
  
# MRL Chart Base 

  newbase <- reactive({if(is.null(base())|is.null(Summary.by.Process()))return(NULL)
                       base <- base()
                       n <- nrow(base)
                       base[n, 2:4] <- Summary.by.Process()[1:3,5]                       
                       base[n, 5:7] <- Summary.by.Process()[1:3,4]
                       base[n,20:22] <- Summary.by.Process()[1:3,6]
                       base[n,23] <- Summary.by.Process.Region()[5,6]
                       base[n, 8:10] <- Summary.by.Process.Region()[1,3:5]
                       base[n,11:13] <- Summary.by.Process.Region()[2,3:5]
                       base[n,14:16] <- Summary.by.Process.Region()[3,3:5]
                       base[n,17:19] <- Summary.by.Process.Region()[4,3:5]
                       base[n,24:27] <- Reg.Summary()[1:4,5]
                       base[n,28:31] <- Reg.Summary()[1:4,6]
                       base})
#---------------------------------------- Plot ----------------------------------------------------------------------- 
# Re-organize "Base" Data for Plot 
  plot.12M <- reactive({if(is.null(newbase()))return(NULL)
                        plot.12M <- newbase()[c((dim(newbase())[1]-11):dim(newbase())[1]),] 
                        plot.12M$Month.Plot <- seq(1, 12, 1)
                        plot.12M})
  
  data1 <- reactive({if(is.null(plot.12M()))return(NULL)
                     data1 <- data.frame(department = rep(c("Medical","Regulatory","Legal"), each = nrow(plot.12M())),
                                         Month = plot.12M()$Month, 
                                         total = as.numeric(c(plot.12M()$Med.tot, plot.12M()$Reg.Tot, plot.12M()$Leg.Tot)),
                                         average = as.numeric(c(plot.12M()$Med_avg, plot.12M()$Reg_avg, plot.12M()$Leg.avg)))})
  
  data2 <- reactive({if(is.null(data1())|is.null(plot.12M()))return(NULL)
                     rbind(data1(),  
                           data.frame(department = "total", 
                                      Month = plot.12M()$Month, 
                                      total = as.numeric(plot.12M()$Total), 
                                      average = NA))})
  
# Plot 
  Trend_Chart <- function()({if(is.null(data1())|is.null(data2())|is.null(plot.12M()))return(NULL)
                             line = ggplot(data = data1()) +
                                    ggtitle("MRL Trend Chart for Document View Time and Number of Reviews") +
                                    labs(x = "", y = "Avg. Review Time (Calendar Day)") +
                                    theme(plot.title = element_text(size = 14, hjust = 0.5)) +
                                    geom_point(aes(x = Month, y = average, group = department, color = department)) +
                                    geom_line(aes(x = Month, y = average, group = department, color = department)) +
                                    geom_label_repel(aes(x = Month, y = average, group = department, color = department, label = average), show.legend = F )+
                                    scale_x_discrete(limits = plot.12M()$Month) +
                                    scale_y_continuous("Avg. Review Time (Calendar Day)") + 
                                    scale_color_manual(values = c(Legal = "chartreuse3", Medical = "blueviolet", Regulatory = "deeppink3")) +
                                    geom_hline(yintercept = 10,color = 'blue', lty = 2) +
                                    annotate("text", x = 2, y = 9, label = "Target <=10 Calendar Days", color = 'blue', size = 3) +
                                    theme(text = element_text(size = 10),legend.title = element_blank(), legend.position = "right",
                                          axis.title.y = element_text(size = 10))
    
                             bar = ggplot(data = data1()) +
                                   #ggtitle("MRL Trend Chart for Number of Reviews")+
                                   labs(x = "", y = "Number of Reviews") +
                                   theme(plot.title = element_text(size = 14, hjust = 0.5)) +
                                   geom_bar(aes(x = Month, y = total, color = department, fill = department), stat = "identity", width = 0.5) +
                                   geom_text(aes(x = Month, y = total, label = total, group=department), color = "white", position = position_stack(vjust = 0.5),
                                             show.legend = F) +
                                   geom_line(data = subset(data2(), department=="total"), aes(x = Month, y = total, group = 1), color = "red", lwd = 1.2) + 
                                   geom_text(data = subset(data2(), department=="total"), aes(x = Month, y = total, label = total), color = "red",
                                             position = position_stack(vjust = 1.04),show.legend = F) +
                                   scale_x_discrete(limits = plot.12M()$Month) +
                                   scale_y_continuous("Number of Reviews") +
                                   scale_color_manual(values = c(Medical = "blueviolet", Regulatory = "deeppink3", Legal = "chartreuse3")) +
                                   scale_fill_manual(values = c(Medical = "blueviolet", Regulatory = "deeppink3", Legal = "chartreuse3")) +
                                   theme(text = element_text(size = 10), legend.key = element_blank(), legend.title = element_blank(), legend.position = "right",
                                         axis.title.y = element_text(size = 10))
    
    grid.newpage() 
    grid.draw(rbind(ggplotGrob(line), ggplotGrob(bar), size = "last"))})    
#------------------------------------------------------------------------------------------------------------------    
  
  #output$Trend_Chart <- renderPlot({Trend_Chart()},height = 800, width = 1300, res = 150)
  #output$table <- renderTable({base()}, rownames = TRUE)
  
  
  output$base <- downloadHandler(filename = function(){paste("MRL_chart_base_", format(input$dateRange[1],"%B %Y"), ".csv", sep="")},
                                 content = function(file){withProgress(message = paste0("Downloading Base File"),
                                                                       value = 0, 
                                                                       {for (i in 1:10) {incProgress(1/10, detail = paste(i*10, "%"))
                                                                         Sys.sleep(0.1)}
                                                          write.csv(newbase(), file, sep=",", row.names = FALSE)
                                                          })
                                 })
  output$outlier <- downloadHandler(filename = function() {paste("MRL_outliers_", format(input$dateRange[1],"%B %Y"), ".xlsx",sep = "")},
                                    content = function(file) {wb <- createWorkbook()
                                                              addWorksheet(wb = wb, sheet = "Med Outliers" )
                                                              writeData(wb = wb, sheet = 1, x =  MedOutliers())
                                                              addWorksheet(wb = wb, sheet = "Reg Outliers" )
                                                              writeData(wb = wb, sheet = 2, x =  RegOutliers())
                                                              addWorksheet(wb = wb, sheet = "Leg Outliers" )
                                                              writeData(wb = wb, sheet = 3, x =  LegOutliers())
                                                              withProgress(message = paste0("Downloading Outliers File"),
                                                                           value = 0, 
                                                                           {for (i in 1:10) {incProgress(1/10, detail = paste(i*10, "%"))
                                                                            Sys.sleep(0.1)}
                                                              
                                                              saveWorkbook(wb, file, overwrite = T)
                                                                           })
                                                               })
  
  output$historical <- downloadHandler(filename = function() {paste("MRL_historical_data_", format(input$dateRange[1],"%B %Y"), ".xlsx",sep = "")},
                                       content = function(file) {wb <- createWorkbook()
                                       
                                       addWorksheet(wb = wb, sheet = "Med_start" )
                                       writeData(wb = wb, sheet = 1, x =  medswf())
                                       addWorksheet(wb = wb, sheet = "Med_end" )
                                       writeData(wb = wb, sheet = 2, x =  medreview())
                                       addWorksheet(wb = wb, sheet = "Reg_start" )
                                       writeData(wb = wb, sheet = 3, x =  regswf())
                                       addWorksheet(wb = wb, sheet = "Reg_end" )
                                       writeData(wb = wb, sheet = 4, x =  regreview())
                                       addWorksheet(wb = wb, sheet = "Leg_start" )
                                       writeData(wb = wb, sheet = 5, x =  legswf())
                                       addWorksheet(wb = wb, sheet = "Leg_end" )
                                       writeData(wb = wb, sheet = 6, x =  legreview())
                                       withProgress(message = paste0("Downloading Historical File"),
                                                    value = 0, 
                                                    {for (i in 1:10) {incProgress(1/10, detail = paste(i*10, "%"))
                                                      Sys.sleep(0.1)}
                                       saveWorkbook(wb, file, overwrite = T)
                                                    })
                                       })

  output$report <- downloadHandler(filename = function() {paste("MRL_Report_", format(input$dateRange[1],"%B %Y"),".pptx", sep = "")},
                                   content = function(file) {
                                                 PR <- Summary.by.Process.Region()
                                                 SP <- Summary.by.Process()
                                                 RSP <- Reg.Summary()
                                                 Outliers <- outliers()
                                                 std_border = fp_border(color="black") 
                                     
                                                 title1 <- block_list(
                                                                fpar(ftext("MRL Dashboard", shortcuts$fp_bold(font.size = 36, color = "white")))
                                                 )
                                                 
                                                 title2 <- block_list(
                                                                fpar(ftext(paste("Data Coverage: through", format(input$dateRange[2], "%h %d %Y")), shortcuts$fp_bold(font.size = 18, color = "white"))),
                                                                fpar(ftext(paste("Data Freeze by: CEDAR Admin"), shortcuts$fp_bold(font.size = 18, color = "white"))),
                                                                fpar(ftext(paste("CEDAR data received on", input$received), shortcuts$fp_bold(font.size = 18, color = "white")))
                                                 )
                                                 
                                                 title3 <- block_list(
                                                                fpar(ftext(paste("Report Date:", Sys.time()), shortcuts$fp_bold(font.size = 18, color = "white"))),
                                                                fpar(ftext(paste("Qi Zhang on behalf of Metrics Central"), shortcuts$fp_bold(font.size = 18, color = "white"))),
                                                                fpar(ftext(paste(" and Quality Compliance"), shortcuts$fp_bold(font.size = 18, color = "white")))
                                                 )
                                     
                                                 caption1 <- block_list(
                                                                  fpar(ftext("Total Documents Processed by Process and Region", shortcuts$fp_bold(font.size = 22, color = "cyan4")))
                                                 )
                                     
                                                 caption2 <- block_list(
                                                                  fpar(ftext("Summary of Review Cycle Time by Process", shortcuts$fp_bold(font.size = 22, color = "cyan4")))
                                                 )
                                     
                                                 caption3_1 <- block_list(
                                                                    fpar(ftext("Regulatory", shortcuts$fp_bold(font.size = 22, color = "orange")))
                                                 )
                                     
                                                 caption3_2 <- block_list(
                                                                    fpar(ftext("Summary of Review Cycle Time by Process", shortcuts$fp_bold(font.size = 22, color = "cyan4")))
                                                 )
                                                 
                                                 caption4 <- block_list(
                                                                  fpar(ftext("Number of Outliers (> 10 days)", shortcuts$fp_bold(font.size = 22, color = "cyan4")))
                                                 )
                                            
                                                 legend1 <- block_list(
                                                                 fpar(ftext("GB: Great Britain;", shortcuts$fp_italic(color="black"))),
                                                                 fpar(ftext("MulM: Multi-Markets;", shortcuts$fp_italic(color="black"))),
                                                                 fpar(ftext("USA: the United States of America;", shortcuts$fp_italic(color="black"))),
                                                                 fpar(ftext("OTHER: all other countries combined.", shortcuts$fp_italic(color="black")))   
                                                 )
                                     
                                                 legend2 <- block_list(
                                                                fpar(ftext("Days: calendar days", shortcuts$fp_italic(color="black")))
                                                 )
                                                 
                                                 comment0 <- block_list(
                                                                  fpar(ftext(paste(input$slide2), shortcuts$fp_bold(font.size = 16)))  
                                                 )
                                     
                                                 comment1 <- block_list(
                                                                  fpar(ftext(paste(input$slide3), shortcuts$fp_bold(font.size = 16)))  
                                                 )
                                     
                                                 comment2 <- block_list(
                                                                  fpar(ftext(paste(input$slide4), shortcuts$fp_bold(font.size = 16)))  
                                                 )
                                     
                                                 comment3 <- block_list(
                                                                  fpar(ftext(paste(input$slide5), shortcuts$fp_bold(font.size = 16)))  
                                                 )
                                                 
                                                 comment4 <- block_list(
                                                                  fpar(ftext(paste(input$slide6), shortcuts$fp_bold(font.size = 16)))  
                                                 )
                                     
                                     
                                                 table1 <- flextable(PR) %>% 
                                       
                                                           set_header_labels(Na = "Targeted Market", 
                                                                             Na1 = "Targeted Market", 
                                                                             med = "Medical", reg = "Regulatory", 
                                                                             leg = "Legal", tot = "Total") %>%
                                       
                                                           merge_h(part = "header") %>% 
                                                           bold(part = "header") %>%
                                                           fontsize(size = 16, part = "all") %>%  
                                                           bg(bg = "cyan4", part = "header") %>%
                                                           color(color = "white", part = "header") %>%
                                                           bg(i = c(1,3,5), bg = "azure2", part = "body") %>%
                                                           merge_at(i = 1, j = 1:2, part = "body")%>% 
                                                           merge_at(i = 2, j = 1:2, part = "body")%>% 
                                                           merge_at(i = 3, j = 1:2, part = "body")%>% 
                                                           merge_at(i = 4, j = 1:2, part = "body")%>% 
                                                           merge_at(i = 5, j = 1:2, part = "body")%>%
                                       
                                                           width(width = 1.3) %>% 
                                                           height(height = 0.55, part = "header") %>%
                                                           height(height = 0.50, part = "body") %>%
                                                           align(align = "center", part = "all")%>%
                                                           border_outer(border = std_border, part = "all") %>%
                                                           border_inner(border = std_border, part = "all") 
                                     
                                                 table2 <- flextable(SP) %>% 
                                                           
                                                           set_header_labels(Na = "Process", Na1 = "Process", Na2 = "Process",
                                                                             tot = "Tot # Reviewed", avg = "Avg Review Time (Days)", 
                                                                             percent = "% Target Achieved") %>%
                                       
                                                           merge_h(part = "header") %>% 
                                                           bold(part = "header") %>%
                                                           fontsize(size = 16, part = "all") %>% 
                                                           bg(bg = "cyan4", part = "header") %>%
                                                           color(color = "white", part = "header") %>%
                                                           bg(i = c(1,3), bg = "azure2", part = "body") %>%
                                                           merge_at(i = 1, j = 1:3, part = "body")%>% 
                                                           merge_at(i = 2, j = 1:3, part = "body")%>% 
                                                           merge_at(i = 3, j = 1:3, part = "body")%>% 
                                                   
                                                           width(width = 1.3) %>%  
                                                           height(height = 0.55, part = "header") %>%
                                                           height(height = 0.5, part = "body") %>%
                                       
                                                           align(align = "center", part = "all")%>%
                                                           border_outer(border = std_border, part = "all") %>%
                                                           border_inner(border = std_border, part = "all")   
                                     
                                     
                                                 table3 <- flextable(RSP) %>% 
                                                   
                                                           set_header_labels(Na = "Regulatory by Region", 
                                                                             Na1 = "Regulatory by Region", 
                                                                             Na2 = "Regulatory by Region",
                                                                             tot = "Tot # Reviewed", 
                                                                             avg = "Avg Review Time (Days)", 
                                                                             percent = "% Target Achieved") %>%
                                                           fontsize(size = 16, part = "all") %>% 
                                                           merge_h(part = "header") %>% 
                                                           bold(part = "header") %>%
                                                           color(color = "white", part = "header") %>%
                                                           bg(bg = "orange", part = "header") %>%
                                                           bg(i = c(1,3, 5), bg = "antiquewhite1", part = "body") %>%
                                       
                                                           merge_at(i = 1, j = 1:3, part = "body")%>% 
                                                           merge_at(i = 2, j = 1:3, part = "body")%>% 
                                                           merge_at(i = 3, j = 1:3, part = "body")%>% 
                                                           merge_at(i = 4, j = 1:3, part = "body")%>% 
                                                           merge_at(i = 5, j = 1:3, part = "body")%>% 
                                       
                                                           width(width = 1.3) %>% 
                                                           height(height = 0.55, part = "header") %>%
                                                           height(height = 0.5, part = "body") %>%
                                       
                                                           align(align = "center", part = "all")%>%
                                                           border_outer(border = std_border, part = "all") %>%
                                                           border_inner(border = std_border, part = "all") 
                                     
                                     
                                                 table4 <- flextable(Outliers) %>% 
                                                           fontsize(size = 16, part = "all") %>% 
                                                           bold(part = "header") %>%
                                                           bg(bg = "cyan4", part = "header") %>%
                                                           color(color = "white", part = "header") %>%
                                                           bg(i = c(1,3), bg = "azure2", part = "body") %>%
                                                           merge_at(j = 1:2, part = "header") %>% 
                                                           merge_at(i = 1, j = 1:2, part = "body")%>% 
                                                           merge_at(i = 2, j = 1:2, part = "body")%>% 
                                                           merge_at(i = 3, j = 1:2, part = "body")%>%
                                       
                                                           width(width = 1.6) %>% 
                                                           height(height = 0.55, part = "header") %>%
                                                           height(height = 0.5, part = "body") %>%
                                       
                                                           align(align = "center", part = "all")%>%
                                                           border_outer(border = std_border, part = "all") %>%
                                                           border_inner(border = std_border, part = "all") 
                                     
                                     
                                     example_pp <- read_pptx() %>% 
                                       
                                         add_slide(layout = "Title Slide", master = "Office Theme") %>% 
                                                   ph_with(external_img(src = paste0(loc_path,"title_page.png")), 
                                                           location = ph_location(left = 0, top = 0, width = 10, height = 7.5) )%>%
                                                   ph_with(location = ph_location(top = 0.5, left = 0.2, width = 10, height = 1),
                                                           value = title1) %>%
                                                   ph_with(location = ph_location(top = 1.5, left = 0.2, width = 10, height = 1),
                                                           value = title2) %>%
                                                   ph_with(location = ph_location(top = 3, left = 0.2, width = 10, height = 1),
                                                           value = title3) %>%
                                       
                                        add_slide(layout = "Title and Content", master = "Office Theme") %>%
                                                  ph_with(external_img(src = paste0(loc_path,"left_corner.png")), 
                                                          location = ph_location(left = 0, top = 0, width = 2.2, height = 1.5) )%>%
                                                  ph_with(external_img(src = paste0(loc_path,"logo.png")), 
                                                          location = ph_location(left = 7.8, top = 6, width = 2.2, height = 1.5) )%>%
                                                  ph_with(dml(Trend_Chart()), 
                                                          location = ph_location(left = 0.5, top = 0.5, width = 9, height = 6.5))%>%
                                                  ph_with(external_img(src = paste0(loc_path,"legend.png")), 
                                                          location = ph_location(left = 8.5, top = 5.7, width = 0.5, height = 0.3) )%>%
                                                  ph_with(value = comment0,
                                                          location = ph_location(left = 0.5, top = 6.5, width = 5, height = 1, label = "comment0")) %>%
                                       
                                        add_slide(layout = "Title and Content", master = "Office Theme") %>% 
                                                  ph_with(external_img(src = paste0(loc_path,"left_corner.png")), 
                                                          location = ph_location(left = 0, top = 0, width = 2.2, height = 1.5) )%>%
                                                  ph_with(external_img(src = paste0(loc_path,"logo.png")), 
                                                          location = ph_location(left = 7.8, top = 6, width = 2.2, height = 1.5) )  %>%
                                                  ph_with(location = ph_location(top = 0.8, left = 1.2, width = 10, height = 1),
                                                          value = caption1) %>% 
                                                  ph_with(value = table1,
                                                          location = ph_location(left = 0.9, top = 2) )%>%
                                                  ph_with(value = legend1,
                                                          location = ph_location(left = 0.9, top = 5, width = 3, height = 1, label = "legend1")) %>%
                                                  ph_with(value = comment1,
                                                          location = ph_location(left = 1.5, top = 6.4, width = 5, height = 1, label = "comment1")) %>%
                                       
                                       
                                        add_slide(layout = "Title and Content", master = "Office Theme") %>% 
                                                  ph_with(external_img(src = paste0(loc_path,"left_corner.png")), 
                                                          location = ph_location(left = 0, top = 0, width = 2.2, height = 1.5) )%>%
                                                  ph_with(external_img(src = paste0(loc_path,"logo.png")), 
                                                          location = ph_location(left = 7.8, top = 6, width = 2.2, height = 1.5) )  %>%
                                                  ph_with(location = ph_location(top = 1, left = 1.6, width = 10, height = 1),
                                                          value = caption2) %>% 
                                                  ph_with(value = table2, 
                                                          location = ph_location(left = 1.2, top = 2.5)) %>%
                                                  ph_with(value = legend2,
                                                          location = ph_location(left = 1.2, top = 4.7, width = 3, height = 1, label = "legend2")) %>%
                                                  ph_with(value = comment2,
                                                          location = ph_location(left = 1.1, top = 5.7, width = 5, height = 1, label = "comment2")) %>%
                                       
                                       
                                        add_slide(layout = "Title and Content", master = "Office Theme") %>% 
                                                  ph_with(external_img(src = paste0(loc_path,"left_corner.png")), 
                                                          location = ph_location(left = 0, top = 0, width = 2.2, height = 1.5) )%>%
                                                  ph_with(external_img(src = paste0(loc_path,"logo.png")), 
                                                          location = ph_location(left = 7.8, top = 6, width = 2.2, height = 1.5) )  %>%
                                                  ph_with(location = ph_location(top = 0.4, left = 4, width = 10, height = 1),
                                                          value = caption3_1) %>% 
                                                  ph_with(location = ph_location(top = 0.9, left = 1.6, width = 10, height = 1),
                                                          value = caption3_2) %>% 
                                                  ph_with(value = table3, 
                                                          location = ph_location(left = 1.2, top = 2)) %>%
                                                  ph_with(value = legend1,
                                                          location = ph_location(left = 1.2, top = 5.7, width = 3, height = 1, label = "legend1")) %>%
                                                  ph_with(value = comment3,
                                                          location = ph_location(left = 1.1, top = 6.6, width = 5, height = 1, label = "comment3")) %>%
                                       
                                        add_slide(layout = "Title and Content", master = "Office Theme") %>% 
                                                  ph_with(external_img(src = paste0(loc_path,"left_corner.png")), 
                                                          location = ph_location(left = 0, top = 0, width = 2.2, height = 1.5) )%>%
                                                  ph_with(external_img(src = paste0(loc_path,"logo.png")), 
                                                          location = ph_location(left = 7.8, top = 6, width = 2.2, height = 1.5) )  %>%
                                                  ph_with(location = ph_location(top = 1.1, left = 2.5, width = 10, height = 1),
                                                          value = caption4) %>% 
                                                  ph_with(value = table4, 
                                                          location = ph_location(left = 2.6, top = 2.3)) %>%
                                                  ph_with(value = comment4,
                                                          location = ph_location(left = 2.5, top = 4.8, width = 5, height = 1, label = "comment4")
                                                  ) 
                                     
                                        withProgress(message=paste0("Downloading PPT File"),
                                                     value=0, 
                                                     {for (i in 1:10) {incProgress(1/10, detail = paste(i*10, "%"))
                                                                       Sys.sleep(0.1)}
                                     
                                        print(example_pp, target = file)
                                                      })
                                   }
  )
  
  
}      

# Run the application 
shinyApp(ui = ui, server = server)