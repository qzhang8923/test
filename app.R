#===========================================================================================================
#
#                                 PMW Project App Phase II (New)
#                                 
#                                 Last Update: 10/5/2020
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
#install.packages("devtools")
#devtools::install_github('rstudio/DT')
#install.packages("data.table")
#install.packages("dplyr")
#install.packages("zoo")
#install.packages("lubridate")
#install.packages("ggplot2")
#library('lubridate')
#install.packages("shinyWidgets")
library(ggplot2)
library(DT)
library(shiny)
library(dplyr)
library(shinythemes)
library(readxl)
library(openxlsx)
library(data.table)
library(zoo) #function 'as.yearmon'
library(lubridate)
library(rsconnect)

options(shiny.maxRequestSize = 30*1024^10)#change uploading file's size
#MCImage <- "https://github.com/qzhang8923/images/raw/main/ConvaTec.PNG"
MCImage <- "https://raw.githubusercontent.com/qzhang15403/pics/main/ConvaTec.PNG"
#-----------------------------------------------------------------------------------------------------------

#-----------------------------------------------------------------------------------------------------------
#                                       Functions
#-----------------------------------------------------------------------------------------------------------
dropdownButton <- function(label = "", status = c("default", "primary", "success", "info", "warning", "danger"), ..., width = NULL) {
  
                              status <- match.arg(status)
                              # dropdown button content
                              html_ul <- list(class = "dropdown-menu",
                                              style = if (!is.null(width)) 
                                                           paste0("width: ", validateCssUnit(width), ";"),
                                                           lapply(X = list(...), FUN = tags$li, style = "margin-left: 10px; margin-right: 10px;")
                                               )
                              # dropdown button apparence
                              html_button <- list(class = paste0("btn btn-", status," dropdown-toggle"),
                                                   type = "button", 
                                                  `data-toggle` = "dropdown"
                                                  )
                              html_button <- c(html_button, list(label))
                              html_button <- c(html_button, list(tags$span(class = "caret")))
                             # final result
                              tags$div(class = "dropdown",
                                       do.call(tags$button, html_button),
                                       do.call(tags$ul, html_ul),
                              tags$script("$('.dropdown-menu').click(function(e) { e.stopPropagation();});") )
}

int_breaks <- function(x, n = 5) pretty(x, n)[pretty(x, n) %% 1 == 0]  #this function guarantees to only produce integer breaks. 

#-----------------------------------------------------------------------------------------------------------
#
#                                         Define UI 
#
#-----------------------------------------------------------------------------------------------------------
ui <- navbarPage(
  
  theme = shinytheme("cerulean"),
  
  "PMW Automation", # App title
  
  tabsetPanel(
    
    tabPanel("Phase II",  #Tab title
             
             sidebarPanel(width=2,
                          helpText(div(strong("Step 1:")), " Upload Most Current Complaints Master Data File.", style = "color:deepskyblue"),
                          helpText("(.xlsx file; contains 51 columns)"),
                          fileInput("Complaints", "", multiple = FALSE, accept = c(".xlsx")),
            
                          helpText(div(strong("Step 2:")), " Upload Most Current Sales Data.", style = "color:deepskyblue"),
                          fileInput("Sales"," ",multiple = FALSE, accept = c(".xlsx")),
                          
                          helpText(div(strong("Step 3:")), "Please Choose Target Information.", style = "color:deepskyblue"),
                          
                          
                          dateRangeInput('dateRange', label = 'Date range input: yyyy-mm-dd',
                                         start = "2019-04-01", end = Sys.Date()),
                          helpText("(Date range no less than 15 months)"),
                          
                          helpText(div(strong("Select Business Unit")), style = "color:black"),
                          dropdownButton(label = "BU", status = "default",
                                         actionButton(inputId = "all BU", label = "(Un)select all"),
                                         verbatimTextOutput(outputId = "BU"),
                                         checkboxGroupInput(inputId = "BU_dropdown", label = "", choices = "Loading...")
                          ),
                          
                          helpText(div(strong("Select PHRI Category")),style = "color:black"),
                          dropdownButton(label = "PHRI", status = "default",
                                         actionButton(inputId = "all PHRI", label = "(Un)select all"),
                                         verbatimTextOutput(outputId = "PHRI"),
                                         checkboxGroupInput(inputId = "PHRI_dropdown", label = "", choices = "Loading...")
                          ),
                          
                          helpText(strong("Select PMR Attribute Category"),style = "color:black"),
                          dropdownButton(label = "PMR", status = "default",
                                         actionButton(inputId = "all PMR", label = "(Un)select all"),
                                         verbatimTextOutput(outputId = "PMR"),
                                         checkboxGroupInput(inputId = "PMR_dropdown", label = "", choices = "Please select PHRI...")
                          ),
                          
                          helpText(strong("Select Harm Code or Malfunction Code"),style = "color:black"),
                          
                          
                          dropdownButton(label = "Harm Code", status = "default",
                                         actionButton(inputId = "all HC", label = "(Un)select all"),
                                         verbatimTextOutput(outputId = "Harm Code"),
                                         checkboxGroupInput(inputId = "HC", label = "", choices = "Please select PHRI...")
                          ),
                          
                          helpText(""),
                          
                          dropdownButton(label = "Malfunction Code", status = "default",
                                         actionButton(inputId = "all MF", label = "(Un)select all"),
                                         verbatimTextOutput(outputId = "Malfuction Code"),
                                         checkboxGroupInput(inputId = "MF", label = "", choices = "Please select PHRI...")
                          ),
                          
                          #pickerInput("locInput","Location", choices=c("New Mexico", "Colorado", "California"), options = list(`actions-box` = TRUE), multiple = T),
                          helpText(""),
                          
                          helpText(div(strong("Step 4:")), " Download Output File.", style = "color:deepskyblue"),
                          downloadButton("download", "Table (by MF)"),
                          downloadButton("download2", "Alert Trips (by MF and PMR)"),
                          downloadButton("download_plot1", "Plot 1"),
                          downloadButton("download_plot2", "Plot 2"),
                          downloadButton("download_plot3", "Plot 3")
             ),
             
             
             mainPanel(
               
                     p("=================================================================================================================================="),
                     
                     fluidRow(
                             column(5, h1("PMW Automation Phase II (NEW)"),
                                       p("Copyright belong to ConvaTec Metrics Central team"),
                                       p("R Cloud v1.0 Updated Date: Oct 27 2020")),
                             column(5, img(src = MCImage, height = 120,width = 400))
                     ),
                     
                     p("=================================================================================================================================="),
                     
                       tabsetPanel(
                         
                         tabPanel("Table by MF", 
                                  dataTableOutput("result"),
                                  
                                  fluidRow(
                                    column(4,
                                           radioButtons(inputId = "code", 
                                                        label = "Choose Report Type", 
                                                        choices = list("Complaints by Malfunction Code" = "MF", "CPM by Malfunction Code" = "CPM_MF6"), 
                                                        selected = "MF"
                                                        #inline = T
                                                         )
                                           
                                     ),
                                    column(7, 
                                           p("###############################################################################"),
                                           strong(" *Trending Flaggings:"),
                                           p(" Flagging on", span ("YELLOW", style = "background-color: yellow"), "for the 4th month if consecutive Complaints/CPM increases for any 3  months;"),
                                           p(" Flagging on", span ("RED", style = "background-color: red"), " if the Most Recent Quarter Average > 3 and % Difference > 50%;"),
                                           p(" Flagging on", span ("ORANGE", style = "background-color: orange"), " if the Most Recent Quarter Average  < or = 3 and % Difference > 50%."),
                                           p("###############################################################################")
                                    )  
                                  ),
                                  
                                  DT::dataTableOutput("table"),
                                  htmlOutput("text")),
                         tabPanel("Table by MF and PMR", 
                                  fluidRow(
                                    column(4,
                                           radioButtons(inputId = "code_PMR", 
                                                        label = "Choose Report Type", 
                                                        choices = list("Complaints by Malfunction Code and PMR" = "MF_PMR", "CPM by Malfunction Code and PMR" = "CPM_MF_PMR6"), 
                                                        selected = "MF_PMR"
                                                        #inline = T
                                           )
                                           
                                    ),
                                    column(7, 
                                           p("###############################################################################"),
                                           strong(" *Trending Flaggings:"),
                                           p(" Flagging on", span ("YELLOW", style = "background-color: yellow"), "for the 4th month if consecutive Complaints/CPM increases for any 3  months;"),
                                           p(" Flagging on", span ("RED", style = "background-color: red"), " if the Most Recent Quarter Average > 3 and % Difference > 50%;"),
                                           p(" Flagging on", span ("ORANGE", style = "background-color: orange"), " if the Most Recent Quarter Average  < or = 3 and % Difference > 50%."),
                                           p("###############################################################################")
                                    )),
                                  DT::dataTableOutput("table_PMR")),
                                  
                         tabPanel("Plot 1: Complaints & CPM Trend Chart", 
                                  plotOutput("Trend_Chart1")),
                         tabPanel("Plot 2: Complaint Analysis Chart", plotOutput("Trend_Chart2")),
                         tabPanel("Plot 3: CPM Analysis Chart", 
                                  plotOutput("Trend_Chart3"))
                         
                         
                       )   
                       
             )
    )
  )
)


#-----------------------------------------------------------------------------------------------------------
#
#                                          Define Server
#
#-----------------------------------------------------------------------------------------------------------


server <- function(input, output, session) {
  
               data <- reactive({if(is.null(input$Complaints)){return(NULL)
                               } else {data <- read_excel(paste(input$Complaints$datapath), 1)
                                       data <- data[data$`Date Created` >= input$dateRange[1] & data$`Date Created` <= input$dateRange[2]+1,]
                                       
                                    
                                       validate(need(nrow(data[is.na(data$`Business Unit (Franchise)`),]) == 0, 
                                                     "Warning: There are blanks in Business Unit"))
                                       validate(need(nrow(data[is.na(data$`PHRI-2`),]) == 0, 
                                                     "Warning: There are blanks in PHRI"))
                                       validate(need(nrow(data[is.na(data$`PMR Attribute Category`),]) == 0, 
                                                     "Warning: There are blanks in PMR Attribute Category"))
                                       validate(need(nrow(data[is.na(data$`MF`),]) == 0, 
                                                     "Warning: There are blanks in MF"))
                                       validate(need(nrow(data[is.na(data$`HC`),]) == 0, 
                                                     "Warning: There are blanks in HC"))
                                       # data$`PHRI`[is.na(data$`PHRI`)] <- "N/A"
                                       # data$`PMR Attribute Category`[is.na(data$`PMR Attribute Category`)] <- "N/A"
                                       # data$MF[is.na(data$MF)] <- "N/A"
                                       # data$HC[is.na(data$HC)] <- "N/A"
                                       data[data$`PMR Attribute Category`=="Aquacel Ag Extra",]$`PMR Attribute Category` <- "Aquacel Ag EXTRA"
                                       data[data$`PMR Attribute Category`=="Aloe VESTA Body Wash",]$`PMR Attribute Category` <- "Aloe Vesta Body Wash"
                                       data[data$`PMR Attribute Category`=="Aloe VESTA Cleansers",]$`PMR Attribute Category` <- "ALOE VESTA Cleansers"
                                       data[data$`PMR Attribute Category`=="Aloe VESTA Lotions & Creams",]$`PMR Attribute Category` <- "ALOE VESTA Lotions & Creams"
                                       #print(nrow(data[data$`PMR Attribute Category`=="2pc-Pouch-Natura+" & data$MF=="OST-PMC3.12"&data$CREATED=="72020",]))
                                       return(data)}
                                 })

               observeEvent(data(), {BU <- sort(unique(data()$`Business Unit (Hyperion)`))
                                     updateCheckboxGroupInput(session = session, inputId = "BU_dropdown", choices = BU, selected = "")
               })
               
               observeEvent(input$"all BU", {BU <- sort(unique(data()$`Business Unit (Hyperion)`))
                                             if (is.null(input$"BU_dropdown")){updateCheckboxGroupInput(session = session, inputId = "BU_dropdown", selected = BU)
                                           } else {updateCheckboxGroupInput(session = session, inputId = "BU_dropdown", selected = "")
                                           }
               })
               
               observeEvent(input$"BU_dropdown", {selected_PHRI <- sort(unique(data()[data()$`Business Unit (Hyperion)` %in% input$"BU_dropdown",]$`PHRI`))
                                                  updateCheckboxGroupInput(session = session, inputId = "PHRI_dropdown", choices = selected_PHRI, selected = selected_PHRI)
               })
               
               observeEvent(input$"all PHRI", {selected_PHRI <- sort(unique(data()[data()$`Business Unit (Hyperion)` %in% input$"BU_dropdown",]$"PHRI"))
                                               if (is.null(input$"PHRI_dropdown")){updateCheckboxGroupInput(session = session, inputId = "PHRI_dropdown", selected = selected_PHRI)
                                               } else {updateCheckboxGroupInput(session = session, inputId = "PHRI_dropdown", selected = "")
                                               }
               })
               
               observeEvent(input$"PHRI_dropdown", {selected_PMR <- sort(unique(data()[data()$`Business Unit (Hyperion)` %in% input$"BU_dropdown" & 
                                                                                       data()$"PHRI" %in% input$"PHRI_dropdown",]$`PMR Attribute Category`))
                                                    updateCheckboxGroupInput(session = session, inputId = "PMR_dropdown", choices = selected_PMR, selected = selected_PMR)
               })
               
               observeEvent(input$"all PMR", {selected_PMR <- sort(unique(data()[data()$`Business Unit (Hyperion)` %in% input$"BU_dropdown" & 
                                                                                 data()$"PHRI" %in% input$"PHRI_dropdown",]$`PMR Attribute Category`))
                                             if (is.null(input$"PMR_dropdown")){updateCheckboxGroupInput(session = session, inputId = "PMR_dropdown", selected = selected_PMR)
                                              } else {updateCheckboxGroupInput(session = session, inputId = "PMR_dropdown", selected = "")
               }
               })
               
               observeEvent(input$"PHRI_dropdown", {selected_HC <- sort(unique(data()[data()$`Business Unit (Hyperion)` %in% input$"BU_dropdown" &
                                                                                      data()$"PHRI" %in% input$"PHRI_dropdown",]$`HC`))
                                                    updateCheckboxGroupInput(session = session, inputId = "HC", choices = selected_HC, selected = selected_HC)
               })
               
               observeEvent(input$"all HC", {selected_HC <- sort(unique(data()[data()$`Business Unit (Hyperion)` %in% input$"BU_dropdown" & 
                                                                               data()$"PHRI" %in% input$"PHRI_dropdown",]$`HC`))
                                            if (is.null(input$"HC")){updateCheckboxGroupInput(session = session, inputId = "HC", selected = selected_HC)
                                            } else {updateCheckboxGroupInput(session = session, inputId = "HC", selected = "")
               }
               })
               
               
               observeEvent(input$"PHRI_dropdown", {selected_MF <- sort(unique(data()[data()$`Business Unit (Hyperion)` %in% input$"BU_dropdown" &
                                                                                      data()$"PHRI" %in% input$"PHRI_dropdown",]$`MF`))
                                                    updateCheckboxGroupInput(session = session, inputId = "MF", choices = selected_MF, selected = selected_MF)
               })
               
               observeEvent(input$"all MF", {selected_MF <- sort(unique(data()[data()$`Business Unit (Hyperion)` %in% input$"BU_dropdown" &
                                                                               data()$"PHRI" %in% input$"PHRI_dropdown",]$`MF`))
                                            if (is.null(input$"MF")){updateCheckboxGroupInput(session = session, inputId = "MF", selected = selected_MF)
                                            } else {updateCheckboxGroupInput(session = session, inputId = "MF", selected = "")
               }
               })
               
               
            
               
               output$"Harm Code" <- renderPrint(input$HC)
               output$"Malfuction Code" <- renderPrint(input$MF)
               output$"PHRI" <- renderPrint(input$"PHRI_dropdown")
               output$"PMR" <- renderPrint(input$"PMR_dropdown")
               
#---------------------------------------------------------------------------------------------------------------------------------------  
#                                       Read Data and Data Manipulations
#---------------------------------------------------------------------------------------------------------------------------------------               
               
               Complaints <- reactive({if(is.null(data())){return(NULL)
                                       } else {Complaints <- data()
                                               Complaints$`Date Created` <- as.Date(Complaints$`Date Created`,"%m %d %Y")
                                               validate(need(input$dateRange[2]-input$dateRange[1] > 450,
                                                             "Warning Message: Selected Date Range is less than 15 months."))
                                               Complaints <- Complaints[Complaints$`Date Created` >= input$dateRange[1] & Complaints$`Date Created` <= input$dateRange[2],]
                                               Complaints <- Complaints[Complaints$`Business Unit (Hyperion)` %in% input$"BU_dropdown",]
                                               #print(nrow(Complaints[Complaints$`PMR Attribute Category`=="2pc-Pouch-Natura+" & Complaints$MF=="OST-PMC3.12" & Complaints$CREATED=="72020",]))
                                               #print(nrow(Complaints[Complaints$`Business Unit (Hyperion)`=="Ostomy" & Complaints$CREATED=="92020",]))
                                               return(Complaints)
                                               }
               })
  
  
               Sales <- reactive({if(is.null(input$Sales)|is.null(input$PHRI_dropdown)){return(NULL)
                                  } else {
                                          sales <- read_excel(paste(input$Sales$datapath), sheet = "Global Eaches", skip = 2, col_types=c("text"))
                                          sales$`Hyperion Business Unit` <- ifelse(sales$`Hyperion Business Unit`=="Critical Care", 
                                                                                       "Total Critical Care",
                                                                                       sales$`Hyperion Business Unit`)
                                          
                                          sales$`Hyperion Business Unit` <- ifelse(sales$`Hyperion Business Unit`=="Continence Care", 
                                                                                       "Total Continence Care",
                                                                                       sales$`Hyperion Business Unit`)
                                          
                                          sales$`Hyperion Business Unit` <- ifelse(sales$`Hyperion Business Unit`=="OSTOMY", 
                                                                                       "Ostomy",
                                                                                       sales$`Hyperion Business Unit`)
                                         
                                          sales$`Hyperion Business Unit` <- ifelse(sales$`Hyperion Business Unit`=="WOUND", 
                                                                                       "Wound",
                                                                                       sales$`Hyperion Business Unit`)
                                      
                                          sales <- sales[sales$`Hyperion Business Unit` %in% input$"BU_dropdown",]  
                                          list <- which(colnames(sales) %in% c('ICC (SR Code)', 'CCC Reference #', 'Hyperion Description', 'PMR Category', 'Hyperion Business Unit', 'Total Eaches')) 
                                          sales <- sales[,-list] 
                                          colnames(sales)[3:ncol(sales)] <- as.character(seq(as.Date("2016-01-01"), by = "month", length.out = ncol(sales)-2))
                                          #colnames(sales)[3:ncol(sales)] <- as.character(as.Date(as.numeric(colnames(sales)[3:ncol(sales)]), origin = "1899-12-30"))
                                          sales[3:ncol(sales)] <- sapply(sales[3:ncol(sales)], as.character)
                                          sales[is.na(sales)|sales=="-"|sales=="'0"] <- "0"
                                          sales[3:ncol(sales)] <- sapply(sales[3:ncol(sales)], as.numeric)
                                          sales <- sales[sales$'PHRI' %in% input$'PHRI_dropdown',]
                                          validate(need(nrow(sales) > 0, "No Sales in current PHRI/PMR Category."))
                                        
                                          tb_s <- data.frame()
                                          for (i in 3:ncol(sales)){
                                                each <- aggregate(sales[i], by=list(sales$`PMR Attribute Category`, sales$PHRI), sum)
                                                each$month <- as.Date(colnames(each)[3])
                                                colnames(each)[3] <- "count"
                                                tb_s <- rbind(tb_s,each)
                                                 }
                                          colnames(tb_s)[1:2] <- c('PMR Attribute Category', 'PHRI') 
                                          #tb_s <- tb_s[tb_s$month >= (input$dateRange[1]%m-% months(6)) & tb_s$month <= (input$dateRange[2]%m-% months(1)),] #obtain data for previous 6 months
                                          #tb_s <- tb_s[tb_s$month >= (input$dateRange[1]) & tb_s$month <= (input$dateRange[2]),]
                                          
                                          return(tb_s)
                                          }
               })
               
               
               Sales_6 <- reactive({if(is.null(Sales())|is.null(input$PMR_dropdown)){return(NULL)
                                     } else {tb_s <- Sales()
                                             tb_s <- tb_s[tb_s$month >= (input$dateRange[1]%m-% months(6)) & tb_s$month <= (input$dateRange[2]%m-% months(1)),] #obtain data for previous 6 months
                                             Sales_s <- tb_s[(tb_s$`PMR Attribute Category` %in% input$PMR_dropdown),]
                                             validate(need(nrow(Sales_s) > 0, "No Sales in current PMR Category."))
                                             tb1 <- aggregate(Sales_s$count, by=list(Sales_s$month),sum)
                                             colnames(tb1) <- c("month", "count")
                                             tb2 <- data.frame(matrix(tb1$count, ncol=length(unique(tb1$month))))
                                             colnames(tb2) <- unique(as.character(as.yearmon(tb1$month)))
                                             rownames(tb2) <- "Total Sales"
                                             mean <- data.frame()
                                             n <- ncol(tb2)
                                             for(i in 6:n){mean[1,(i-5)] <- round(rowMeans(tb2[,(i-5):i]),2)}
                                             return(mean)
                                             }
               })
               
#---------------------------------------------------------------------------------------------------------------------------------------  
#                                            Table (2 Report Options)
#---------------------------------------------------------------------------------------------------------------------------------------               
               
               table_c <- reactive({if(is.null(Complaints())){return(NULL)
                                    } else {table_c <- Complaints() %>% 
                                                       group_by(month = floor_date(`Date Created`,"month"), HC, MF, `PHRI`, `PMR Attribute Category`) %>% 
                                                       summarize(count = n())
                                            table_c$month <- as.yearmon(table_c$month)
                                            #write.csv(table_c, "C:/Users/QI15403/OneDrive - ConvaTec Limited/Desktop/table_c.csv")
                                            return(table_c)
                                            }
               })
                 
               
               HC <- reactive({if(is.null(table_c())|is.null(input$MF)|is.null(input$`PHRI_dropdown`)|is.null(input$PMR_dropdown)|is.null(input$HC)){return(NULL)
                                } else {table_c <- table_c()[(table_c()$MF %in% input$MF) & (table_c()$`PHRI` %in% input$`PHRI_dropdown`) & (table_c()$`PMR Attribute Category` %in% input$PMR_dropdown),]
                                        validate(need(nrow(table_c) > 0, "No complaints in current date range and PMR/PHRI Category."))
                                        tb1<-aggregate(table_c$count, by = list(table_c$month, table_c$HC), sum)
                                        colnames(tb1) <- c("month","HC","count")
                                        dates <- seq(input$dateRange[1], input$dateRange[2], by = "month")
                                        months <- as.yearmon(dates)
                                        catg <- unique(tb1$HC)
                                        tb <- as.data.frame(lapply(expand.grid(months,catg),unlist))
                                        colnames(tb) <- c("month","HC")
                                        tb2 <- merge(tb, tb1, by= c('month','HC'), all.x=T)
                                        tb2$count[is.na(tb2$count)] <- 0
                                        HC <- data.frame(matrix(tb2$count, ncol=length(unique(tb2$month))))
                                        colnames(HC) <- as.character(unique(tb2$month))
                                        rownames(HC) <- unique(tb2$HC)
                                        HC[nrow(HC)+1,] <- round(colSums(HC),digits=0)
                                        rownames(HC)[nrow(HC)] <- "Total per Month"
                                        return(HC)
                                }
               })
                                        
                                        
               MF <- reactive({if(is.null(table_c())|is.null(input$HC)|is.null(input$`PHRI_dropdown`)|is.null(input$PMR_dropdown)|is.null(input$MF)){return(NULL)
                               } else {table_c <- table_c()[(table_c()$HC %in% input$HC) & (table_c()$`PHRI` %in% input$`PHRI_dropdown`) & (table_c()$`PMR Attribute Category` %in% input$PMR_dropdown),]
                                       
                                       validate(need(nrow(table_c) > 0, "No complaints in current date range and PMR/PHRI Category."))
                                   
                                       tb1 <- aggregate(table_c$count, by=list(table_c$month,table_c$MF),sum)
                                       colnames(tb1) <- c("month","MF","count")
                                       dates <- seq(input$dateRange[1], input$dateRange[2], by = "month")
                                       months <- as.yearmon(dates)
                                       catg <- unique(tb1$MF)
                                       tb <- as.data.frame(lapply(expand.grid(months,catg),unlist))
                                       colnames(tb) <- c("month","MF")
                                       tb2 <- merge(tb, tb1, by= c('month','MF'), all.x=T)
                                       tb2$count[is.na(tb2$count)] <- 0
                                       MF <- data.frame(matrix(tb2$count, ncol=length(unique(tb2$month))))
                                       colnames(MF) <- as.character(unique(tb2$month))
                                       rownames(MF) <- unique(tb2$MF)
                                       MF[nrow(MF)+1,] <- round(colSums(MF),digits=0)
                                       rownames(MF)[nrow(MF)] <- "Total per Month"
                                       return(MF)
                               }
               })
               
               
               CPM_HC6 <- reactive({if(is.null(HC())|is.null(Sales_6())){return(NULL)
                                    } else {HC <- as.matrix(HC())
                                            Sales <- as.matrix(Sales_6())
                                            CPM_HC <- t(t(HC)/Sales[1,])
                                            CPM_HC <- round(CPM_HC*1000000, 2)
                                            rownames(CPM_HC)[nrow(CPM_HC)] <- "Total CPM per Month"
                                            return(CPM_HC)
                                           }
               })
               
               CPM_MF6 <- reactive({if(is.null(MF())|is.null(Sales_6())){return(NULL)
                                   } else {MF <- as.matrix(MF())
                                           Sales <- as.matrix(Sales_6())
                                           CPM_MF <- t(t(MF)/Sales[1,])
                                           CPM_MF <- round(CPM_MF*1000000, 2)
                                           rownames(CPM_MF)[nrow(CPM_MF)] <- "Total CPM per Month"
                                           return(CPM_MF)
                                           }
               })
               
               code <- reactive({if(is.null(HC())|is.null(MF())){return(NULL)
                                  } else {switch(input$code,
                                                 #HC = HC(),
                                                 MF = MF(),
                                                 #CPM_HC6 = data.frame(CPM_HC6()),
                                                 CPM_MF6 = data.frame(CPM_MF6())
                                                 )}
               })
  
               tb <- reactive({if(is.null(code())){return(NULL)
                               } else {tb <- code()
                                       tb$'12 Month Average' <- round(rowMeans(tb[,c((ncol(tb)-14):(ncol(tb)-3))] , na.rm = T), 2)
                                       tb$'Most Recent Quarter Average' <- round(rowMeans(tb[,(ncol(tb)-3):(ncol(tb)-1)], na.rm = T),2)
                                       tb$"% Difference" <- ifelse(tb$'12 Month Average'==0 & tb$'Most Recent Quarter Average'==0,
                                                                   tb$"% Difference" <- NA,
                                                                   ifelse(tb$'12 Month Average'==0 & !tb$'Most Recent Quarter Average'==0, 
                                                                          tb$"% Difference" <- 100, 
                                                                          tb$"% Difference" <- (tb$'Most Recent Quarter Average'- tb$'12 Month Average')*100/tb$'12 Month Average'))
                                       tb$"% Difference" <- round(tb$"% Difference",2)
                                       tb$MAX <- apply(tb[,-c((ncol(tb)-3):ncol(tb))], 1, max) 
                                       return(tb)
                                       }
               })
               
               tb1 <- reactive({if(is.null(tb())){return(NULL)
                                } else {tb1 = tb()
                                        x <- ncol(tb1)
                                        y <- nrow(tb1)
                                        dat <- data.frame()
                                        dat2 <- data.frame()
                                        dat3 <- data.frame()
               
                                        for (i in 1:(x-4-3)) {
                                             for (j in 1:y) {
                                                    dat[j,i] <- tb1[j,i+3] > tb1[j,i+2] & tb1[j,i+2] > tb1[j,i+1] & tb1[j,i+1] > tb1[j,i]
                                             }
                                        }
               
                                        for (j in 1:(y-1)) {
                                                    dat2[j,1] <- ifelse (tb1[j,x-2] > 3 & tb1[j,x-1] > 50, 1,
                                                        ifelse (tb1[j,x-2] <= 3 & tb1[j,x-1] > 50, 2,  0))
                                                  
                                        }
                                       dat2[y,1] <- 0
                                       colnames(dat2)[1] <- "index"
                                       tb1 <- cbind(tb1, dat, dat2)
                                       return(tb1)
                                       }
               })
               

               tb_d <- reactive({if(is.null(tb1())){return(NULL)
                                  } else {tb1 <- tb1()[,-c((ncol(tb1())-(ncol(tb1())-8)/2):ncol(tb1()))]
                                          return(tb1)
               
               }})
               
               output$table <- DT::renderDataTable(if(is.null(tb1())){return(NULL)
                                                    } else {
                                                    DT::datatable(tb1(), options = list(
                                                                                         columnDefs = list(list(targets = seq((ncol(tb1())-8)/2+4+4,ncol(tb1())), visible = FALSE)),
                                                                                         #pageLength = nrow(tb1())
                                                                                         pageLength = 15
                                                                                        )
                                                                  ) %>% formatStyle(
                                                                                    tail(head(colnames(tb1()), (ncol(tb1())-8)/2+3), (ncol(tb1())-8)/2),
                                                                                    head(tail(colnames(tb1()),(ncol(tb1())-8)/2+1),(ncol(tb1())-8)/2),
                                                                                    backgroundColor = styleEqual(c(0,1), c('white', 'yellow'))
                                                                  ) %>% formatStyle(
                                                                                    tail(head(colnames(tb1()), (ncol(tb1())-8)/2+6), 2),
                                                                                    tail(colnames(tb1()),1),
                                                                                    backgroundColor = styleEqual(c(0,1,2), c('white', 'red','orange'))
                                                                  ) 
                                                    
               })
               
#---------------------------------------------------------------------------------------------------------------------------------------  
#                                       Output trips for each single PMR in one file
#---------------------------------------------------------------------------------------------------------------------------------------    
              PMR <- reactive({if(is.null(input$PMR_dropdown)){return(NULL)
                              }else{PMR <- input$PMR_dropdown
                                    return(PMR)
                              }
              })    
               
               
              Com_PMR_MF <- reactive({if(is.null(input$PHRI_dropdown)){return(NULL)
                                     }else{table <- data.frame()
                                     
                                           for(i in 1:length(PMR())){
                                               PMR <- PMR()
                                               #MF - complaints by MF
                                               table_c <- table_c()[(table_c()$HC %in% input$HC) & 
                                                                    (table_c()$`PHRI` %in% input$`PHRI_dropdown`) & 
                                                                    (table_c()$`PMR Attribute Category`==PMR[i]),]
                                               tb1 <- aggregate(table_c$count, by = list(table_c$month, table_c$MF), sum)
                                               colnames(tb1) <- c("month","MF","count")
                                               dates <- seq(input$dateRange[1], input$dateRange[2], by = "month")
                                               months <- as.yearmon(dates)
                                               catg <- unique(tb1$MF)
                                               tb <- as.data.frame(lapply(expand.grid(months,catg),unlist))
                                               colnames(tb) <- c("month", "MF")
                                               tb2 <- merge(tb, tb1, by= c('month','MF'), all.x=T)
                                               tb2$count[is.na(tb2$count)] <- 0
                                               MF <- data.frame(matrix(tb2$count, ncol=length(unique(tb2$month))))
                                               colnames(MF) <- as.character(unique(tb2$month))
                                               rownames(MF) <- NULL
                                               MF$MF <- unique(tb2$MF)
                                               MF$PMR <- rep(PMR[i], nrow(MF))
                                               table <- rbind(table, MF)
                                           }
                                               n <- ncol(table)
                                               m <- nrow(table)
                                               table <- cbind(table$PMR, table$MF, table[,1:(n-2)])
                                               colnames(table) <- c("PMR", "MF", as.character(unique(tb2$month)))
                                               table$MF <- as.character(table$MF)
                                               table$PMR <- as.character(table$PMR)
                                               table[m+1,1:2] <- c("", "Total per Month")
                                               table[m+1,3:n] <- colSums(table[1:m,3:n])
                                               table$'12 Month Average' <- round(rowMeans(table[,c((ncol(table)-14):(ncol(table)-3))] , na.rm = T), 2)
                                               table$'Most Recent Quarter Average' <- round(rowMeans(table[,(ncol(table)-3):(ncol(table)-1)], na.rm = T),2)
                                               table$"% Difference" <- ifelse(table$'12 Month Average'==0 & table$'Most Recent Quarter Average'==0,
                                                                           table$"% Difference" <- NA,
                                                                           ifelse(table$'12 Month Average'==0 & !table$'Most Recent Quarter Average'==0,
                                                                                  table$"% Difference" <- 100,
                                                                                  table$"% Difference" <- (table$'Most Recent Quarter Average'- table$'12 Month Average')*100/table$'12 Month Average'))
                                               table$"% Difference" <- round(table$"% Difference",2)
                                               table$MAX <- apply(table[,-c(1, 2, (ncol(table)-3):ncol(table))], 1, max)
                                               x <- ncol(table)
                                               y <- nrow(table)
                                               dat <- data.frame()
                                               dat2 <- data.frame()
                                               dat3 <- data.frame()
                                               
                                               for (i in 3:(x-4-3)) {
                                                 for (j in 1:y) {
                                                   dat[j,i-2] <- table[j,i+3] > table[j,i+2] & table[j,i+2] > table[j,i+1] & table[j,i+1] > table[j,i]
                                                 }
                                               }
                                               
                                               for (j in 1:(y-1)) {
                                                 dat2[j,1] <- ifelse (table[j,x-2] > 3 & table[j,x-1] > 50, 1,
                                                                      ifelse (table[j,x-2] <= 3 & table[j,x-1] > 50, 2,  0))
                                                 
                                               }
                                               
                                               dat2[y,1] <- 0
                                               colnames(dat2)[1] <- "index"
                                               table <- cbind(table, dat, dat2)
                                               return(table)
                                     }
              }) 
              
              
              nosales <- reactive({if(is.null(input$PHRI_dropdown)){return(NULL)
                                  }else{sales <- subset(Sales(), select = c("PHRI", "PMR Attribute Category"))
                                        sales <- unique(sales[c("PHRI", "PMR Attribute Category")])
                                        sales$match <- rep("Yes",nrow(sales))
                                        tb_s <- subset(table_c(), select = c("PHRI", "PMR Attribute Category"))
                                        tb_s <- unique(tb_s[c("PHRI", "PMR Attribute Category")])
                                        tb_s <- tb_s[(tb_s$`PHRI` %in% input$`PHRI_dropdown`) & 
                                                     (tb_s$`PMR Attribute Category` %in% input$`PMR_dropdown`),]

                                        all <- merge(tb_s, sales, 
                                                     by.x = c("PHRI", "PMR Attribute Category"), 
                                                     by.y = c("PHRI", "PMR Attribute Category"), 
                                                     all.x = T)
                                        unmatched <- all[is.na(all$match),1:2]
                                        
                                        # list <- c()
                                        # for(j in 1:length(input$`PHRI_dropdown`)){
                                        #    PMR <- unique(table_c()[table_c()$PHRI==input$`PHRI_dropdown`[j],]$`PMR Attribute Category`)
                                        #    nosales_PMR <- c()
                                        #    nosales_PHRI <- c()
                                        #    for(i in 1:length(PMR)){
                                        #       Sales_s <- tb_s[(tb_s$`PHRI`==input$`PHRI_dropdown`[j]) & 
                                        #                       (toupper(tb_s$`PMR Attribute Category`)==toupper(PMR[i])),]    
                                        #       
                                        #       if(nrow(Sales_s)==0){nosales_PMR <- rbind(nosales_PMR, PMR[i])
                                        #                            nosales_PHRI <- rbind(nosales_PHRI, PHRI[i])
                                        #                      }else{nosales_PMR <- nosales_PMR
                                        #                            nosales_PHRI <- nosales_PHRI}
                                        #    }  
                                        #   
                                        #   if(length(nosales_PMR)==0){list <- list
                                        #                        }else{list <- rbind(list, nosales_PMR)}
                                        # }
                                        return(unmatched)
                                  }}) #no corresponding sales in Sales Eaches File
              
              
              CPM_PMR_MF <- reactive({if(is.null(input$PHRI_dropdown)|is.null(input$PMR_dropdown)|is.null(Sales())|is.null(table_c())){return(NULL)
                                }else{PHRI <- input$PHRI_dropdown
                                      sales <- Sales()
                                      sales <- sales[sales$month >= (input$dateRange[1]%m-% months(6)) & 
                                                     sales$month <= (input$dateRange[2]%m-% months(1)),] #obtain data for previous 6 months
                                      
                                      table_c <- table_c()[(table_c()$HC %in% input$HC) & 
                                                           (table_c()$`PHRI` %in% input$`PHRI_dropdown`) & 
                                                           (table_c()$`PMR Attribute Category`%in% input$`PMR_dropdown`),]
                                      
                                      table <- data.frame()
                                
                                      for(j in 1:length(PHRI)){
                                          PMR <- unique(table_c[table_c$PHRI==PHRI[j],]$`PMR Attribute Category`) 
                                          PMR <- intersect(PMR, input$`PMR_dropdown`)
                                          PMR <- PMR[!(PMR %in% nosales()[nosales()$PHRI==PHRI[j],]$`PMR Attribute Category`)]
                                          validate(need(!is.na(PMR), "Please select at least one PMR category from each selected PHRI."))
                                          
                                          table_PMR <- data.frame()
                                          
                                      for(i in 1:length(PMR)){
                                          #Sales_6  
                                          Sales_s <- sales[(sales$`PHRI`==PHRI[j]) & 
                                                           (toupper(sales$`PMR Attribute Category`)==toupper(PMR[i])),]
                                          tb1 <- aggregate(Sales_s$count, by = list(Sales_s$month), sum)
                                          colnames(tb1) <- c("month", "count")
                                          tb2 <- data.frame(matrix(tb1$count, ncol=length(unique(tb1$month))))
                                          colnames(tb2) <- unique(as.character(as.yearmon(tb1$month)))
                                          mean <- data.frame()
                                          n <- ncol(tb2)
                                          for(k in 6:n){mean[1,(k-5)] <- round(rowMeans(tb2[,(k-5):k]),2)}
                                          Sales_6 <- mean
                                          
                                          #MF - complaints by MF
                                          table_mf <- table_c[(table_c$`PHRI`==PHRI[j]) & 
                                                              (table_c$`PMR Attribute Category`==PMR[i]),]
                                          tb1 <- aggregate(table_mf$count, by = list(table_mf$month, table_mf$MF), sum)
                                          colnames(tb1) <- c("month", "MF", "count")
                                          dates <- seq(input$dateRange[1], input$dateRange[2], by = "month")
                                          months <- as.yearmon(dates)
                                          catg <- unique(tb1$MF)
                                          tb <- as.data.frame(lapply(expand.grid(months,catg),unlist))
                                          colnames(tb) <- c("month", "MF")
                                          tb2 <- merge(tb, tb1, by= c('month','MF'), all.x=T)
                                          tb2$count[is.na(tb2$count)] <- 0
                                          MF <- data.frame(matrix(tb2$count, ncol = length(unique(tb2$month))))
                                          colnames(MF) <- as.character(unique(tb2$month))
                                          rownames(MF) <- NULL
                                          MF$MF <- unique(tb2$MF)
                                          MF$PMR <- rep(PMR[i], nrow(MF))
                                   
                                          
                                          #CPM_MF6 
                                          CPM_MF <- as.matrix(MF[,1:(ncol(MF)-2)])
                                          Sales <- as.matrix(Sales_6)
                                          CPM_MF <- t(t(CPM_MF)/Sales[1,])
                                          CPM_MF <- round(CPM_MF*1000000, 2)
                                          CPM_MF <- data.frame(CPM_MF)
                                          CPM_MF$MF <- unique(tb2$MF)
                                          CPM_MF$PMR <- rep(PMR[i], nrow(MF))
                                          table_PMR <- rbind(table_PMR, CPM_MF)
                                      }
                                          
                                          table <- rbind(table, table_PMR)
                                      }
                                      
                                          n <- ncol(table)
                                          m <- nrow(table)
                                          table <- cbind(table$PMR, table$MF, table[,1:(n-2)])
                                          colnames(table) <- c("PMR", "MF", as.character(unique(tb2$month)))
                                          table$MF <- as.character(table$MF)
                                          table$PMR <- as.character(table$PMR)
                                          table[m+1,1:2] <- c("", "Total per Month")
                                          table[m+1,3:n] <- colSums(table[1:m,3:n])
                                          table$'12 Month Average' <- round(rowMeans(table[,c((ncol(table)-14):(ncol(table)-3))] , na.rm = T), 2)
                                          table$'Most Recent Quarter Average' <- round(rowMeans(table[,(ncol(table)-3):(ncol(table)-1)], na.rm = T),2)
                                          table$"% Difference" <- ifelse(table$'12 Month Average'==0 & table$'Most Recent Quarter Average'==0,
                                                                         table$"% Difference" <- NA,
                                                                         ifelse(table$'12 Month Average'==0 & !table$'Most Recent Quarter Average'==0,
                                                                                table$"% Difference" <- 100,
                                                                                table$"% Difference" <- (table$'Most Recent Quarter Average'- table$'12 Month Average')*100/table$'12 Month Average'))
                                          table$"% Difference" <- round(table$"% Difference",2)
                                          table$MAX <- apply(table[,-c(1, 2, (ncol(table)-3):ncol(table))], 1, max)
                                          x <- ncol(table)
                                          y <- nrow(table)
                                          dat <- data.frame()
                                          dat2 <- data.frame()
                                      
                                          for (i in 3:(x-4-3)) {
                                              for (j in 1:y) {
                                                  dat[j,i-2] <- table[j,i+3] > table[j,i+2] & 
                                                                table[j,i+2] > table[j,i+1] & 
                                                                table[j,i+1] > table[j,i]
                                                  }
                                          }
                                      
                                          for (j in 1:(y-1)) {
                                              dat2[j,1] <- ifelse (table[j,x-2] > 3 & table[j,x-1] > 50, 1,
                                                              ifelse (table[j,x-2] <= 3 & table[j,x-1] > 50, 2,  0))
                                        
                                          }
                                      
                                          dat2[y,1] <- 0
                                          colnames(dat2)[1] <- "index"
                                          table <- cbind(table, dat, dat2)
                                          return(table)
                                }
              }) 
              
              CPM_PMR_MF2 <- reactive({if(is.null(CPM_PMR_MF())){return(NULL)
                                   } else {CPM <- CPM_PMR_MF()
                                           n <- ncol(CPM)
                                           CPM <- CPM[is.finite(rowSums(CPM[,3:((n-8)/2+2)])), ]
                                           return(CPM)
              }
              }) #without sales=0
              
              sales0 <- reactive({if(is.null(CPM_PMR_MF())){return(NULL)
                                } else {CPM <- CPM_PMR_MF()
                                        n <- ncol(CPM)
                                        CPM <- CPM[-nrow(CPM),]
                                        sales0 <- CPM[!is.finite(rowSums(CPM[,3:((n-8)/2+2)])), ]
                                        return(sales0)
              }
              })
              
              end_date <-  reactive(paste(as.yearmon(input$dateRange[2])))
              
              
              trips_com1 <- reactive({if(is.null(Com_PMR_MF())){return(NULL)
                                    } else {Com_PMR_MF <- Com_PMR_MF() 
                                            n <- ncol(Com_PMR_MF)
                                            trips1 <- Com_PMR_MF[which(Com_PMR_MF[,(n-1)]==1),]
                                            alert <- tail(head(colnames(trips1),(n-10)/2+5), 4)
                                            trips1 <- trips1[,colnames(trips1) %in% c("PMR", "MF", alert)]
                                            Month <- data.frame(Month = rep(end_date(), nrow(trips1)))
                                            trips1 <- cbind(Month, trips1)
                                            return(trips1)
                                }
                                })
                
              trips_com2 <- reactive({if(is.null(Com_PMR_MF())){return(NULL)
                                    } else {Com_PMR_MF <- Com_PMR_MF() 
                                            n <- ncol(Com_PMR_MF)
                                            trips2 <- Com_PMR_MF[which(Com_PMR_MF[,n]==1),]
                                            alert <- tail(head(colnames(trips2),(n-10)/2+8), 3)
                                            trips2 <- trips2[,colnames(trips2) %in% c("PMR", "MF", alert)]
                                            Month <- data.frame(Month = rep(end_date(), nrow(trips2)))
                                            trips2 <- cbind(Month, trips2)
                                            return(trips2)
              }
              })
              
              trips_CPM1 <- reactive({if(is.null(CPM_PMR_MF2())){return(NULL)
                                    } else {CPM_PMR_MF2 <- CPM_PMR_MF2() 
                                            n <- ncol(CPM_PMR_MF2)
                                            trips1 <- CPM_PMR_MF2[which(CPM_PMR_MF2[,(n-1)]==1),]
                                            alert <- tail(head(colnames(trips1),(n-10)/2+5), 4)
                                            trips1 <- trips1[,colnames(trips1) %in% c("PMR", "MF", alert)]
                                            Month <- data.frame(Month = rep(end_date(), nrow(trips1)))
                                            trips1 <- cbind(Month, trips1)
                                            return(trips1)
              }
              })
              
              trips_CPM2 <- reactive({if(is.null(CPM_PMR_MF2())){return(NULL)
                                    } else {CPM_PMR_MF2 <- CPM_PMR_MF2() 
                                            n <- ncol(CPM_PMR_MF2)
                                            trips2 <- CPM_PMR_MF2[which(CPM_PMR_MF2[,n]==1),]
                                            alert <- tail(head(colnames(trips2),(n-10)/2+8), 3)
                                            trips2 <- trips2[,colnames(trips2) %in% c("PMR", "MF", alert)]
                                            Month <- data.frame(Month = rep(end_date(), nrow(trips2)))
                                            trips2 <- cbind(Month, trips2)
                                            return(trips2)
              }
              })
              
              code_PMR <- reactive({if(is.null(Com_PMR_MF())){return(NULL)
                                  } else {switch(input$code_PMR,
                                                 MF_PMR = Com_PMR_MF(),
                                                 CPM_MF_PMR6 = data.frame(CPM_PMR_MF2())
              )}
              })
              
              
              output$table_PMR <- DT::renderDataTable(if(is.null(code_PMR())){return(NULL)
                                     } else {DT::datatable(code_PMR(),
                                             options = list(columnDefs = list(list(targets = seq((ncol(code_PMR())-10)/2+6+4, ncol(code_PMR())), visible = FALSE)),
                                                            pageLength = 15
                                                            )
                                                            ) %>% formatStyle(tail(head(colnames(code_PMR()), (ncol(code_PMR())-10)/2+5), (ncol(code_PMR())-10)/2),
                                                                              head(tail(colnames(code_PMR()),(ncol(code_PMR())-10)/2+1),(ncol(code_PMR())-10)/2),
                                                                              backgroundColor = styleEqual(c(0,1), c('white', 'yellow'))
                                                            ) %>% formatStyle(tail(head(colnames(code_PMR()), (ncol(code_PMR())-10)/2+8), 2),
                                                                              tail(colnames(code_PMR()),1),
                                                                              backgroundColor = styleEqual(c(0,1,2), c('white', 'red','orange'))
                )
              })
              
               

#---------------------------------------------------------------------------------------------------------------------------------------  
#                                       Plot 1 Complaints and CPM Trend Chart
#---------------------------------------------------------------------------------------------------------------------------------------
               
               
               plot1_data <- reactive({plot1_data <- CPM_MF6()})
               

               Trend_Chart1 <- function()({if(is.null(MF())|is.null(CPM_MF6())|is.null(Sales_6())){return(NULL)
                                         }else{validate(need(nrow(Sales_6()) >= 1, "No Sales in current PHRI/PMR Category/Date Range."))
                                               MF <- MF()
                                               CPM <- plot1_data()
                                               data <- data.frame(month = as.yearmon(colnames(MF)),
                                                                  Complaint = as.numeric(t(MF[nrow(MF),])),
                                                                  CPM = as.numeric(CPM[nrow(CPM),]))
                                               if(median(data$Complaint) == 0 | median(data$CPM) == 0){
                                                     b <- 1
                                               }
                                                else{
                                                     b <- median(data$Complaint)/median(data$CPM)
                                                }
                                               
                                               plot <- ggplot(data = data)+
                                                       geom_bar(stat = "identity",
                                                                aes(x = as.Date(month), y = Complaint), 
                                                                fill = "dodgerblue4",
                                                                width = 12
                                                                ) +
                                                      
                                                       geom_line(aes(x = as.Date(month), y = CPM * b),
                                                                 colour="red",
                                                                 lwd = 0.3
                                                                 ) + 
                   
                                                       geom_point(aes(x = as.Date(month), y = CPM * b),
                                                                  colour="red",
                                                                  size = 0.8
                                                                  ) + 
                                                       geom_text(aes(x = as.Date(month), y = CPM * b), 
                                                                 label = data$CPM,
                                                                 colour="red",
                                                                 size = 1,
                                                                 vjust = -1.5) +
                                                 
                                                       scale_x_date("", 
                                                                    date_breaks = "month", 
                                                                    date_labels = "%b-%Y") +
                                                 
                                                       scale_y_continuous(name = "Total Complaints", 
                                                                          sec.axis = sec_axis(~./ b, name = "CPM")) +
                                                    
                                                       ggtitle(paste("Complaint and CPM Analysis for", input$"BU_dropdown")) +
                                                 
                                                       theme(plot.title = element_text(size = 7, hjust = 0.5),
                                                             axis.text.x = element_text(angle = 90, hjust = 1),
                                                             axis.title.y.left = element_text(color = "dodgerblue4"),
                                                             axis.text.y.left = element_text(color = "dodgerblue4"),
                                                             axis.title.y.right = element_text(color = "red"),
                                                             axis.text.y.right = element_text(color = "red"),
                                                             text = element_text(size = 7),
                                                             legend.title=element_blank(),
                                                             legend.position="bottom"
                                                             )
                                                       return(plot)
                                                       }
               })
               
             
               

#---------------------------------------------------------------------------------------------------------------------------------------  
#                                       Plot 2 Complaint Analysis Chart
#--------------------------------------------------------------------------------------------------------------------------------------- 
               
               Plot2_data <- reactive({if(is.null(table_c())|is.null(input$MF)|is.null(input$`PHRI_dropdown`)|is.null(input$PMR_dropdown)|is.null(input$HC)){return(NULL)
                                   } else {table_c <- table_c()[(table_c()$`PHRI` %in% input$`PHRI_dropdown`) & (table_c()$`PMR Attribute Category` %in% input$PMR_dropdown),]
                                           validate(need(nrow(table_c) > 0, "No complaints in current date range and PHRI Category."))
                                           tb1<-aggregate(table_c$count, by=list(table_c$month,table_c$`PMR Attribute Category`),sum)
                                           colnames(tb1) <- c("month","PMR","count")
                                           dates <- seq(input$dateRange[1], input$dateRange[2], by = "month")
                                           months <- as.yearmon(dates)
                                           catg <- unique(tb1$PMR)
                                           tb <- as.data.frame(lapply(expand.grid(months,catg),unlist))
                                           colnames(tb) <- c("month","PMR")
                                           tb2 <- merge(tb, tb1, by= c('month','PMR'), all.x=T)
                                           tb2$count[is.na(tb2$count)] <- 0
                                           return(tb2)
                                           }
               })
               
               Trend_Chart2 <- function()({if(is.null(Plot2_data())){return(NULL)
                                      }else{data <- Plot2_data()

                                            plot <- ggplot()+
                                                    geom_line(data = data, 
                                                              aes(x = as.Date(month), y = count, group = PMR, color = PMR),
                                                              lwd = 0.5
                                                              ) + 
                                                    geom_point(data = data, 
                                                               aes(x = as.Date(month), y = count, group = PMR, color = PMR),
                                                               size = 0.5
                                                               ) + 
                                                    scale_x_date("", 
                                                                 date_breaks = "month", 
                                                                 date_labels = "%b-%Y") +
                                                    scale_y_continuous(name = "Complaints", breaks = int_breaks)+
                                                    ggtitle(paste("Complaint Analysis for",input$"BU_dropdown")) +
                                                    
                                                    theme(plot.title = element_text(size = 7, hjust = 0.5),
                                                          axis.text.x = element_text(angle = 90, hjust = 1),
                                                          text = element_text(size = 7),
                                                          legend.title = element_blank(),
                                                          legend.position="bottom",
                                                          legend.text = element_text(size = 3),
                                                          legend.key.size = unit(0.3,"line")
                                                           )
                                                    
                                                    return(plot)}
               })
               
#---------------------------------------------------------------------------------------------------------------------------------------  
#                                       Plot 3 CPM Analysis Chart
#--------------------------------------------------------------------------------------------------------------------------------------- 
  
               
               plot3_CPM <- reactive({if(is.null(table_c())|is.null(Sales())|is.null(input$MF)|is.null(input$`PHRI_dropdown`)|is.null(input$PMR_dropdown)|is.null(input$HC)){return(NULL)
                                   } else {table_c <- table_c()[(table_c()$`PHRI` %in% input$`PHRI_dropdown`) & (table_c()$`PMR Attribute Category` %in% input$PMR_dropdown),]
                                             validate(need(nrow(table_c) > 0, "No complaints in current date range and PHRI/PMR Category."))
                                             tb1 <- aggregate(table_c$count, by = list(table_c$`PMR Attribute Category`, table_c$month),sum) 
                                             colnames(tb1) <- c('PMR Attribute Category', 'month','count_c')
                                             tb_s <- Sales()[(Sales()$`PHRI` %in% input$`PHRI_dropdown`) & (Sales()$`PMR Attribute Category` %in% input$PMR_dropdown),]
                                             tb_s$month <- as.yearmon(tb_s$month)
                                             validate(need(nrow(tb_s) > 0, "No sales in current date range and PHRI/PMR Category."))
                                             
                                             dates <- seq(input$dateRange[1], input$dateRange[2], by = "month")
                                             months <- as.yearmon(dates)
                                             catg <- unique(tb_s$`PMR Attribute Category`)  
                                             full <- as.data.frame(lapply(expand.grid(months,catg),unlist))
                                             colnames(full) <- c("month","PMR Attribute Category")
                                             full <- merge(full, tb1, by= c('month',"PMR Attribute Category"), all.x=T)
                                             full$count[is.na(full$count)] <- 0
                                             tb1 <- full
                                             
                                  
                                             tb2 <- aggregate(tb_s$count, by = list(tb_s$`PMR Attribute Category`, tb_s$month),sum) 
                                             colnames(tb2) <- c('PMR Attribute Category', 'month','count_s')
                                             combined <- merge(tb1,tb2,by=c("month",  "PMR Attribute Category"))
                                             combined <- combined[,-3]
                                             colnames(combined) <- c( 'month','PMR Attribute Category',"complaints", "sales1")
                                             table <- data.frame()
                                             for(i in 1:6){
                                                           t1 <- combined
                                                           t1$month <- t1$month + (i/12)
                                                           table <- rbind(table,t1)
                                             }
                                             sales6 <- aggregate(table$sales, by = list(table$`PMR Attribute Category`, table$month), FUN = mean)
                                             colnames(sales6) <- c( 'PMR Attribute Category','month', "sales6")
                                             sales6 <- sales6[sales6$month >= as.yearmon(input$dateRange[1]) & sales6$month <= as.yearmon(input$dateRange[2]),]
                                             
                                             combined <- merge(combined, sales6, by = c("month",  "PMR Attribute Category"))
                                             colnames(combined) <- c( 'month','PMR Attribute Category',"complaints", "sales1","sales6")
                                             return(combined)
                                             }
               })
               
               plot3_CPM6 <- reactive({if(is.null(plot3_CPM())){return(NULL)
                                    } else {combined <- plot3_CPM()
                                            combined[,ncol(combined)+1] <- round(combined$complaints/combined$sales6*1000000, 2)
                                            colnames(combined) <- c( 'month','PMR',"complaints", "sales1","sales6", "CPM")
                                            return(combined)
                                            }
               })
                 
              
               
               plot3_data <- reactive({if(is.null(plot3_CPM6())){return(NULL)
                                    } else {plot3_data <- plot3_CPM6()
                                            return(plot3_data)
               }})
               
               Trend_Chart3 <- function()({if(is.null(plot3_data())){return(NULL)
                                         }else{data <- plot3_data()
               
                                               plot <- ggplot()+
                                                       geom_line(data = data, 
                                                                 aes(x = as.Date(month), y = CPM, group = PMR, color = PMR),
                                                                 lwd = 0.5
                                                                  ) + 
                                                       geom_point(data = data, 
                                                                  aes(x = as.Date(month), y = CPM, group = PMR, color = PMR),
                                                                      size = 0.5
                                                                  ) + 
                                                       scale_x_date("", 
                                                                    date_breaks = "month", 
                                                                    date_labels = "%b-%Y") +
                                                                    scale_y_continuous(name = "CPM", breaks = int_breaks)+
                                                                    ggtitle(paste("CPM Analysis for",input$"BU_dropdown")) +
                 
                                                       theme(plot.title = element_text(size = 7, hjust = 0.5),
                                                             axis.text.x = element_text(angle = 90, hjust = 1),
                                                             text = element_text(size = 7),
                                                             legend.title = element_blank(),
                                                             legend.position="bottom",
                                                             legend.text = element_text(size = 3),
                                                             legend.key.size = unit(0.3,"line")
                                                             )
                                                        
                                                 return(plot)}
               })
               
#---------------------------------------------------------------------------------------------------------------------------------------  
#                                          Outputs and Downloads
#--------------------------------------------------------------------------------------------------------------------------------------- 
               
               #output$result <- renderDataTable({Sales_1()})
               #output$result <- renderDataTable({plot1_data1()})
               
               #output$text <- renderText(if(is.null(tb())){return(NULL)
               #                         } else {paste("*Trending flagging:
               #                                        flagging on yellow for the 4th month if consecutive Complaints/CPM increases for any 3  months;
               #                                        flagging on red if the recent quarter average CPM > 3 and % Difference > 50%;
               #                                        flagging on orange if the recent quarter average CPM < or = 3 and % Difference > 50%")                  
               #})
               
               
               output$Trend_Chart1 <- renderPlot(Trend_Chart1(),height = 800, width = 1200, res = 300)
               output$Trend_Chart2 <- renderPlot(Trend_Chart2(),height = 900, width = 1300, res = 300)
               output$Trend_Chart3 <- renderPlot(Trend_Chart3(),height = 900, width = 1300, res = 300)
               
               
               output$download2 <- downloadHandler(filename = function() {paste0("Trips ", end_date(), ".xlsx",sep = "")},
                                                   content = function(file) {#write.xlsx(tb_d(), file, rowNames=T) #download excel file without highlighting
                                                   wb <- createWorkbook()
                                                   addWorksheet(wb = wb, sheet = "Complaint Alert 1")
                                                   writeData(wb = wb, sheet = 1, x = trips_com1(), rowNames = F)
                                                   addWorksheet(wb = wb, sheet = "Complaint Alert 2")
                                                   writeData(wb = wb, sheet = 2, x = trips_com2(), rowNames = F)
                                                   addWorksheet(wb = wb, sheet = "Complaint Table")
                                                   writeData(wb = wb, sheet = 3, x = Com_PMR_MF(), rowNames = F)
                                                   addWorksheet(wb = wb, sheet = "CPM Alert 1")
                                                   writeData(wb = wb, sheet = 4, x = trips_CPM1(), rowNames = F)
                                                   addWorksheet(wb = wb, sheet = "CPM Alert 2")
                                                   writeData(wb = wb, sheet = 5, x = trips_CPM2(), rowNames = F)
                                                   addWorksheet(wb = wb, sheet = "CPM Table")
                                                   writeData(wb = wb, sheet = 6, x = CPM_PMR_MF2(), rowNames = F)
                                                   addWorksheet(wb = wb, sheet = "No PMR in Sales")
                                                   writeData(wb = wb, sheet = 7, x = nosales(), rowNames = F)
                                                   addWorksheet(wb = wb, sheet = "Sales=0")
                                                   writeData(wb = wb, sheet = 8, x = sales0(), rowNames = F)
                                                     
                                                   yellow <- createStyle(fontColour = "#000000", fgFill = "#ffff00")
                                                   red <- createStyle(fontColour = "#000000", fgFill = "#ff0000")
                                                   orange <- createStyle(fontColour = "#000000", fgFill = "#ffa500")

                                                   n <- ncol(Com_PMR_MF())
                                                   nr <- nrow(Com_PMR_MF())

                                                   index1 <- Com_PMR_MF()[,((n-10)/2+10):(n-1)]
                                                   index2 <- Com_PMR_MF()[,n]
                                                   lapply(1:nr, function(i){
                                                      addStyle(wb = wb, sheet = 3, style = yellow,
                                                               rows = i + 1,
                                                               cols = which(index1[i,]=="TRUE") + 5
                                                       )
                                                     }
                                                     )

                                                     addStyle(wb = wb, sheet = 3, style = red,
                                                              rows = which(index2=="1") + 1,
                                                              cols = ((n-10)/2+7):((n-10)/2+8),
                                                              gridExpand = T,
                                                              stack = T
                                                     )

                                                     # addStyle(wb = wb, sheet = 3, style = orange,
                                                     #          rows = which(index2=="2") + 1,
                                                     #          cols = ((n-10)/2+7):((n-10)/2+8),
                                                     #          gridExpand = T,
                                                     #          stack = T
                                                     # )
                                                     
                                                     n_CPM <- ncol(CPM_PMR_MF2())
                                                     nr_CPM <- nrow(CPM_PMR_MF2())
                                                     
                                                     index1_CPM <- CPM_PMR_MF2()[,((n-10)/2+10):(n-1)]
                                                     index2_CPM <- CPM_PMR_MF2()[,n]
                                                     lapply(1:nr_CPM, function(i){
                                                       addStyle(wb = wb, sheet = 6, style = yellow,
                                                                rows = i + 1,
                                                                cols = which(index1_CPM[i,]=="TRUE") + 5
                                                       )
                                                     }
                                                     )
                                                     
                                                     addStyle(wb = wb, sheet = 6, style = red,
                                                              rows = which(index2_CPM=="1") + 1,
                                                              cols = ((n-10)/2+7):((n-10)/2+8),
                                                              gridExpand = T,
                                                              stack = T
                                                     )

                                                     deleteData(wb, sheet = 3, cols = ((n-10)/2+10):(n + 1), rows = 1:(nr+1), gridExpand = TRUE)
                                                     deleteData(wb, sheet = 6, cols = ((n-10)/2+10):(n + 1), rows = 1:(nr+1), gridExpand = TRUE)
                                                     deleteData(wb, sheet = 8, cols = ((n-10)/2+10):(n + 1), rows = 1:(nr+1), gridExpand = TRUE)
                                                     saveWorkbook(wb, file, overwrite = T)
                                                   })
               
               output$download <- downloadHandler(filename = function() {paste("table.xlsx",sep = "")},
                                                  content = function(file) {#write.xlsx(tb_d(), file, rowNames=T) #download excel file without highlighting
                                                    wb <- createWorkbook()
                                                    addWorksheet(wb = wb, sheet = "Sheet 1")
                                                    writeData(wb = wb, sheet = 1, x = tb1(), rowNames = T)
                                                
                                                    yellow <- createStyle(fontColour = "#000000", fgFill = "#ffff00")
                                                    red <- createStyle(fontColour = "#000000", fgFill = "#ff0000")
                                                    orange <- createStyle(fontColour = "#000000", fgFill = "#ffa500")
                                                    
                                                    n <- ncol(tb1())
                                                    nr <- nrow(tb1())
                                                    
                                                    index1 <- tb1()[,((n-8)/2+8):(n-1)]
                                                    index2 <- tb1()[,n]
                                                    lapply(1:nrow(tb1()), function(i){
                                                                                      addStyle(wb = wb, sheet = 1, style = yellow, 
                                                                                               rows = i + 1,
                                                                                               cols = which(index1[i,]=="TRUE") + 4
                                                                                               )
                                                                                     }
                                                    )
                                                    
                                                    addStyle(wb = wb, sheet = 1, style = red, 
                                                             rows = which(index2=="1") + 1,
                                                             cols = ((n-8)/2+6):((n-8)/2+7),
                                                             gridExpand = T, 
                                                             stack = T
                                                    )
                                                     
                                                    addStyle(wb = wb, sheet = 1, style = orange, 
                                                             rows = which(index2=="2") + 1,
                                                             cols = ((n-8)/2+6):((n-8)/2+7),
                                                             gridExpand = T, 
                                                             stack = T
                                                    )
                                                    
                                                    deleteData(wb, sheet = 1, cols = ((n-8)/2+9):(n + 1), rows = 1:(nr+1), gridExpand = TRUE)
                                                    saveWorkbook(wb, file, overwrite = T)
                                                    })
               
               output$download_plot1 <- downloadHandler(filename = "Complaints & CPM Trend Chart.png",
                                                       content = function(file) {png(file, height = 800, width = 1200, res = 300)
                                                                                 print(Trend_Chart1())
                                                                                 dev.off()
                                                                                 },
                                                       contentType = "image/png"
                                               
               ) 
               
               output$download_plot2 <- downloadHandler(filename = "Complaints Analysis Chart.png",
                                                        content = function(file) {png(file, height = 800, width = 1200, res = 300)
                                                                                  print(Trend_Chart2())
                                                                                  dev.off()
                                                        },
                                                        contentType = "image/png"
                                                        
               ) 
               
               output$download_plot3 <- downloadHandler(filename = "CPM Analysis Chart.png",
                                                        content = function(file) {png(file, height = 800, width = 1200, res = 300)
                                                                                  print(Trend_Chart3())
                                                                                  dev.off()
                                                       },
                                                       contentType = "image/png"
                                                       
               )
               
}

# Run the application 
shinyApp(ui = ui, server = server)