#===========================================================================================================
#
#                                  Article 88 Shiny App
#                                 
#                                 Last Update: 3/11/2021
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
#install.packages("readxl")
#install.packages("openxlsx")
#install.packages("DT")
#install.packages("dplyr")
library(shiny)
library(shinythemes)
library(readxl)
library(openxlsx)
library(DT)
library(dplyr)
library('lubridate')
library(zoo)

options(shiny.maxRequestSize = 30*1024^10)  #change uploading file's size
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

#-----------------------------------------------------------------------------------------------------------
#
#                                         Define UI 
#
#-----------------------------------------------------------------------------------------------------------
ui <- navbarPage(
  
  theme = shinytheme("cerulean"),
  
  "Article 88", # App title
  
  tabsetPanel(
    
    tabPanel("Article 88",
             sidebarPanel(width=2,
                          
                          helpText(div(strong("Step 1:")), " Upload Master Complaints File (.xlsx).", style = "color:deepskyblue"),
                          fileInput("Complaints", "Master Complaint", multiple = FALSE, accept = c(".xlsx")),
                        
                          helpText(div(strong("Step 2:")), " Upload Sales Eaches File (.xlsx).", style = "color:deepskyblue"),
                          fileInput("Sales", "Sales Eaches", multiple = FALSE, accept = c(".xlsx")),
                          
                          helpText(div(strong("Step 3:")), " Upload PMS mapping File (.xlsx)", style = "color:deepskyblue"),
                          fileInput("PMS_mapping", "PMS Mapping File", multiple = FALSE, accept = c(".xlsx")),
                          
                          helpText(div(strong("Step 4:")), " Upload Influential Points  File (.xlsx)", style = "color:deepskyblue"),
                          fileInput("excl_list", "Exclusion List", multiple = FALSE, accept = c(".xlsx")),
                          
                          helpText(div(strong("Step 5:")), "Please Choose Target Information.", style = "color:deepskyblue"),
                          dateRangeInput('Range1', label = 'Range 1: Date Range for Control Chart: yyyy-mm-dd',
                                         start = "2019-1-01", end = "2020-12-31"),
                          helpText("(Date range need to be 24 months; Start date must be after July 2016.)"),
                          dateRangeInput('Range2', label = 'Range 2: Date Range for Most Recent Period: yyyy-mm-dd',
                                         start = "2020-7-01", end = "2020-12-31"),
                          
                          helpText(div(strong("Select Severity Rate")), style = "color:black"),
                          dropdownButton(label = "SR", status = "default",
                                         actionButton(inputId = "all SR", label = "(Un)select all"),
                                         verbatimTextOutput(outputId = "SR"),
                                         checkboxGroupInput(inputId = "SR_dropdown", label = "", choices = "Loading...")
                          ),
                          
                          helpText(div(strong("Select Business Unit")), style = "color:black"),
                          dropdownButton(label = "BU", status = "default",
                                         actionButton(inputId = "all BU", label = "(Un)select all"),
                                         verbatimTextOutput(outputId = "BU"),
                                         checkboxGroupInput(inputId = "BU_dropdown", label = "", choices = "Loading...")
                          ),
                          
                          helpText(div(strong("Select PMS Category")), style = "color:black"),
                          dropdownButton(label = "PMS", status = "default",
                                         actionButton(inputId = "all PMS", label = "(Un)select all"),
                                         verbatimTextOutput(outputId = "PMS"),
                                         checkboxGroupInput(inputId = "PMS_dropdown", label = "", choices = "Loading...")
                          ),
                          
                          helpText(div(strong("Select PHRI Category")),style = "color:black"),
                          dropdownButton(label = "PHRI", status = "default",
                                         actionButton(inputId = "all PHRI", label = "(Un)select all"),
                                         verbatimTextOutput(outputId = "PHRI"),
                                         checkboxGroupInput(inputId = "PHRI_dropdown", label = "", choices = "Please select BU and SR...")
                          ),
                          
                          helpText(strong("Select PMR Attribute Category"),style = "color:black"),
                          dropdownButton(label = "PMR", status = "default",
                                         actionButton(inputId = "all PMR", label = "(Un)select all"),
                                         verbatimTextOutput(outputId = "PMR"),
                                         checkboxGroupInput(inputId = "PMR_dropdown", label = "", choices = "Please select BU and PHRI...")
                          ),
                          
                          helpText(strong("Select Malfunction Code"),style = "color:black"),
                          dropdownButton(label = "MF", status = "default",
                                         actionButton(inputId = "all MF", label = "(Un)select all"),
                                         verbatimTextOutput(outputId = "MF"),
                                         checkboxGroupInput(inputId = "MF_dropdown", label = "", choices = "Please select PHRI and PMR...")
                          ),
                          
                          helpText(div(strong("Step 6:")), " Download Output Files.", style = "color:deepskyblue"),
                          downloadButton("control_chart", "Control Charts"),
                          downloadButton("control_chart_trips", "Tripped Control Charts"),
                          downloadButton("index", "Index Table")
             ),
             
             mainPanel(p("=================================================================================================================================="),
                                      
                       fluidRow(
                                column(5, h1("Article 88"),
                                       p("Copyright belong to ConvaTec Metrics Central team"),
                                       p("Updated Date: Mar 2021")),
                                       column(4, img(src = MCImage, height = 120,width = 400))
                                      ),
                       p("=================================================================================================================================="),
                       DT::dataTableOutput("result") 
             )
    )
)
)

#-----------------------------------------------------------------------------------------------------------
#
#                                         Define Server 
#
#-----------------------------------------------------------------------------------------------------------
server <- function(input, output, session) {
  
  
  raw <- reactive({if(is.null(input$Complaints)|is.null(input$PMS_mapping)){return(NULL)
                  } else {data <- read_excel(paste(input$Complaints$datapath), 1)
                          PMS_mapping_file <- read_excel(paste(input$PMS_mapping$datapath), 1)
                          data <- subset(data, select = c(`PR ID`, `Date Created`, `Business Unit (Hyperion)`, PHRI, `PMR Attribute Category`, MF, `Malfunction Code`, `HC`, `Part Number`, `Severity Rating`))
                          data$`Date Created` <- as.Date(data$`Date Created`,"%m %d %Y")
                          validate(need(nrow(data[is.na(data$`Business Unit (Hyperion)`),]) == 0, 
                                  "Warning: There are blanks in Business Unit"))
                          validate(need(nrow(data[is.na(data$`PHRI`),]) == 0, 
                                   "Warning: There are blanks in PHRI"))
                          validate(need(nrow(data[is.na(data$`PMR Attribute Category`),]) == 0, 
                                  "Warning: There are blanks in PMR Attribute Category"))
                          validate(need(nrow(data[is.na(data$`MF`),]) == 0, 
                                  "Warning: There are blanks in MF"))
                          validate(need(nrow(data[is.na(data$`HC`),]) == 0, 
                                 "Warning: There are blanks in HC"))
                          data[data$`PMR Attribute Category`=="Aquacel Ag Extra",]$`PMR Attribute Category` <- "Aquacel Ag EXTRA"
                          data[data$`PMR Attribute Category`=="Aloe VESTA Body Wash",]$`PMR Attribute Category` <- "Aloe Vesta Body Wash"
                          data[data$`PMR Attribute Category`=="Aloe VESTA Cleansers",]$`PMR Attribute Category` <- "ALOE VESTA Cleansers"
                          data[data$`PMR Attribute Category`=="Aloe VESTA Lotions & Creams",]$`PMR Attribute Category` <- "ALOE VESTA Lotions & Creams"
                          
                          data <- merge(data, PMS_mapping_file, by.x = "Part Number", by.y = "ICC", all.x = TRUE)
                          print(paste0("There are ", nrow(data[is.na(data$PMS),]), " complainst without PMS Plan #."))
                          data[is.na(data$PMS),]$PMS <- "TBD"
                          #print(data[data$`Business Unit (Hyperion)`=="Ostomy" & data$PMS == "CCC PMS Plan 007",])
                          return(data)}
  })
  
  data <- reactive({if(is.null(input$Complaints)|is.null(input$PMS_mapping)){return(NULL)
            } else {data <- raw()
                    data <- data[data$`Date Created` >= input$Range1[1] & data$`Date Created` <= input$Range1[2]+1,]
                    return(data)
            }
  })
  
  
  observeEvent(data(), {selected_SR <- sort(unique(data()$`Severity Rating`))
                       updateCheckboxGroupInput(session = session, inputId = "SR_dropdown", choices = selected_SR, selected = "")
  })
  
  observeEvent(input$"all SR", {selected_SR <- sort(unique(data()$"Severity Rating"))
                                if (is.null(input$"SR_dropdown")){updateCheckboxGroupInput(session = session, inputId = "SR_dropdown", selected = selected_SR)
                                                          } else {updateCheckboxGroupInput(session = session, inputId = "SR_dropdown", selected = "")
  }
  })  
  
  observeEvent(input$"SR_dropdown", {BU <- sort(unique(data()[data()$"Severity Rating" %in% input$"SR_dropdown",]$`Business Unit (Hyperion)`))
                                     updateCheckboxGroupInput(session = session, inputId = "BU_dropdown", choices = BU, selected = "")
  })
  
  observeEvent(input$"all BU", {BU <- sort(unique(data()[data()$"Severity Rating" %in% input$"SR_dropdown",]$`Business Unit (Hyperion)`))
                                if (is.null(input$"BU_dropdown")){updateCheckboxGroupInput(session = session, inputId = "BU_dropdown", selected = BU)
                                                          } else {updateCheckboxGroupInput(session = session, inputId = "BU_dropdown", selected = "")
  }
  })
  
  observeEvent(input$"BU_dropdown", {selected_PMS <- sort(unique(data()[data()$`Business Unit (Hyperion)` %in% input$"BU_dropdown" &
                                                                        data()$"Severity Rating" %in% input$"SR_dropdown",]$`PMS`))
                                    updateCheckboxGroupInput(session = session, inputId = "PMS_dropdown", choices = selected_PMS, selected = selected_PMS)
  })
  
  observeEvent(input$"all PMS", {selected_PMS <- sort(unique(data()[data()$`Business Unit (Hyperion)` %in% input$"BU_dropdown"&
                                                                    data()$"Severity Rating" %in% input$"SR_dropdown",]$`PMS`))
                                if (is.null(input$"PMS_dropdown")){updateCheckboxGroupInput(session = session, inputId = "PMS_dropdown", selected = selected_PMS)
                                                           } else {updateCheckboxGroupInput(session = session, inputId = "PMS_dropdown", selected = "")
  }
  })
  
  
  observeEvent(input$"PMS_dropdown", {selected_PHRI <- sort(unique(data()[data()$`Business Unit (Hyperion)` %in% input$"BU_dropdown" &
                                                                          data()$"Severity Rating" %in% input$"SR_dropdown" &
                                                                          data()$PMS %in% input$"PMS_dropdown",]$`PHRI`))
                                     updateCheckboxGroupInput(session = session, inputId = "PHRI_dropdown", choices = selected_PHRI, selected = selected_PHRI)
  })
  
  observeEvent(input$"all PHRI", {selected_PHRI <- sort(unique(data()[data()$`Business Unit (Hyperion)` %in% input$"BU_dropdown"&
                                                                      data()$"Severity Rating" %in% input$"SR_dropdown"&
                                                                      data()$PMS %in% input$"PMS_dropdown",]$`PHRI`))
                                  if (is.null(input$"PHRI_dropdown")){updateCheckboxGroupInput(session = session, inputId = "PHRI_dropdown", selected = selected_PHRI)
                                                              } else {updateCheckboxGroupInput(session = session, inputId = "PHRI_dropdown", selected = "")
  }
  })
  
  
  observeEvent(input$"PHRI_dropdown", {selected_PMR <- sort(unique(data()[data()$`Business Unit (Hyperion)` %in% input$"BU_dropdown" &
                                                                          data()$"PHRI" %in% input$"PHRI_dropdown"&
                                                                          data()$"Severity Rating" %in% input$"SR_dropdown" &
                                                                          data()$PMS %in% input$"PMS_dropdown",]$`PMR Attribute Category`))
                                       updateCheckboxGroupInput(session = session, inputId = "PMR_dropdown", choices = selected_PMR, selected = selected_PMR)
  })
  
  observeEvent(input$"all PMR", {selected_PMR <- sort(unique(data()[data()$`Business Unit (Hyperion)` %in% input$"BU_dropdown" &
                                                                    data()$"PHRI" %in% input$"PHRI_dropdown"&
                                                                    data()$"Severity Rating" %in% input$"SR_dropdown" &
                                                                    data()$PMS %in% input$"PMS_dropdown",]$`PMR Attribute Category`))
                                 if (is.null(input$"PMR_dropdown")){updateCheckboxGroupInput(session = session, inputId = "PMR_dropdown", selected = selected_PMR)
                                                            } else {updateCheckboxGroupInput(session = session, inputId = "PMR_dropdown", selected = "")
  }
  })
  
  observeEvent(input$"PMR_dropdown", {selected_MF <- sort(unique(data()[data()$`Business Unit (Hyperion)` %in% input$"BU_dropdown" &
                                                                        data()$"PHRI" %in% input$"PHRI_dropdown" &
                                                                        data()$"Severity Rating" %in% input$"SR_dropdown" &
                                                                        data()$`PMR Attribute Category` %in% input$"PMR_dropdown" &
                                                                        data()$PMS %in% input$"PMS_dropdown",]$`MF`))
                                      updateCheckboxGroupInput(session = session, inputId = "MF_dropdown", choices = selected_MF, selected = selected_MF)
  })
  
  observeEvent(input$"all MF", {selected_MF <- sort(unique(data()[data()$`Business Unit (Hyperion)` %in% input$"BU_dropdown" &
                                                                  data()$"PHRI" %in% input$"PHRI_dropdown"&
                                                                  data()$"Severity Rating" %in% input$"SR_dropdown" &
                                                                  data()$`PMR Attribute Category` %in% input$"PMR_dropdown" &
                                                                  data()$PMS %in% input$"PMS_dropdown",]$`MF`))
                                if (is.null(input$"MF_dropdown")){updateCheckboxGroupInput(session = session, inputId = "MF_dropdown", selected = selected_MF)
                                                          } else {updateCheckboxGroupInput(session = session, inputId = "MF_dropdown", selected = "")
  }
  })
  
  output$"BU" <- renderPrint(input$"BU_dropdown")
  output$"SR" <- renderPrint(input$"SR_dropdown")
  output$"PHRI" <- renderPrint(input$"PHRI_dropdown")
  output$"PMR" <- renderPrint(input$"PMR_dropdown")
  output$"MF" <- renderPrint(input$"MF_dropdown")
  
#---------------------------------------------------------------------------------------------------------------------------------------  
#                                       Read Data and Data Manipulations
#--------------------------------------------------------------------------------------------------------------------------------------- 
  
  Complaints <- reactive({if(is.null(data())|is.null(input$"SR_dropdown")|is.null(input$"BU_dropdown")|
                             is.null(input$"PHRI_dropdown")|is.null(input$"PMR_dropdown")|is.null(input$"MF_dropdown")){return(NULL)
                        } else {Complaints <- data()
                                Complaints$`Date Created` <- as.Date(Complaints$`Date Created`,"%m %d %Y")
                                Complaints <- Complaints[Complaints$`Date Created` >= input$Range1[1] & Complaints$`Date Created` <= input$Range1[2],]
                                Complaints <- Complaints[Complaints$`Severity Rating` %in% input$"SR_dropdown" &
                                                         Complaints$`Business Unit (Hyperion)` %in% input$"BU_dropdown" &
                                                         Complaints$PMS %in% input$"PMS_dropdown" &
                                                         Complaints$"PHRI" %in% input$"PHRI_dropdown" & 
                                                         Complaints$`PMR Attribute Category` %in% input$"PMR_dropdown" &
                                                         Complaints$`MF` %in% input$"MF_dropdown",]
                                return(Complaints)
  }
  })
  
  Complaints_noharm <- reactive({if(is.null(Complaints())){return(NULL)
                         } else {Complaints <- Complaints()
                                 Complaints <- Complaints[!is.na(Complaints$HC) & Complaints$HC!="TBD" & Complaints$HC!="No Harm"
                                                          & Complaints$HC!="HC09.01" & Complaints$HC!="HC09.02" & Complaints$HC!="HC09.03"
                                                          & Complaints$HC!="HC09.04" & Complaints$HC!="HC09.05" & Complaints$HC!="HC09.06"
                                                          & Complaints$HC!="HC09.07",]
  }})
  
  MF_fullname <- reactive({if(is.null(Complaints())){return(NULL)
                      } else {Complaint <- subset(Complaints(), select = c(`Business Unit (Hyperion)`, PHRI, `PMR Attribute Category`, MF, `Malfunction Code`))
                              list <- distinct(Complaint, `Business Unit (Hyperion)`, PHRI, `PMR Attribute Category`, MF, .keep_all= TRUE)
                              return(list)
  }})
  
  Highest_severity <- reactive({if(is.null(Complaints())){return(NULL)
                        } else {Complaint <- subset(raw(), select = c(`Date Created`, `Business Unit (Hyperion)`, PHRI, `PMR Attribute Category`, MF, `Severity Rating`))
                                Complaint <- Complaint[Complaint$`Date Created` >= (input$Range2[1]) & 
                                                       Complaint$`Date Created` <= (input$Range2[2]),] 
                                Complaint <- Complaint[order(-Complaint$`Severity Rating`),]
                                HS <- distinct(Complaint, `Business Unit (Hyperion)`, PHRI, `PMR Attribute Category`, MF, .keep_all= TRUE)
                                HS <- subset(HS, select = c(`Business Unit (Hyperion)`, PHRI, `PMR Attribute Category`, MF, `Severity Rating`))
                                return(HS)
  }})
  
  Highest_severity_ref <- reactive({if(is.null(raw())){return(NULL)
                            } else {Complaint <- subset(raw(), select = c(`Date Created`, `Business Unit (Hyperion)`, PHRI, `PMR Attribute Category`, MF, `Severity Rating`))
                                    Complaint <- Complaint[Complaint$`Date Created` < (input$Range2[1]),] 
                                    Complaint <- Complaint[order(-Complaint$`Severity Rating`),]
                                    HS <- distinct(Complaint, `Business Unit (Hyperion)`, PHRI, `PMR Attribute Category`, MF, .keep_all= TRUE)
                                    HS <- subset(HS, select = c(`Business Unit (Hyperion)`, PHRI, `PMR Attribute Category`, MF, `Severity Rating`))
                                    return(HS)
  }})
  
  PMS_plan <- reactive({if(is.null(Complaints())){return(NULL)
                } else {Complaint <- subset(raw(), select = c(`Business Unit (Hyperion)`, PHRI, `PMR Attribute Category`, MF, PMS))
                        Complaint <- Complaint[Complaint$PMS!="TBD" & !is.na(Complaint$PMS),]
                        PMS <- distinct(Complaint, `Business Unit (Hyperion)`, PHRI, `PMR Attribute Category`, MF, PMS, .keep_all= TRUE)
                        return(PMS)
  }})
  
  table_c <- function(file_name){
                          if(is.null(file_name)){return(NULL)
                          }else{
                          table_c <- file_name %>% 
                                     group_by(month = floor_date(`Date Created`,"month"), `Business Unit (Hyperion)`, `PHRI`, `PMR Attribute Category`, MF) %>% dplyr::summarize(count = n())
                          table_c$month <- as.yearmon(table_c$month)
                          return(table_c)}
  }
  
  Sales <- reactive({if(is.null(input$Sales)|is.null(input$PHRI_dropdown)|is.null(input$PMR_dropdown)){return(NULL)
                } else {
                        sales <- read_excel(paste(input$Sales$datapath), sheet = "Global Eaches", skip = 2)
                        
                        sales$`Hyperion Business Unit` <- ifelse(sales$`Hyperion Business Unit`=="Continence Care", 
                                                                "Total Continence Care",
                                                                 sales$`Hyperion Business Unit`)
                        
                        sales$`Hyperion Business Unit` <- ifelse(sales$`Hyperion Business Unit`=="Critical Care", 
                                                                "Total Critical Care",
                                                                 sales$`Hyperion Business Unit`)
                        
                        # sales$`Hyperion Business Unit` <- ifelse(sales$`Hyperion Business Unit`=="Wound",
                        #                                         "Wound Therapeutics",
                        #                                          sales$`Hyperion Business Unit`)
                        
                        list <- which(colnames(sales) %in% c('ICC (SR Code)', 'CCC Reference #', 'Hyperion Description', 'PMR Category', 'Total Eaches')) 
                        sales <- sales[,-list] 
                        colnames(sales)[4:ncol(sales)] <- as.character(seq(as.Date("2016-01-01"), by = "month", length.out = ncol(sales)-3))
                        #colnames(sales)[3:ncol(sales)] <- as.character(as.Date(as.numeric(colnames(sales)[3:ncol(sales)]), origin = "1899-12-30"))
                        sales[4:ncol(sales)] <- sapply(sales[4:ncol(sales)], as.character)
                        sales[is.na(sales)|sales=="-"|sales=="'0"] <- "0"
                        sales[4:ncol(sales)] <- sapply(sales[4:ncol(sales)], as.numeric)
                        sales <- sales[sales$'Hyperion Business Unit' %in% input$"BU_dropdown" &
                                       sales$'PHRI' %in% input$'PHRI_dropdown' &
                                       sales$'PMR Attribute Category' %in% input$'PMR_dropdown',]
                        validate(need(nrow(sales) > 0, "No Sales in current PHRI/PMR Category."))
    
                        tb_s <- data.frame()
                        for (i in 4:ncol(sales)){
                                      each <- aggregate(sales[i], by = list(sales$'Hyperion Business Unit', sales$`PMR Attribute Category`, sales$PHRI), sum)
                                      each$month <- as.Date(colnames(each)[4])
                                      colnames(each)[4] <- "count"
                                      tb_s <- rbind(tb_s,each)
                         }
                        colnames(tb_s)[1:3] <- c("BU", 'PMR', 'PHRI') 
                        tb_s <- tb_s[tb_s$month >= (input$Range1[1]%m-% months(6)) & 
                                     tb_s$month <= (input$Range1[2]),] #obtain data for previous 6 months
                      #write.xlsx(tb_s,"C:/Users/QI15403/Desktop/sales.xlsx")
                        return(tb_s)
  }
  })
  
  Sales_PMR_MF <- reactive({if(is.null(input$PHRI_dropdown)|is.null(input$PMR_dropdown)){return(NULL)
                 }else{table_BU <- data.frame()
                       Sales <- Sales()
                       Sales$month <- as.yearmon(Sales$month)
                       BU <- input$BU_dropdown

                       for(k in 1:length(BU)){   
                           PHRI <- unique(Sales[(Sales$`BU`==BU[k]),]$PHRI)
                           select_PHRI <- input$PHRI_dropdown
                           PHRI <- intersect(PHRI, select_PHRI)
                           table_PHRI <- data.frame()
    
                          for(j in 1:length(PHRI)){
                                PMR <- unique(Sales[(Sales$`BU`==BU[k]) & (Sales$PHRI==PHRI[j]),]$`PMR`)
                                select_PMR <- input$PMR_dropdown
                                PMR <- intersect(PMR, select_PMR)
      
                            
                               #MF - complaints by MF 
                                table <- Sales[(Sales$`BU`==BU[k]) & (Sales$`PHRI`==PHRI[j]),]
                                tb1 <- aggregate(table$count, by = list(table$month, table$`PMR`), sum)
                                colnames(tb1) <- c("month", "PMR", "count")
                                dates <- seq(input$Range1[1]%m-% months(6), input$Range1[2], by = "month")
                                months <- as.yearmon(dates)
                                catg <- PMR
                                tb <- as.data.frame(lapply(expand.grid(months, catg), unlist))
                                colnames(tb) <- c("month", "PMR")
                                tb2 <- merge(tb, tb1, by= c('month','PMR'), all.x = T)
                                tb2 <- tb2[order(tb2$'month'),]
                                tb2$count[is.na(tb2$count)] <- 0
                                table_PMR <- data.frame(matrix(tb2$count, ncol = length(unique(tb2$month))))
                                colnames(table_PMR) <- as.character(unique(tb2$month))
                                rownames(table_PMR) <- NULL
                                table_PMR$PMR <- unique(tb2$PMR)
                                table_PMR$PHRI <- rep(PHRI[j], nrow(table_PMR))
                                table_PMR$BU <- rep(BU[k], nrow(table_PMR))   
                                table_PHRI <- rbind(table_PHRI, table_PMR)     
                          }
    
                     table_BU <- rbind(table_BU, table_PHRI) 
  }
  
                     table_BU <- cbind(table_BU$BU, table_BU$PHRI, table_BU$PMR, table_BU[,1:(ncol(table_BU)-3)])
                     colnames(table_BU) <- c("BU", "PHRI", "PMR", as.character(unique(tb2$month)))
                     table_BU$BU <- as.character(table_BU$BU)
                     table_BU$PHRI <- as.character(table_BU$PHRI)
                     table_BU$PMR <- as.character(table_BU$PMR)
                     return(table_BU)
  }
  }) 
  
  
  Com_PMR_MF <- function(filename){
                        table_BU <- data.frame()
                        table_c <- table_c(filename)
                        BU <- input$BU_dropdown
                          
                             
                        for(k in 1:length(BU)){   
                          PHRI <- unique(table_c[(table_c$`Business Unit (Hyperion)`==BU[k]),]$PHRI)
                          select_PHRI <- input$PHRI_dropdown
                          PHRI <- intersect(PHRI, select_PHRI)
                          table_PHRI <- data.frame()
                          
                           for(j in 1:length(PHRI)){
                              PMR <- unique(table_c[(table_c$`Business Unit (Hyperion)`==BU[k]) & (table_c$PHRI==PHRI[j]),]$`PMR Attribute Category`)
                              select_PMR <- input$PMR_dropdown
                              PMR <- intersect(PMR, select_PMR)
                              table_PMR <- data.frame()
                              
                                  for(i in 1:length(PMR)){
                                           #MF - complaints by MF 
                                           table <- table_c[(table_c$`Business Unit (Hyperion)`==BU[k]) &
                                                            (table_c$`PHRI`==PHRI[j]) & 
                                                            (table_c$`PMR Attribute Category`==PMR[i]),]
                                           tb1 <- aggregate(table$count, by = list(table$month, table$MF), sum)
                                           colnames(tb1) <- c("month", "MF", "count")
                                           dates <- seq(input$Range1[1], input$Range1[2], by = "month")
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
                                           MF$PHRI <- rep(PHRI[j], nrow(MF))
                                           MF$BU <- rep(BU[k], nrow(MF))
                                           table_PMR <- rbind(table_PMR, MF)    
                                  }
                              
                             table_PHRI <- rbind(table_PHRI, table_PMR)     
                            
                          }
                          
                          table_BU <- rbind(table_BU, table_PHRI) 
                        }
                             
                          table_BU <- cbind(table_BU$BU, table_BU$PHRI, table_BU$PMR, table_BU$MF, table_BU[,1:(ncol(table_BU)-4)])
                          colnames(table_BU) <- c("BU", "PHRI", "PMR", "MF", as.character(unique(tb2$month)))
                          table_BU$BU <- as.character(table_BU$BU)
                          table_BU$PHRI <- as.character(table_BU$PHRI)
                          table_BU$MF <- as.character(table_BU$MF)
                          table_BU$PMR <- as.character(table_BU$PMR)
                          return(table_BU)
                       }

  
  nosales <- reactive({if(is.null(input$PHRI_dropdown)){return(NULL)
                    }else{sales <- subset(Sales(), select = c("BU", "PHRI", "PMR"))
                          sales <- unique(sales[c("BU", "PHRI", "PMR")])
                          sales$match <- rep("Yes", nrow(sales))
                          tb_s <- subset(table_c(Complaints()), select = c("Business Unit (Hyperion)", "PHRI", "PMR Attribute Category"))
                          tb_s <- unique(tb_s[c("Business Unit (Hyperion)", "PHRI", "PMR Attribute Category")])
  
                          all <- merge(tb_s, sales, 
                                       by.x = c("Business Unit (Hyperion)", "PHRI", "PMR Attribute Category"), 
                                       by.y = c("BU", "PHRI", "PMR"), 
                                       all.x = T)
                          unmatched <- all[is.na(all$match),1:3]
                          return(unmatched)
   }}) #no corresponding sales in Sales Eaches File
  
  CPM_PMR_MF <- reactive({if(is.null(input$PHRI_dropdown)|is.null(input$PMR_dropdown)|is.null(Sales())|is.null(table_c(Complaints()))){return(NULL)
                       }else{
                             sales <- Sales()[Sales()$month!=unique(Sales()$month)[length(unique(Sales()$month))],] # delete most recent month 
                             table_c <- table_c(Complaints())
                             table_BU <- data.frame()
                             BU <- input$BU_dropdown    
                             
                             for(k in 1:length(BU)){   
                               PHRI <- unique(table_c[(table_c$`Business Unit (Hyperion)`==BU[k]),]$PHRI)
                               select_PHRI <- input$PHRI_dropdown
                               PHRI <- intersect(PHRI, select_PHRI)
                               table_PHRI <- data.frame()
                               
                               for(j in 1:length(PHRI)){
                                   PMR <- unique(table_c[(table_c$`Business Unit (Hyperion)`==BU[k]) & (table_c$PHRI==PHRI[j]),]$`PMR Attribute Category`)
                                   select_PMR <- input$PMR_dropdown
                                   PMR <- intersect(PMR, select_PMR)
                                   PMR <- PMR[!(PMR %in% nosales()$`PMR Attribute Category`)]
                                   #validate(need(!is.na(PMR), "Please select at least one PMR category from each selected PHRI."))
                                   if(length(PMR)==0){
                                   
                                     table_PMR <- data.frame()
                                     
                                   }else{
                                     
                                     table_PMR <- data.frame()
    
                                     for(i in 1:length(PMR)){

                                     #Sales_6  
                                     Sales_s <- sales[(sales$'BU'==BU[k]) & (sales$`PHRI`==PHRI[j]) & (toupper(sales$`PMR`)==toupper(PMR[i])),]
                                     tb1 <- Sales_s[,colnames(Sales_s) %in% c("month","count")]
                                     tb2 <- data.frame(matrix(tb1$count, ncol = length(unique(tb1$month))))
                                     colnames(tb2) <- unique(as.character(as.yearmon(tb1$month)))
                                     mean <- data.frame()
                                     n <- ncol(tb2)
                                     for(l in 6:n){mean[1,(l-5)] <- round(rowMeans(tb2[,(l-5):l]),2)}
                                     Sales_6 <- mean
     
                                     #MF - complaints by MF
                                     table_mf <- table_c[(table_c$`Business Unit (Hyperion)`==BU[k]) & (table_c$`PHRI`==PHRI[j]) & (table_c$`PMR Attribute Category`==PMR[i]),]
                                     tb1 <- aggregate(table_mf$count, by = list(table_mf$month, table_mf$MF), sum)
                                     colnames(tb1) <- c("month", "MF", "count")
                                     dates <- seq(input$Range1[1], input$Range1[2], by = "month")
                                     months <- as.yearmon(dates)
                                     catg <- unique(tb1$MF)
                                     tb <- as.data.frame(lapply(expand.grid(months,catg),unlist))
                                     colnames(tb) <- c("month", "MF")
                                     tb2 <- merge(tb, tb1, by = c('month','MF'), all.x=T)
                                     tb2$count[is.na(tb2$count)] <- 0
                                     MF <- data.frame(matrix(tb2$count, ncol = length(unique(tb2$month))))
                                     colnames(MF) <- as.character(unique(tb2$month))
                                     rownames(MF) <- NULL
                                     MF$MF <- unique(tb2$MF)
      
                                     #CPM_MF6 
                                     CPM_MF <- as.matrix(MF[,1:(ncol(MF)-1)])
                                     Sales <- as.matrix(Sales_6)
                                     CPM_MF <- t(t(CPM_MF)/Sales[1,])
                                     CPM_MF <- round(CPM_MF*1000000, 2)
                                     CPM_MF <- data.frame(CPM_MF)
                                     CPM_MF$MF <- unique(tb2$MF)
                                     CPM_MF$PMR <- rep(PMR[i], nrow(MF))
                                     CPM_MF$PHRI <- rep(PHRI[j], nrow(MF))
                                     CPM_MF$BU <- rep(BU[k], nrow(MF))
                                     table_PMR <- rbind(table_PMR, CPM_MF)
                           
                              }
                                   }
                                   table_PHRI <- rbind(table_PHRI, table_PMR)
                           
                               }
                                   table_BU <- rbind(table_BU, table_PHRI)
                          
                             }
                       

                                   table_BU <- cbind(table_BU$BU, table_BU$PHRI, table_BU$PMR, table_BU$MF, table_BU[,1:(ncol(table_BU)-4)])
                                   colnames(table_BU) <- c("BU", "PHRI", "PMR", "MF", as.character(unique(tb2$month)))
                                   table_BU$MF <- as.character(table_BU$MF)
                                   table_BU$PMR <- as.character(table_BU$PMR)
                                   table_BU$PHRI <- as.character(table_BU$PHRI)
                                   table_BU$BU <- as.character(table_BU$BU)
                                   return(table_BU)
  }
  }) # supporting data
  
  CPM_PMR_MF2 <- reactive({if(is.null(CPM_PMR_MF())){return(NULL)
                      } else {CPM <- CPM_PMR_MF()
                              n <- ncol(CPM)
                              CPM <- CPM[is.finite(rowSums(CPM[,(5:n)])) & apply(CPM[,(5:n)], 1, function(x)all(x>=0)),] 
                              return(CPM)
  }
  }) #without sales <= 0
  
  # CPM_PMR_MF3 <- reactive({if(is.null(CPM_PMR_MF2())){return(NULL)
  #                     } else {CPM <- CPM_PMR_MF2()
  #                             n <- ncol(CPM)
  #                             CPM <- CPM[rowSums(CPM[,((n-5):n)])!=0,] 
  #                             return(CPM)
  # }
  # }) # exclude rows with most recent 6 months complaints=0
  
  sales0 <- reactive({if(is.null(CPM_PMR_MF())){return(NULL)
                 } else {CPM <- CPM_PMR_MF()
                         n <- ncol(CPM)
                         #CPM <- CPM[-nrow(CPM),]
                         sales0 <- CPM[!is.finite(rowSums(CPM[,(5:n)])),]
                         return(sales0)
  }
  })  # sales = 0
  
  sales_neg <- reactive({if(is.null(CPM_PMR_MF())){return(NULL)
                 } else {CPM <- CPM_PMR_MF()
                         n <- ncol(CPM)
                         sales_neg <- CPM[apply(CPM[,(5:n)], 1, function(x)all(x<0)),] 
                         return(sales_neg)
  }
  })  # sales < 0
  
  index_tb <- reactive({if(is.null(CPM_PMR_MF2())){return(NULL)
                     }else{index_tb <- subset(CPM_PMR_MF2(), select=c("BU", "PHRI", "PMR", "MF"))
                           index_tb$Page <- c(1:nrow(index_tb))
                        
                           CPM <- CPM_PMR_MF2()
                           Complaint <- Com_PMR_MF(Complaints())
                           Complaint_noharm <- Com_PMR_MF(Complaints_noharm())
                           Sales <- Sales()   
                           Sales_PMR_MF <- Sales_PMR_MF()
                           excl_list <- read_excel(paste(input$excl_list$datapath), 1)
                           excl_list$Month <- ifelse(nrow(excl_list)!=0, 
                                                     as.character(format(excl_list$Month, "%b %Y")),
                                                     excl_list$Month)
                      
                        
                           for(i in 1:nrow(CPM)){
                             BU <- CPM[i,]$BU
                             PHRI <- CPM[i,]$PHRI
                             PMR <- CPM[i,]$PMR
                             MF <- CPM[i,]$MF
                             
                             # 6M avg Sales
                             Sales_s <- Sales[Sales$'BU'==BU & 
                                              Sales$`PHRI`==PHRI & 
                                              toupper(Sales$`PMR`)==toupper(PMR),]
                             Sales_6 <- data.frame()
                             Sales_s <- Sales_s[-nrow(Sales_s),] # delete most recent month
                             n <- nrow(Sales_s)
                             for(l in 6:n){Sales_6[1,(l-5)] <- round(mean(Sales_s$count[(l-5):l]), 2)}
                             
                             # Complaints
                             Complaint_n <- Complaint[Complaint$BU==BU & 
                                                      Complaint$PHRI==PHRI & 
                                                      toupper(Complaint$PMR)==toupper(PMR) & 
                                                      Complaint$MF==MF,]
                             
                             Complaint_noharm_n <- Complaint_noharm[Complaint_noharm$BU==BU & 
                                                                    Complaint_noharm$PHRI==PHRI & 
                                                                    toupper(Complaint_noharm$PMR)==toupper(PMR) & 
                                                                    Complaint_noharm$MF==MF,]
                             if(nrow(Complaint_noharm_n)==0){Complaint_noharm_n <- Complaint_n
                                                             Complaint_noharm_n[1,5:ncol(Complaint_noharm_n)] <- rep(0, (ncol(Complaint_noharm_n)-4))
                                                       }else{Complaint_noharm_n <- Complaint_noharm_n} # if no corresponding complaints, assign as zeros.
                            
                             # Sales
                             Sales_n <- Sales_PMR_MF[Sales_PMR_MF$BU==BU & 
                                                     Sales_PMR_MF$PHRI==PHRI & 
                                                     toupper(Sales_PMR_MF$PMR)==toupper(PMR),]
                      
                             plot_data <- data.frame("Month"= as.character(format(seq(input$Range1[1], by = "month", length.out = ncol(Sales_6)),"%b %Y")),
                                                     "Complaint" = as.data.frame(t(Complaint_n[1,-c(1:4)])),
                                                     "Complaint_noharm" = as.data.frame(t(Complaint_noharm_n[1,-c(1:4)])),
                                                     "Sales" = as.data.frame(t(Sales_n[1,-c(1:9)])),
                                                     "Sales6" = as.data.frame(t(Sales_6)),
                                                     "CPM" = as.data.frame(t(CPM[i,-c(1:4)])))
                             rownames(plot_data) <- NULL
                             colnames(plot_data) <- c("Month", "Complaint", "Complaint_noharm", "Sales", "Sales6", "CPM")
                             
                             # Calculation
                             cl <- sum(plot_data$Complaint)/sum(plot_data$Sales6/1000000)
                             sd <- sqrt(cl/(plot_data$Sales6/1000000))
                             ucl5 <- cl + sd*5
                             
                             Month <- c()
                     
                             if(nrow(excl_list)==0){
                               
                                 Month <- Month
                                 
                             } else{
                               
                                 for(h in 1:nrow(excl_list)){
                               
                                    if(BU==excl_list$BU[h] & PHRI==excl_list$PHRI[h] & PMR==excl_list$PMR[h] & MF==excl_list$MF[h]){
                                      
                                               Month[length(Month) + 1] <- excl_list$Month[h]
                                   
                                      }else{Month <- Month}
                                 }
                       
                             }
                     
                             
                             if(length(Month)==0){
                                cl <- sum(plot_data[which(plot_data$CPM <= ucl5),]$Complaint)/sum(plot_data[which(plot_data$CPM <= ucl5),]$Sales6/1000000)
                                 # only exclude outliers from calculation (above 5 sd. UCL)
                              }else{
                                cl <- sum(plot_data[which(plot_data$CPM <= ucl5 & !plot_data$Month %in% Month),]$Complaint)/
                                      sum(plot_data[which(plot_data$CPM <= ucl5 & !plot_data$Month %in% Month),]$Sales6/1000000)
                                 # excluse points on the list file and above 5 sd.
                              } 
                             
                             
                             sd <- sqrt(cl/(plot_data$Sales6/1000000))
                             
                             # 2 out 3 consecutive points between UCL2 and UCL3
                             dataset <- plot_data$CPM
                             index <- ifelse(dataset >= cl + sd * 2 & dataset < cl + sd * 3, 1, 0) 
                             index2 <- NA
                             for(j in 1:(length(index) - 2)){
                               index2[j] <- ifelse(sum(index[j], index[j + 1], index[j + 2]) >= 2, 1, 0)
                             }
                             index2 <- which(index2==1) + 2 # location of 3rd point for "2 out 3 consecutive points greater than a fixed value "
                             consecutive_index <- unique(c(index2 - 2, index2 - 1, index2))
                             greater_index <- which(index==1)
                             highlight <- consecutive_index[consecutive_index %in% greater_index]
                             
                             # above ULC3
                             alert <- ifelse(dataset > cl + sd * 3, 1, 0)
                             
                             alert[which(is.na(alert))] <- 0 # 6M avg sales < 0 ==> CPM is NA ==> assign alert equal to 0
                            
                             # influential point
                             influential <- ifelse(dataset > ucl5, 1, 0)
                             excl_point <- ifelse(plot_data$Month %in% Month, 1, 0)
                             influential[which(is.na(influential))] <- 0
                             
                             # 3 or more complaints !=0 in the most recent 6M
                             zero_complaints <- ifelse(plot_data$Complaint=="0", 0, 1)
                             
                             # changes in complaints
                             cplt_changes <- c()
                             cplt <- plot_data$Complaint
                             for(h in 4:length(cplt)){
                               previous_3M_avg <- mean(c(cplt[h-1], cplt[h-2], cplt[h-3]))
                               #cplt_changes[h-3] <- scales::percent(round((cplt[h] - previous_3M_avg)/previous_3M_avg, 4))
                               cplt_changes[h-3] <- round((cplt[h] - previous_3M_avg)/previous_3M_avg, 3) * 100
                             }
                             
                             # changes in sales
                             sales_changes <- c()
                             sales <- plot_data$Sales6
                             for(g in 2:length(sales)){
                               sales_changes[g-1] <- round(6*(sales[g] - sales[g-1])/sales[g-1], 3) * 100
                             }
                             
                             # changes in first 3M and most recent 3M CPM
                             CPM_list <- plot_data$CPM
                             CPM_n <- length(plot_data$CPM)
                             CPM_3M1 <- round((mean(CPM_list[(CPM_n-5):(CPM_n-3)])-mean(CPM_list[(CPM_n-17):(CPM_n-6)]))/mean(CPM_list[(CPM_n-17):(CPM_n-6)]), 3) * 100
                             CPM_3M2 <- round((mean(CPM_list[(CPM_n-2):(CPM_n)])-mean(CPM_list[(CPM_n-14):(CPM_n-3)]))/mean(CPM_list[(CPM_n-14):(CPM_n-3)]), 3) * 100
                             #CPM_3M1 <- ifelse(is.infinite(CPM_3M1), "NA", CPM_3M1)
                             #CPM_3M2 <- ifelse(is.infinite(CPM_3M2), "NA", CPM_3M2)
                             range2_avg <- round(mean(CPM_list[(CPM_n-5):(CPM_n)]), 2) # most recent 6M CPM avg
                             range3_avg <- round(mean(CPM_list[(CPM_n-17):(CPM_n-6)]), 2) # 12M CPM avg prior to range 2 
                             range2_total_sales <- round(sum(plot_data$Sales[(CPM_n-5):(CPM_n)]), 2)
                             range2_total_complaints <- round(sum(plot_data$Complaint[(CPM_n-5):(CPM_n)]), 2)
                             range2_total_complaints_noharm <- round(sum(plot_data$Complaint_noharm[(CPM_n-5):(CPM_n)]), 2)
                               
                             # index for Range 1
                             index_tb$`Range1: Out of Control (>3sd)`[i] <- ifelse(sum(alert) > 0, "Yes", "No")
                             index_tb$`Range1: Extreme Outliers (>5sd)`[i] <- ifelse(sum(influential) > 0, "Yes", "No")
                             index_tb$`Range1: 2/3 Successive Points between 2sd line and UCL`[i] <- ifelse(sum(highlight) > 0, "Yes", "No")
                             index_tb$`Range1: Influential Points (Ad-hoc)`[i] <- ifelse(sum(excl_point) > 0, "Yes", "No")
                             
                             # index for Range 2
                             Range1_n <- length(seq(from = input$Range1[1], to = input$Range1[2], by='month'))
                             Range2_n <- length(seq(from = input$Range2[1], to = input$Range2[2], by='month'))
                             index_tb$`Range2: Out of Control (>3sd) Trips`[i] <- ifelse(sum(alert[(Range1_n-(Range2_n - 1)):Range1_n]) > 0, "Yes", "No")
                             index_tb$`Range2: Extreme Outliers (>5sd)`[i] <- ifelse(sum(influential[(Range1_n-(Range2_n - 1)):Range1_n]) > 0, "Yes", "No")
                             index_tb$`Range2: 2/3 Successive Points between 2sd line and UCL`[i] <- ifelse(any(highlight > (Range1_n - Range2_n)), "Yes", "No")
                             index_tb$`Range2: Influential Points (Ad-hoc)`[i] <- ifelse(sum(excl_point[(Range1_n-(Range2_n - 1)):Range1_n]) > 0, "Yes", "No")
                             index_tb$`Range2: > or = 3 Non-zero Complaint Months`[i] <- ifelse(sum(zero_complaints[(Range1_n-(Range2_n - 1)):Range1_n]) >= 3, "Yes", "No")
                             #index_tb$`Range2: any complaints > or = 3`[i] <- ifelse(any(plot_data$Complaint[(Range1_n-(Range2_n - 1)):Range1_n] >= 3), "Yes", "No")
                             cplt_changes <- cplt_changes[(Range1_n-(Range2_n - 1)-3):(Range1_n-3)]
                             sales_changes <- sales_changes[(Range1_n-Range2_n):(Range1_n-1)]
                             index_tb$`Range2: % of complaint changes - first tripped signal compared to the avg of previous 3M`[i] <- ifelse(index_tb$`Range2: Out of Control (>3sd) Trips`[i]=="No", NA,
                                                                                                                cplt_changes[head(which(alert[(Range1_n-(Range2_n - 1)):Range1_n]==1), 1)])
                             index_tb$`Range2: % of sales changes on the first tripped signal`[i] <- ifelse(index_tb$`Range2: Out of Control (>3sd) Trips`[i]=="No", NA,
                                                                                                            sales_changes[head(which(alert[(Range1_n-(Range2_n - 1)):Range1_n]==1), 1)])
                             index_tb$`Range2: % of CPM changes - first 3M compared to previous 12M`[i] <- CPM_3M1
                             index_tb$`Range2: % of CPM changes - most recent 3M compared to previous 12M`[i] <- CPM_3M2
                             index_tb$`Range2: CPM Avg (current 6 months) `[i] <- range2_avg
                             index_tb$`Range3: CPM Avg (prev 12 months prior to Range 2)`[i] <- range3_avg
                             index_tb$`Range2: total sales eaches for specific PMR Cat`[i] <- range2_total_sales
                             index_tb$`Range2: # complaints for MF codes tripped`[i] <- range2_total_complaints
                             index_tb$`Range2: P1 (prob of MF occurring) = V/U`[i] <- round(range2_total_complaints/range2_total_sales, 10)
                             index_tb$`Range2: # complaints with mf code (only with harm codes)`[i] <- range2_total_complaints_noharm 
                             index_tb$`Range2: P2 (prob of haz leading to harm) = X/V`[i] <- round(range2_total_complaints_noharm/range2_total_complaints, 10)
                             index_tb$`Range2: P0 (prob of harm occuring) = W*Y`[i] <- round((range2_total_complaints/range2_total_sales) * 
                                                                                       (range2_total_complaints_noharm/range2_total_complaints), 10) # need modify
                             index_tb$`Range2: CPM/PPM conversion`[i] <- round(index_tb$`Range2: P0 (prob of harm occuring) = W*Y`[i] * 1000000, 4)
                           }
                           
                           index_tb <- merge(index_tb, Highest_severity_ref(), 
                                             by.x = c("BU", "PHRI", "PMR", "MF"), 
                                             by.y = c("Business Unit (Hyperion)", "PHRI", "PMR Attribute Category", "MF"), 
                                             all.x = T)
                           
                           index_tb <- merge(index_tb, Highest_severity(), 
                                             by.x = c("BU", "PHRI", "PMR", "MF"), 
                                             by.y = c("Business Unit (Hyperion)", "PHRI", "PMR Attribute Category", "MF"), 
                                             all.x = T)
                           index_tb <- merge(index_tb, PMS_plan(), 
                                             by.x = c("BU", "PHRI", "PMR", "MF"), 
                                             by.y = c("Business Unit (Hyperion)", "PHRI", "PMR Attribute Category", "MF"), 
                                             all.x = T)
                           
                           index_tb <- index_tb[order(index_tb$Page),]
                           names(index_tb)[names(index_tb)=="Severity Rating.x"] <- "Highest Severity Rating before Most Recent 6M"
                           names(index_tb)[names(index_tb)=="Severity Rating.y"] <- "Highest Severity Rating in Most Recent 6M"
                           index_tb[is.na(index_tb$PMS),]$PMS <- "TBD"
                           
                           return(index_tb)
  }})
  

  
  Highlight <- reactive({if(is.null(index_tb())){return(NULL)
                     }else{highlight <- index_tb()[index_tb()[, 10]=="Yes" & index_tb()[,14]=="Yes", ]
                           return(highlight)
  }
  })
  
  
    
  makePdf <- function(filename, select){
                   pdf(filename, width = 11 , height = 8.5)
    
                   CPM <- merge(select[,colnames(select) %in% c("BU", "PHRI", "PMR", "MF")], CPM_PMR_MF2(), by = c("BU", "PHRI", "PMR", "MF"), all.x = T, sort = F)
    
                   # CPM <- CPM_PMR_MF2()
                   Complaint <- Com_PMR_MF(Complaints())
                   Sales <- Sales()
                   Sales <- Sales[Sales$month!=unique(Sales$month)[length(unique(Sales$month))],] # delete most recent month
                   excl_list <- read_excel(paste(input$excl_list$datapath), 1)
                   excl_list$Month <- ifelse(nrow(excl_list)!=0, 
                                             as.character(format(excl_list$Month, "%b %Y")),
                                             excl_list$Month)
                   MF_fullname <- MF_fullname()
  
                   for(i in 1:nrow(CPM)){
                       BU <- CPM[i,]$BU
                       PHRI <- CPM[i,]$PHRI
                       PMR <- CPM[i,]$PMR
                       MF <- CPM[i,]$MF

                       # Sales
                       Sales_s <- Sales[Sales$'BU'==BU & 
                                        Sales$`PHRI`==PHRI & 
                                        toupper(Sales$`PMR`)==toupper(PMR),]
                       Sales_6 <- data.frame()
                       n <- nrow(Sales_s)
                       for(l in 6:n){Sales_6[1,(l-5)] <- round(mean(Sales_s$count[(l-5):l]), 2)}
                     
                       neg_loc <- which(Sales_6[1,]<0)
                       Sales_6[1, neg_loc] <- mean(as.numeric(Sales_6[1,])) # set negative 6M avg equal to the avg
  
                      # Complaints
                       Complaint_n <- Complaint[Complaint$BU==BU & 
                                                Complaint$PHRI==PHRI & 
                                                toupper(Complaint$PMR)==toupper(PMR) & 
                                                Complaint$MF==MF,]
          
                       plot_data <- data.frame("Month"= as.character(format(seq(input$Range1[1], by = "month", length.out = ncol(Sales_6)),"%b %Y")),
                                               "Complaint" = as.data.frame(t(Complaint_n[1,-c(1:4)])),
                                               "Sales" = as.data.frame(t(Sales_6)),
                                               "CPM" = as.data.frame(t(CPM[i,-c(1:4)])))
                       rownames(plot_data) <- NULL
                       colnames(plot_data) <- c("Month", "Complaint", "Sales", "CPM")
                       
                       # Calculation
                       cl <- sum(plot_data$Complaint)/sum(plot_data$Sales/1000000)
                       sd <- sqrt(cl/(plot_data$Sales/1000000))
                       
                       ucl5 <- cl + sd*5 
                       
                       Month <- c()
                       
                       if(nrow(excl_list)==0){
                         
                         Month <- Month
                         
                       } else{
                         
                         for(h in 1:nrow(excl_list)){
                           
                           if(BU==excl_list$BU[h] & PHRI==excl_list$PHRI[h] & PMR==excl_list$PMR[h] & MF==excl_list$MF[h]){
                             
                             Month[length(Month) + 1] <- excl_list$Month[h]
                             
                           }else{Month <- Month}
                         }
                         
                       }
                       
                       
                       if(length(Month)==0){
                         cl <- sum(plot_data[which(plot_data$CPM <= ucl5),]$Complaint)/sum(plot_data[which(plot_data$CPM <= ucl5),]$Sales/1000000)
                         # exclude outliers from calculation (above 5 sd. UCL)
                       }else{
                         cl <- sum(plot_data[which(plot_data$CPM <= ucl5 & !plot_data$Month %in% Month),]$Complaint)/
                           sum(plot_data[which(plot_data$CPM <= ucl5 & !plot_data$Month %in% Month),]$Sales/1000000)
                         # excluse points on the list file and above 5 sd.
                       } 
                       
                       sd <- sqrt(cl/(plot_data$Sales/1000000)) # exclude outliers from calculation (above 5 sd. UCL)
                       
                       
                       ucl3 <- list("x" = 1:nrow(plot_data), "y" = cl + sd*3 )
                       ucl2 <- list("x" = 1:nrow(plot_data), "y" = cl + sd*2 )
                       ucl1 <- list("x" = 1:nrow(plot_data), "y" = cl + sd*1 )
                       
                       # 2 out 3 consecutive points between UCL2 and UCL3
                       dataset <- plot_data$CPM
                       index <- ifelse(dataset >= cl + sd * 2 & dataset < cl + sd * 3, 1, 0) 
                       index2 <- NA
                       for(j in 1:(length(index) - 2)){
                         index2[j] <- ifelse(sum(index[j], index[j + 1], index[j + 2]) >= 2, 1, 0)
                       }
                       index2 <- which(index2==1) + 2 # location of 3rd point for "2 out 3 consecutive points greater than a fixed value "
                       
                       consecutive_index <- unique(c(index2 - 2, index2 - 1, index2))
                       greater_index <- which(index==1)
                       highlight <- consecutive_index[consecutive_index %in% greater_index]
                       
                       # above ULC3
                       alert <- ifelse(dataset > cl + sd * 3, 1, 0)
                     
                       # influential point
                       influential <- ifelse(dataset > ucl5, 1, 0)
                       excl_point <- ifelse(plot_data$Month %in% Month, 1, 0)
                       
                       
                       ucl5 <- list("x" = 1:nrow(plot_data), "y" = cl + sd*5 )
                       
                       # MF full name
                       MF_des <- MF_fullname[MF_fullname$`Business Unit (Hyperion)`==BU & MF_fullname$PHRI==PHRI & 
                                             toupper(MF_fullname$`PMR Attribute Category`)==toupper(PMR) & MF_fullname$MF==MF,]$`Malfunction Code`
                       
                       # Plot
                       par(oma=c(5, 3, 3, 3)) # outer margin area
                       plot(1:nrow(plot_data), plot_data$CPM, 
                            type = "l", 
                            xlab = "", ylab = "", xaxt = "n", main = paste0("U Chart - ", PMR, " (", MF, ") CPM"),
                            ylim = c(min(0, min(plot_data$CPM)), 
                                     max(max(cl + sd*5) * 1.05, max(plot_data$CPM) * 1.05)))
                       par(bg = 'white',  bty = "L", font.axis = 1.2)
                       abline(h = cl, lwd = 2, col = "orange")
                       lines(ucl5, lty = 1, cex= 1, lwd = 2, col = "darkred")
                       lines(ucl3, lty = 1, cex= 1, lwd = 2, col = "red")
                       lines(ucl2, lty = 2, cex= 1, lwd = 2, col = "orange")
                       lines(ucl1, lty = 2, cex= 1, lwd = 2, col = "orange")
                       mtext(side = 4, text = "Central Line", at = cl, col = "orange", las = 1, cex= 0.8)
                       mtext(side = 4, text = "UCL", at = mean(cl + sd[length(sd)]*3), col = "red", las = 1, cex= 0.8)
                       mtext(side = 4, text = "+5 sd", at = mean(cl + sd[length(sd)]*5), col = "darkred", las = 1, cex= 0.8)
                       mtext(side = 4, text = "+2 sd", at = mean(cl + sd[length(sd)]*2), col = "orange", las = 1, cex= 0.8)
                       mtext(side = 4, text = "+1 sd", at = mean(cl + sd[length(sd)]*1), col = "orange", las = 1, cex= 0.8)
                       mtext(side = 2, line = 3, "CPM", col = "black", font = 2, cex = 1.2)
                       points(x = 1:nrow(plot_data), y = plot_data$CPM, pch=20, cex = 1.5)
                       points(x = (1:nrow(plot_data))[which(alert==1)],
                              y = plot_data[which(alert==1),]$CPM, col = "red", pch = 20)
                       points(x = (1:nrow(plot_data))[highlight],
                              y = plot_data[highlight,]$CPM, col = "orange", pch = 20)
                       points(x = (1:nrow(plot_data))[which(influential==1)],
                              y = plot_data[which(influential==1),]$CPM, col = "blue", lwd = 1.5, pch = 4, cex = 2)
                       points(x = (1:nrow(plot_data))[which(excl_point==1)],
                              y = plot_data[which(excl_point==1),]$CPM, col = "orange", lwd = 1.5, pch = 4, cex = 2)
                       axis(1, at = plot_data$Month, labels = F)
                       text(1:nrow(plot_data), plot_data$CPM, labels = paste0(plot_data$CPM, "(", plot_data$Complaint, ")"), col = "black", cex = 0.6, pos = 3)
                       text(1:nrow(plot_data), par("usr")[3], srt = 45, label = plot_data$Month, xpd = TRUE, cex = 0.7)
                       legend(x="bottomleft", 
                              legend=c("In Control", "Out of Control (> 3sd)", "2/3  Successive Points between 2sd and UCL", "Extreme Outliers (> 5sd)", "Influential Points (Ad-hoc)"),
                              pch = c(20, 20, 20, 4, 4), 
                              col = c('black', 'red', "orange", 'blue', "orange"), title = "", 
                              text.width=c(0, 2, 3, 4.5, 4.7),
                              cex = 0.7, xpd=TRUE, horiz = T, 
                              pt.cex = 1, inset=c(0,-0.1), bty="n", 
                              y.intersp = 0.8, x.intersp = 0.8)
                       mtext(paste0(MF_des) , side = 1 , outer = T, line = 0, cex= 0.7)
                       mtext(paste0("Page ", i) , side = 1 , outer = T, line = 3)
                        }
                  dev.off()
  }
  
  
  
  
  
  output$control_chart <- downloadHandler(
                          filename = paste0("Control Chart_", format(input$Range1[2],"%B %Y"), ".pdf"),
                          function(theFile){makePdf(theFile, CPM_PMR_MF2())}
  )
  
  output$control_chart_trips <- downloadHandler(
                                filename = paste0("Control Chart_Trips", format(input$Range1[2],"%B %Y"), ".pdf"),
                                function(theFile){makePdf(theFile, Highlight())}
  )
    
  output$result <- renderDataTable({index_tb()})
  
  output$index <- downloadHandler(filename = function() {paste("Index Table_", input$BU_dropdown, " ", format(input$Range1[2],"%B %Y"), ".xlsx",sep = "")},
                                         content = function(file) {
                                           
                                           wb <- createWorkbook()
                                           addWorksheet(wb = wb, sheet = "Index")
                                           writeData(wb = wb, sheet = "Index", x =  index_tb())
                                           addWorksheet(wb = wb, sheet = "Negative CPM")
                                           writeData(wb = wb, sheet = "Negative CPM", x =  sales_neg())
                                           addWorksheet(wb = wb, sheet = "6M AVG Sales=0")
                                           writeData(wb = wb, sheet = "6M AVG Sales=0", x =  sales0())
                                           addWorksheet(wb = wb, sheet = "No Sales Data")
                                           writeData(wb = wb, sheet = "No Sales Data", x =  nosales())
                                           addWorksheet(wb = wb, sheet = "Complaint")
                                           writeData(wb = wb, sheet = "Complaint", x =  Com_PMR_MF(Complaints()))
                                           addWorksheet(wb = wb, sheet = "Sales")
                                           writeData(wb = wb, sheet = "Sales", x =  Sales_PMR_MF())
                                           addWorksheet(wb = wb, sheet = "CPM")
                                           writeData(wb = wb, sheet = "CPM", x =  CPM_PMR_MF())
                                          
                                           # yellow <- createStyle(fontColour = "#000000", fgFill = "#ffff00")
                                           # red <- createStyle(fontColour = "#000000", fgFill = "#ff0000")
                                           # orange <- createStyle(fontColour = "#000000", fgFill = "#ffa500")
                                           # 
                                           # n <- ncol(index_tb())
                                           # nr <- nrow(index_tb())
                                           # 
                                           # 
                                           # index <- c()
                                           # for (i in 1:nr){
                                           #   index[i] <- ifelse(index_tb()[i, 10]=="Yes" & index_tb()[i,14]=="Yes" &
                                           #                     (index_tb()[i, 17] > 50|index_tb()[i, 18] > 50),
                                           #                     "Yes", "No"
                                           # 
                                           #   )
                                           # }
                                           # 
                                           # addStyle(wb = wb, sheet = "Index", style = red,
                                           #          rows = which(index=="Yes") + 1,
                                           #          cols = 1:18,
                                           #          gridExpand = T,
                                           #          stack = T
                                           # )
                                           # 
                                           
                                           
                                           withProgress(message = paste0("Downloading Index Table"),
                                                        {for (i in 1:10) {incProgress(1/10, detail = paste(i*10, "%"))
                                                          Sys.sleep(0.1)}
                                                        })
                                           
                                           saveWorkbook(wb, file, overwrite = T)
                                           
                                         })
  
}




# Run the application 
shinyApp(ui = ui, server = server)

