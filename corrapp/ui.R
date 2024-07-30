library(shiny)
library(shinythemes)
library(readxl)
library(tidyquant)
library(TTR)
library(stringr)
library(tidyverse)
library(ggcorrplot)
library(corrr)
library(xlsx)
library(devtools)
library(DT)
source_url("https://raw.githubusercontent.com/kassambara/r2excel/master/R/r2excel.r")
source_url("https://raw.githubusercontent.com/statisticsguru1/Additional-xlsx-functions/main/myfuns")


options(shiny.maxRequestSize=1000000*1024^2)
ui<-shinyUI(fluidPage(theme =shinytheme('flatly'),
                      navbarPage("Corrapp",theme = "flatlyST_bootstrap_CTedit.css",
                                 tabPanel(icon("house"),
                                          
                                          fluidRow(offset=1,
                                                   column(1,),
                                                   column(10,
                                                          p("This app features 6 tabs,each dedicated to carry out a specific task, they include;",
                                                            strong("Getting started,Upload data,Constituent Ranking,Performance based graphs,Correlation based graphs"),
                                                            "and",strong("Download full report"),
                                                            style="text-align:justify;color:white;background-color:black;padding:15px;border-radius:10px"),
                                                          br(),
                                                          
                                                          p(icon("house"), "- This tab outlines the app use instructions",
                                                            style="text-align:justify;color:black;background-color:lavender;padding:15px;border-radius:10px"
                                                          ),
                                                          
                                                          p(strong("Upload data"), "- This tab features data upload and display",
                                                            "Click",strong("browse"),"to upload an excel file with constituent lists and raw data",
                                                            "for instance, excel sheets with names; Lithium Lists,Uranium Lists,potassium Lists ... are understood as having constituent names, while the sheet Data is understood as having the data.",
                                                            "Under", strong("Display"),"choose to display either the processed data or daily percentage returns.",
                                                            "Under", strong("Select group"),"choose the group to display",
                                                            
                                                            style="text-align:justify;color:black;background-color:white;padding:15px;border-radius:10px"
                                                          ),
                                                          
                                                          p(strong("Constituent Ranking"), "- This tab has two choices,Performance based rankings and correlation based rankings. performance based rankings tab displays constituent ranking based on price performance, while correlation based rankings display constituent rankings based on average correlation total correlation",
                                                            "i.e how a constituent say 1MC prices correlating with the total prices of the other constituents on that group", "correlation here means that the price changes of that constituent has a potential of changing the prices of the whole group" ,
                                                            "this makes the constituent a good indicator of the whole group.",
                                                            style="text-align:justify;color:black;background-color:orange;padding:15px;border-radius:10px"
                                                          ),
                                                          
                                                          p(strong("performance based graphs"), "- This tab displays the performance based graphs which includes Price percentage oscillator, moving average, Top and quantile correlation.",
                                                            "from",strong("graph type"),"choose which of the three graph types to display",
                                                            "Percentage price oscillator will give a table with Moving Average Convergence Divergence(MACD), based on the top constituents you choose",
                                                            "by default you get an MACD for top 3 and top 6 constituent(you can know them from performance ranking table on the previous tab)", "you can choose more 'Top' constituent lines to add",
                                                            "you can adjust MACD inputs such as the slow and fast exponential moving average inputs",
                                                            "Moving average gives the moving average lines for the top constituents while top and bottom quintiles option gives the top bottom quantile days interms of performance",
                                                            "you can adjust the rolling widow as desired"," Select group helps choose the group to display",
                                                            style="text-align:justify;color:black;background-color:lavender;padding:15px;border-radius:10px"
                                                          ),
                                                          
                                                          p(strong("Correlation based graphs"), "- This tab displays correlation based graphs","you can choose the average correlation oscillator, which is an oscillator barcharts on the average-total correlation for top ranking groups",
                                                            "moving correlation are the n- day average rolling correlation lines for each top constituent group",
                                                            "Top quintile is the n day rolling top and bottom quintile correlations",
                                                            
                                                            style="text-align:justify;color:white;background-color:black;padding:15px;border-radius:10px"
                                                          ),
                                                          p(strong("Correlation based graphs"),"- As you can see from the other tabs, you could download the charts on each tab",
                                                            "this final tab offers an oppotunity to download all the analysis into a single excel sheet ",
                                                            style="text-align:justify;color:black;background-color:lavender;padding:15px;border-radius:10px"
                                                          )
                                                   ),
                                                   column(1,)
                                                   
                                          )),
                                 tabPanel("Upload data",
                                          sidebarPanel(width = 4,style="background-color:white;border-radius: 10px;border-color: black",
                                                       fluidRow(
                                                         br()
                                                       ),
                                                       fluidRow(offset=1,
                                                                fileInput("file1", "Browse file",
                                                                          accept = c(
                                                                            "text/csv",
                                                                            "text/comma-separated-values,text/plain",
                                                                            ".csv","xlsx","xls")),
                                                                br(),
                                                                selectInput("dtype", "Display",
                                                                            choices=c("Raw data","Returns"),
                                                                            selected = "Raw data"),
                                                                br(),
                                                                selectInput("inSelect", "Select group",
                                                                            choices=c(""),
                                                                            selected = "")
                                                                
                                                       ),
                                                       
                                                       
                                                       fluidRow(offset=1,
                                                                br(),
                                                                
                                                                fluidRow(
                                                                  column(2,offset=1,),
                                                                  column(4,offset=1,
                                                                         downloadButton("downloadData", "Download")),
                                                                  column(2,offset=1,)
                                                                )
                                                                
                                                       )
                                                       
                                          ),
                                          
                                          mainPanel(width = 8,
                                                    h4(tags$strong(textOutput('maindt'))),
                                                    tags$head(
                                                      tags$style(
                                                        HTML(".shiny-notification {
              height: 100px;
              width: 400px;
              position:fixed;
              top: calc(50% - 50px);;
              left: calc(80% - 400px);;
            }
           "
                                                        )
                                                      )
                                                    ),
                                                    DT::dataTableOutput("maindata"))
                                 ),
                                 navbarMenu("Constituent Ranking",
                                            tabPanel("Performance based ranking",
                                                     fluidRow(
                                                       column(2,
                                                              selectInput("inSelect1", "Select group",
                                                                          choices=c(""),
                                                                          selected ="")     
                                                       ),
                                                       column(8,),
                                                       column(2,
                                                              downloadButton("downloadData1", "Download rankings")
                                                              
                                                       )),
                                                     
                                                     fluidRow(h4(tags$strong(textOutput("perform"))),
                                                              DT::dataTableOutput("performanceranking")
                                                     )),
                                            
                                            tabPanel("Correlation based ranking",
                                                     fluidRow(
                                                       column(2,
                                                              selectInput("inSelect2", "Select group",
                                                                          choices=c(""),
                                                                          selected = "")
                                                              
                                                       )
                                                       ,
                                                       column(8,),
                                                       column(2,
                                                              downloadButton("downloadData2", "Download rankings")
                                                              
                                                       )
                                                     ),
                                                     fluidRow(
                                                       h4(tags$strong(textOutput("corr"))),
                                                       DT::dataTableOutput("correlationranking")
                                                     )
                                            )
                                 ),
                                 tabPanel("Performance based graphs",
                                          sidebarPanel(width = 4,style="background-color:white;border-radius: 10px;border-color: black",
                                                       fluidRow(
                                                         column(10,offset=1,
                                                                h4(tags$strong("Graph type input"),style="color:orange"),
                                                                
                                                         )     
                                                       ),
                                                       
                                                       fluidRow(
                                                         column(10,offset=1,
                                                                selectInput("pergraph","select graph type",
                                                                            choices=c("Percentage point oscillator","Moving Averages",
                                                                                      "Top and bottom quintile"),
                                                                            selected =""),
                                                         )),
                                                       fluidRow(
                                                         column(10,offset=1,
                                                                h4(tags$strong("Top constituents input"),style="color:orange"),
                                                                
                                                         )     
                                                       ),
                                                       fluidRow(
                                                         column(10,offset=1,
                                                                selectizeInput(
                                                                  "Tops"
                                                                  , "Enter top constituents to consider"
                                                                  , choices = seq(1,100,1),
                                                                  selected=c(3,6)
                                                                  , multiple = T
                                                                )
                                                         )),
                                                       
                                                       conditionalPanel(
                                                         condition = "input.pergraph == 'Percentage point oscillator'",
                                                         
                                                         fluidRow(
                                                           column(10,offset=1,
                                                                  h4(tags$strong("PP oscillator inputs"),style="color:orange"),
                                                           )     
                                                         ),
                                                         
                                                         fluidRow(offset=1,
                                                                  column(5,offset=1,
                                                                         numericInput("nFast","nFast",
                                                                                      value=3,min = 2,
                                                                                      step =1)),
                                                                  column(5,offset=1,
                                                                         numericInput("nSlow","nSlow",
                                                                                      value=6,min = 2,
                                                                                      step =1,width = '1200px'))),
                                                         
                                                         fluidRow(offset=1,
                                                                  column(5,offset=1,
                                                                         numericInput("nSig","nSig",
                                                                                      value=6,min = 2,
                                                                                      step =1)),
                                                                  column(5,offset=1,
                                                                         selectInput("percent","percent",
                                                                                     c("TRUE"=TRUE,"FALSE"=FALSE),
                                                                                     selected="TRUE",width = '1200px')))),
                                                       conditionalPanel(
                                                         condition = "input.pergraph != 'Percentage point oscillator'",
                                                         fluidRow(
                                                           column(10,offset=1,
                                                                  h4(tags$strong("Rolling window input"),style="color:orange"),
                                                           )     
                                                         ),
                                                         fluidRow(offset=1,
                                                                  column(10,offset=1,
                                                                         numericInput("roll","Rolling window",
                                                                                      value=6,min = 2,
                                                                                      step =1)))),
                                                       
                                                       
                                                       fluidRow(
                                                         column(10,offset=1,
                                                                h4(tags$strong("Constituent group input"),style="color:orange"),
                                                         )     
                                                       ),
                                                       fluidRow(offset=1,
                                                                column(10,offset=1,
                                                                       selectInput("inSelect3", "Select group",
                                                                                   choices=c(""),
                                                                                   selected =""))
                                                       ),
                                                       br()
                                          ),
                                          
                                          mainPanel(width = 8,
                                                    fluidRow(
                                                      column(4,
                                                             h4(tags$strong(textOutput("pps")))),
                                                      column(4,),
                                                      column(4,
                                                             downloadButton("downloadData3", "Download perfomance graphs")
                                                             
                                                      )),
                                                    tabsetPanel(
                                                      tabPanel("Data view",DT::dataTableOutput("contents")),
                                                      tabPanel("Graph view",shiny::plotOutput("contentsgraph")))
                                          )),
                                 tabPanel("correlation based graphs",
                                          sidebarPanel(width = 4,style="background-color:white;border-radius: 10px;border-color: black",
                                                       fluidRow(
                                                         column(10,offset=1,
                                                                h4(tags$strong("Graph type input"),style="color:orange"),
                                                                
                                                         )     
                                                       ),
                                                       
                                                       fluidRow(
                                                         column(10,offset=1,
                                                                selectInput("corrgraph","select graph type",
                                                                            choices=c("average correlation oscillator","average correlation linegraph",
                                                                                      "top and bottom quintile correlations"),
                                                                            selected ="Average correlation oscillator"),
                                                         )),
                                                       fluidRow(
                                                         column(10,offset=1,
                                                                h4(tags$strong("Top constituents input"),style="color:orange"),
                                                                
                                                         )     
                                                       ),
                                                       fluidRow(
                                                         column(10,offset=1,
                                                                selectizeInput(
                                                                  "Tops1"
                                                                  , "Enter top constituents to consider"
                                                                  , choices = seq(1,100,1),
                                                                  selected=c(3,6)
                                                                  , multiple = T
                                                                )
                                                         )),
                                                       fluidRow(
                                                         column(10,offset=1,
                                                                h4(tags$strong("Rolling window input"),style="color:orange"),
                                                         )     
                                                       ),
                                                       fluidRow(offset=1,
                                                                column(10,offset=1,
                                                                       numericInput("roll1","Rolling window",
                                                                                    value=6,min = 2,
                                                                                    step =1))),
                                                       
                                                       fluidRow(
                                                         column(10,offset=1,
                                                                h4(tags$strong("Constituent group input"),style="color:orange"),
                                                         )     
                                                       ),
                                                       fluidRow(offset=1,
                                                                column(10,offset=1,
                                                                       selectInput("inSelect4", "Select group",
                                                                                   choices=c(""),
                                                                                   selected =""))
                                                       ),
                                                       br()
                                          ),
                                          
                                          mainPanel(width = 8,
                                                    
                                                    fluidRow(
                                                      column(4,
                                                             h4(tags$strong(textOutput("pps1")))),
                                                      column(4,),
                                                      column(4,
                                                             downloadButton("downloadData4", "Download correlation graphs")
                                                             
                                                      )),
                                                    tabsetPanel(
                                                      
                                                      tabPanel("Data view",DT::dataTableOutput("contents1")),
                                                      
                                                      tabPanel("Graph view",shiny::plotOutput("contentsgraph1"))))
                                 ),
                                 tabPanel("Download full report",
                                          
                                          fluidRow(offset=1,),
                                          br(),
                                          br(),
                                          br(),
                                          br(),
                                          br(),
                                          br(),
                                          br(),
                                          br(),
                                          fluidRow(offset=1,
                                                   column(5,),
                                                   column(3,
                                                          downloadButton("downloadData5", "Download full report")
                                                   ),
                                                   column(4,)
                                          )
                                          
                                          
                                          
                                 )
                      )))

## server script


