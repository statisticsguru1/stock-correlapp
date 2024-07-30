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
library(foreach)
source_url("https://raw.githubusercontent.com/kassambara/r2excel/master/R/r2excel.r")
source_url("https://raw.githubusercontent.com/statisticsguru1/Additional-xlsx-functions/main/myfuns")

options(shiny.maxRequestSize=1000000*1024^2)

server <- function(input, output, session){
  
  #### updating select inputs according to the loaded data
  
  observe({
    file <- input$file1
    if (is.null(file)){
      x <- character(0)
      selectsize<-100
    }else if(tools::file_ext(file$datapath)=="xlsx"|tools::file_ext(file$datapath)=="xls"){
      
      sheets<-excel_sheets(file$datapath)
      constituent_lists<-sheets[endsWith(sheets,"List")]
      x<-constituent_lists
      asset_groups<-lapply(constituent_lists, read_excel, path =file$datapath,col_names = F)
      names(asset_groups)<-constituent_lists
      selectsize<-asset_groups%>%map(nrow)%>%bind_rows()%>%max()
      
    }else{
      x <- character(0)
      selectsize<-100
    }
    
    updateSelectInput(session, "inSelect",
                      label = paste("Select group"),
                      choices = x,
                      selected = head(x, 1)
    )
    
    updateSelectInput(session, "inSelect1",
                      label = paste("Select group"),
                      choices = x,
                      selected = head(x, 1)
    )
    
    updateSelectInput(session, "inSelect2",
                      label = paste("Select group"),
                      choices = x,
                      selected = head(x, 1)
    )
    
    updateSelectInput(session, "inSelect3",
                      label = paste("Select group"),
                      choices = x,
                      selected = head(x, 1)
    )
    
    updateSelectInput(session, "inSelect4",
                      label = paste("Select group"),
                      choices = x,
                      selected = head(x, 1)
    )
    
    updateSelectizeInput(session, "Tops",
                         label = paste("Enter top constituents to consider"),
                         choices = seq(1,selectsize,1),
                         selected=c(3,6)
    )
    updateSelectizeInput(session, "Tops1",
                         label = paste("Enter top constituents to consider"),
                         choices = seq(1,selectsize,1),
                         selected=c(3,6)
    )
  })
  
  ### reading uploaded data
  
  datainput<-reactive({
    withProgress(message = 'Preparing data',
                 detail = 'This may take a while...', value = 0, {
                   for (i in 1:10) {
                     incProgress(1/10)
                     Sys.sleep(0.25)
                   }
                   
                   
                   file1 <- input$file1
                   if (is.null(input$file1)){
                     return(NULL)
                   }else{
                     my_filepath<-file1$datapath
                     
                     sheets<-excel_sheets(my_filepath)
                     constituent_lists<-sheets[endsWith(sheets,"List")]
                     
                     asset_groups<-lapply(constituent_lists, read_excel, path =my_filepath,col_names = F)
                     names(asset_groups)<-constituent_lists
                     
                     
                     
                     
                     # read main data
                     
                     ## names
                     
                     name<-Data <- read_excel(my_filepath,sheet = "Data",n_max=1,skip = 3)%>%
                       select(-1)%>%
                       select(!starts_with("..."))%>%
                       names()
                     
                     re<-foreach(i=1:length(name),.combine=c)%do%{
                       paste(c("Date","close"),name[i],sep="_") 
                     }
                     
                     ## Read data
                     
                     Data <- read_excel(my_filepath,sheet = "Data",skip = 4)%>%
                       setNames(str_squish(colnames(.)))%>%
                       select(-1)%>%
                       setNames(re)
                     
                     # map longest date
                     long<-(Data%>%
                              select(names(sort(colSums(!is.na(Data%>%
                                                                 select(contains("Date")))),decreasing =T)[1]))%>%
                              as.data.frame())[,1]
                     
                     ad<-foreach(i=1:length(name),.combine = rbind)%do%{
                       dta<-Data%>%select(paste(c("Date","close"),name[i],sep="_"))%>%
                         setNames(c("Date","close"))%>%
                         filter(!(is.na(Date)&is.na(close)))%>%
                         mutate(Constituents=name[i])
                       if(sum(!is.na(dta$Date))<sum(!is.na(dta$close))){
                         dta$Date<-long[1:length(dta$Date)]
                       }
                       dta
                     }
                     
                     Data<-ad%>%
                       mutate( Constituents=gsub("\\..*", "",  Constituents))%>%
                       distinct()%>% 
                       pivot_wider( names_from =Constituents,values_from =close)%>%
                       mutate(Date=as.Date(Date))%>%
                       arrange(Date)
                     
                     # split it by asset group and save each asset group data separately(constituents data)
                     
                     asset_group_data<-list()
                     for (asset_group in constituent_lists){
                       js<-asset_groups[[asset_group]]
                       asset_group_data[[asset_group]]<- Data%>%select(Date,js$`...1`)
                     }
                     return(asset_group_data)
                   }
                 })
  })
  
  
  ## rendering main data header according to the selection
  
  output$maindt <- renderText({
    
    if(input$inSelect!=""){
      out<-paste(strsplit(input$inSelect," ")[[1]][1],input$dtype)
    }else{
      out<-input$dtype
    }
    out
  })
  
  ## rendering performance based ranking header according to the selection
  
  output$perform <- renderText({
    if(input$inSelect1!=""){
      out1<-paste(strsplit(input$inSelect1," ")[[1]][1],"Rankings")
    }else{
      out1<-"Rankings"
    }
    out1
    
  })
  
  ## rendering correlation based ranking header according to the selection
  
  output$corr <- renderText({
    if(input$inSelect2!=""){
      out2<-paste(strsplit(input$inSelect2," ")[[1]][1],"Rankings")
    }else{
      out2<-"Rankings"
    }
    out2})
  
  ## rendering moving average header according to the selection
  
  output$pps <- renderText({
    if(input$inSelect3!=""){
      if(input$pergraph=="Moving Averages"){
        out3<-paste(strsplit(input$inSelect3," ")[[1]][1],input$roll,"day moving Averages")
      }else{
        out3<-paste(strsplit(input$inSelect3," ")[[1]][1],input$pergraph)
      }
    }else{
      out3<-"percentage price oscillators"
    }
    out3})
  
  ## rendering moving average header according to the selection 
  
  output$pps1 <- renderText({
    if(input$inSelect4!=""){
      out4<-paste(strsplit(input$inSelect4," ")[[1]][1],input$roll1,"day",input$corrgraph)
    }else{
      out4<-"rolling correlation"
    }
    out4})
  
  
  ## getting returns 
  
  returns<-reactive({
    asset_group_data<-datainput()
    
    returns<-asset_group_data%>%
      map(get_return)
    returns
  })
  
  ## render main data
  
  output$maindata<-DT::renderDataTable({
    
    if(input$dtype=="Raw data"){
      datalist<-datainput()
    }else{
      datalist<-returns()
    }
    display<-datalist[[input$inSelect]]
    
    datatable(display,
              options=list(
                scrollX = TRUE,   ## enable scrolling on X axis
                scrollY = TRUE,   ## enable scrolling on Y axis
                autoWidth = TRUE,
                rownames = FALSE,
                initComplete = JS(
                  "function(settings, json) {",
                  "$('td').css({'border': '1px solid black'});",
                  "$('th').css({'border': '1px solid black'});",
                  "}")))
  })
  
  ## Compute performance based rankings
  
  perfomancebasedranking<-reactive({
    asset_group_data<-datainput()
    
    asset_group_data%>%
      map(get_performance_rankings)
  })
  
  ## render performance ranking
  
  
  output$performanceranking<-DT::renderDataTable({
    assetlist<-perfomancebasedranking()
    assetlist[[input$inSelect1]]
  })
  
  
  ## get correlation based rankings
  
  correlationbasedrankings<-reactive({
    asset_group_data<-datainput()
    assetlist<-asset_group_data%>%map(rankfun)
    assetlist
  })
  
  ## render correlation based ranking
  
  output$correlationranking<-DT::renderDataTable({
    assetlist<-correlationbasedrankings()
    assetlist[[input$inSelect2]]
  })
  
  
  
  # get TPS
  
  TPs<-reactive({
    Tops<-as.numeric(input$Tops)
    asset_group_data<-datainput()
    performance_rankings<-perfomancebasedranking()
    
    
    TPs<-list()
    for(dat in names(asset_group_data)){
      data<-asset_group_data[[dat]]
      rankings<-performance_rankings[[dat]]
      
      for(num in Tops){
        data<-data%>%
          rowwise()%>%
          mutate(!!as.name(paste("Top_",num,sep="")):=mean(c_across(rankings$Constituent[1:num]),na.rm=T))
      }
      
      TPs[[dat]]<-data%>%select(Date,contains("Top"))%>%
        mutate_all( ~replace(., is.nan(.), NA))%>%
        as.tibble()%>%
        ungroup()
      # percentage price oscillators 
    }
    TPs
  })
  
  ## compute macds
  
  macds<-reactive({
    TPs<-TPs()
    nFast<-input$nFast
    nSlow<-input$nSlow
    nSig<-input$nSig
    percent<-input$percent
    
    macds<-list()
    for(dat in names(TPs)){
      data<-TPs[[dat]]
      
      macds[[dat]]<-data%>%
        mutate(across(.cols=contains("Top"),
                      .fns =pp,
                      nFast = nFast,
                      nSlow = nSlow,
                      nSig = nSig,
                      percent = percent))
    }
    macds
  })
  
  ## generate percentage point oscillators graphs
  
  ppgraphss<-reactive({
    ppgraphs<-list()
    macds<-macds()
    Tops<-as.numeric(input$Tops)
    roll<-input$roll
    
    for(dat in names(macds)){
      data<-macds[[dat]]
      ppgraphs[[dat]]<-data%>%
        pivot_longer(cols=contains("Top_"),names_to="series",values_to="macd")%>%
        mutate(series=fct_relevel(series,paste("Top_",Tops,sep="")))%>%
        mutate(pos=macd>=0,Date=as.Date(Date))%>%
        ggplot(aes(x=Date, y=macd,fill=pos))+
        geom_histogram(stat = 'identity')+
        labs(title=paste(strsplit(dat," ")[[1]][1],"Constituents"),y = 'percentage price oscillator')+
        theme_bw()+
        theme(legend.position = 'None',axis.text.x = element_text(angle = 90))+
        facet_wrap(~series,ncol=2)+
        scale_x_date(date_breaks = "1 month")
    }
    ppgraphs
  })
  
  ## moving average data
  
  moving_avg_data<-reactive({
    TPs<-TPs()
    roll<-input$roll
    TPs%>%map(moving_avg,rolling_widow=roll)
  })
  
  
  ## moving average graph
  
  moving_avg_graphs<-reactive({
    moving_avg_graphs<-list()
    moving_avg_data<-moving_avg_data()
    Tops<-as.numeric(input$Tops)
    roll<-input$roll
    for(dat in names(moving_avg_data)){
      data<-moving_avg_data[[dat]]
      
      moving_avg_graphs[[dat]]<-data%>%
        pivot_longer(cols=contains("Top_"),names_to="series",values_to="price")%>%
        mutate(series=fct_relevel(series,paste("Top_",Tops,sep="")),Date=as.Date(Date))%>%
        ggplot(aes(x=Date, y=price,colour=series))+
        geom_line()+
        labs(title =paste(strsplit(dat," ")[[1]][1],"Constituents"), y = paste(roll,'day price rolling average'))+
        theme_bw()+
        scale_x_date(date_breaks = "1 month")+
        theme(axis.text.x = element_text(angle = 90))
    }
    moving_avg_graphs
  })
  
  ## Topbottom quintiles
  
  Topbottomquintiles_data<-reactive({
    Tops<-as.numeric(input$Tops)
    moving_avg_data<-moving_avg_data()
    moving_avg_data%>%map(Topbottom,Tops=Tops)
  })
  
  Topbottomquintiles_graphs<-reactive({
    Topbottomquintiles_data<-Topbottomquintiles_data()
    Topbottomquintiles_graphs<-list()
    roll<-input$roll
    
    for(dat in names(Topbottomquintiles_data)){
      data<-Topbottomquintiles_data[[dat]]
      
      Topbottomquintiles_graphs[[dat]]<-data%>%
        ggplot(aes(x=Date,y=`Average price`,colour=quantile))+
        geom_line()+
        geom_point()+
        theme_bw()+
        labs(title =paste(strsplit(dat," ")[[1]][1],"Constituents"), y = paste(roll,'day price rolling average'))+
        facet_wrap(~series,ncol=2)+
        theme(legend.position = "bottom",axis.text.x = element_text(angle=90))+
        scale_x_date(date_breaks = "1 month")}
    
    Topbottomquintiles_graphs
  })
  
  # render content
  output$contents<-DT::renderDataTable({
    macds<-macds()
    moving_avg_data<-moving_avg_data()
    Tops<-input$Tops
    Topbottomquintiles_data<-Topbottomquintiles_data()
    
    if(input$pergraph=="Percentage point oscillator"){
      outd<-macds[[input$inSelect3]]
    }else if(input$pergraph=="Moving Averages"){
      outd<-moving_avg_data[[input$inSelect3]]
    }else{
      outd<-Topbottomquintiles_data[[input$inSelect3]]
    }
    
    outd
  })
  
  # render performance based graph
  
  output$contentsgraph<-shiny::renderPlot({
    ppgraphss<-ppgraphss()
    moving_avg_graphs<-moving_avg_graphs()
    Topbottomquintiles_graphs<-Topbottomquintiles_graphs()
    
    if(input$pergraph=="Percentage point oscillator"){
      outg<-ppgraphss[[input$inSelect3]]
    }else if(input$pergraph=="Moving Averages"){
      outg<-moving_avg_graphs[[input$inSelect3]]
    }else{
      outg<-Topbottomquintiles_graphs[[input$inSelect3]]
    }
    outg
  })
  
  ##### Correlation based graphs 
  
  
  average_corr<-reactive({
    asset_group_data<-datainput()
    corr_rankings<-correlationbasedrankings()
    roll<-input$roll1
    Tops<-as.numeric(input$Tops1)
    
    average_corr<-list()
    
    withProgress(message = 'Preparing data',
                 detail = 'This may take a while...', value = 0, {
                   for (i in 1:10) {
                     incProgress(1/10)
                     Sys.sleep(0.25)
                   }
                   for(dat in names(asset_group_data)){
                     data<-asset_group_data[[dat]]
                     dates<-data$Date
                     
                     data<-data%>%
                       slider::slide(flat,.before=roll-1,.complete=T)%>%
                       bind_rows()%>%
                       mutate_if(is.numeric,atanh)%>%
                       mutate_all(.funs = na_if,-Inf)%>%
                       mutate_all(.funs = na_if,Inf)
                     
                     rankings<- corr_rankings[[dat]]
                     
                     for(num in Tops){
                       data<-data%>%
                         rowwise()%>%
                         mutate(!!as.name(paste("Top_",num,sep="")):=mean(c_across(rankings$term[1:num]),na.rm=T))
                     }
                     
                     average_corr[[dat]]<-data%>%
                       ungroup()%>%
                       mutate(Date=dates[-(1:(roll-1))])%>%
                       select(Date,contains('Top_'))%>%
                       mutate_if(is.numeric,tanh)
                   }
                 })              
    average_corr
  })
  
  
  average_corr_graphs<-reactive({
    average_corr<-average_corr()
    corr_rankings<-correlationbasedrankings()
    roll<-input$roll1
    Tops<-as.numeric(input$Tops1)
    
    average_corr_graphs<-list()
    for(dat in names(average_corr)){
      data<-average_corr[[dat]]
      
      average_corr_graphs[[dat]]<-data%>%
        pivot_longer(cols=contains("Top_"),names_to="series",values_to="Average correlation")%>%
        mutate(series=fct_relevel(series,paste("Top_",Tops,sep="")))%>%
        mutate(pos=`Average correlation`>=0,Date=as.Date(Date))%>%
        ggplot(aes(x=Date, y=`Average correlation`,fill=pos))+
        geom_bar(stat = 'identity')+
        labs(title=paste(strsplit(dat," ")[[1]][1],"Constituents"), y = 'Average correlation')+
        theme_bw()+
        theme(legend.position = 'None',axis.text.x = element_text(angle=90))+
        facet_wrap(~series,ncol=2)+
        scale_x_date(date_breaks = "1 month")
    }
    average_corr_graphs
  })
  
  average_corr_graphs2<-reactive({
    # Average corr lines
    average_corr<-average_corr()
    corr_rankings<-correlationbasedrankings()
    roll<-input$roll1
    Tops<-as.numeric(input$Tops1)
    
    average_corr_graphs2<-list()
    
    for(dat in names(average_corr)){
      data<-average_corr[[dat]]
      
      average_corr_graphs2[[dat]]<-data%>%
        pivot_longer(cols=contains("Top_"),names_to="series",values_to="Average correlation")%>%
        mutate(series=fct_relevel(series,paste("Top_",Tops,sep="")))%>%
        mutate(pos=`Average correlation`>=0,Date=as.Date(Date))%>%
        ggplot(aes(x=Date, y=`Average correlation`,colour=series))+
        geom_line(stat = 'identity')+
        labs(title=paste(strsplit(dat," ")[[1]][1],"Constituents"), y = 'Average correlation')+
        theme_bw()+
        theme(legend.position = "bottom",axis.text.x = element_text(angle=90))+
        scale_x_date(date_breaks = "1 month")
    }
    average_corr_graphs2
  })
  
  
  # Top quantile correlations 
  average_corr_graphs3<-reactive({
    average_corr<-average_corr()
    corr_rankings<-correlationbasedrankings()
    roll<-input$roll1
    Tops<-as.numeric(input$Tops1)
    
    average_corr_graphs3<-list()
    
    for(dat in names(average_corr)){
      data<-average_corr[[dat]]
      
      average_corr_graphs3[[dat]]<-data%>%
        pivot_longer(cols=contains("Top_"),names_to="series",values_to="Average correlation")%>%
        mutate(series=fct_relevel(series,paste("Top_",Tops,sep="")))%>%
        group_by(series)%>%
        mutate(quantile=ifelse(`Average correlation`>=quantile(`Average correlation`,na.rm=T)[4],"Top 25%",
                               ifelse(`Average correlation`<=quantile(`Average correlation`,na.rm=T)[2],"Bottom 25%","middle50")),
               Date=as.Date(Date))%>%
        filter(quantile!='middle50')%>%
        ggplot(aes(x=Date,y=`Average correlation`,colour=quantile))+
        geom_line()+
        geom_point()+
        theme_bw()+
        labs(title=paste(strsplit(dat," ")[[1]][1],"Constituents"),y=paste(roll,"day rolling average"))+
        theme(legend.position = "bottom",axis.text.x = element_text(angle=90))+
        facet_wrap(~series,ncol=2)+
        scale_x_date(date_breaks = "1 month")
    }
    average_corr_graphs3
  }) 
  
  # render content
  output$contents1<-DT::renderDataTable({
    average_corr<-average_corr()
    outd1<-average_corr[[input$inSelect4]]
    outd1
  })
  
  # render performance based graph
  
  output$contentsgraph1<-shiny::renderPlot({
    average_corr_graphs<-average_corr_graphs()
    average_corr_graphs2<-average_corr_graphs2()
    average_corr_graphs3<-average_corr_graphs3()
    
    if(input$corrgraph=="average correlation oscillator"){
      outg1<-average_corr_graphs[[input$inSelect4]]
    }else if(input$corrgraph=="average correlation linegraph"){
      outg1<-average_corr_graphs2[[input$inSelect4]]
    }else{
      outg1<-average_corr_graphs3[[input$inSelect4]]
    }
    outg1
  })
  
  
  output$downloadData <- downloadHandler(
    filename = function() {
      paste(input$dtype, Sys.Date(), ".xlsx", sep="")
    },
    content = function(filex) {
      
      
      wb<-paste(input$dtype,Sys.time())
      assign(wb,createWorkbook(type="xlsx"))
      
      ################################### formating ##############################################################
      
      # Excel Formatings
      
      CellStyle(get(wb), dataFormat=NULL, alignment=NULL,
                border=NULL, fill=NULL, font=NULL)
      
      # Define some cell styles
      #++++++++++++++++++++
      # Title and sub title styles
      TITLE_STYLE <- CellStyle(get(wb))+ Font(get(wb),  heightInPoints=16, 
                                              color="blue", isBold=TRUE, underline=1)
      SUB_TITLE_STYLE <- CellStyle(get(wb)) + 
        Font(get(wb),  heightInPoints=14, 
             isItalic=TRUE, isBold=TRUE,color="orange")
      
      SUB_TITLE_STYLE1 <- CellStyle(get(wb)) + 
        Font(get(wb),  heightInPoints=14, 
             isItalic=TRUE, isBold=FALSE)
      # Styles for the data table row/column names
      TABLE_ROWNAMES_STYLE <- CellStyle(get(wb)) + Font(get(wb), isBold=TRUE)
      TABLE_COLNAMES_STYLE <- CellStyle(get(wb)) + Font(get(wb), isBold=TRUE) +
        Alignment(wrapText=TRUE, horizontal="ALIGN_CENTER") +
        Border(color="black", position=c("TOP", "BOTTOM"), 
               pen=c("BORDER_THIN", "BORDER_THICK")) 
      
      #++++++++++++++++++++++++
      ########################################################################################
      
      # raw data
      if(input$dtype=="Raw data"){     
        asset_group_data<-datainput()
        
        for(dat in names(asset_group_data)){
          data<-asset_group_data[[dat]]
          sheetname<-paste(strsplit(dat," ")[[1]][1],"Data",sep="_")
          sheetlabel<-paste(strsplit(dat," ")[[1]][1],"Data")
          
          assign(sheetname,createSheet(get(wb), sheetName = paste(sheetlabel)))
          
          # Add title
          xlsx.addTitle(get(sheetname), rowIndex=1, title=paste(sheetlabel),
                        titleStyle = TITLE_STYLE)
          # Add sub title
          xlsx.addTitle(get(sheetname), rowIndex=2, 
                        title=paste("Dataset for",strsplit(dat," ")[[1]][1],"constituents"),
                        titleStyle = SUB_TITLE_STYLE1)
          
          
          # Add a table
          addDataFrame(data,get(sheetname), startRow=3, startColumn=1, 
                       colnamesStyle = TABLE_COLNAMES_STYLE,
                       rownamesStyle = TABLE_ROWNAMES_STYLE)
          
          # Change column width
          setColumnWidth(get(sheetname), colIndex=c(1:ncol(data)), colWidth=11)
          setColumnWidth(get(sheetname), colIndex=2, colWidth=19)
        } 
        
      }else{
        
        all_returns<-returns()
        for(dat in names(all_returns)){
          data<-all_returns[[dat]]
          
          sheetname<-paste(strsplit(dat," ")[[1]][1],"returns",sep="_")
          sheetlabel<-paste(strsplit(dat," ")[[1]][1],"returns")
          assign(sheetname,createSheet(get(wb), sheetName = paste(sheetlabel)))
          
          # Add title
          xlsx.addTitle(get(sheetname), rowIndex=1, title=paste(sheetlabel),
                        titleStyle = TITLE_STYLE)
          # Add sub title
          xlsx.addTitle(get(sheetname), rowIndex=2, 
                        title=paste("Calculated",strsplit(dat," ")[[1]][1], "daily percentage returns"),
                        titleStyle = SUB_TITLE_STYLE1)
          
          # Add a table
          addDataFrame(data, get(sheetname), startRow=3, startColumn=1, 
                       colnamesStyle = TABLE_COLNAMES_STYLE,
                       rownamesStyle = TABLE_ROWNAMES_STYLE)
          # Change column width
          setColumnWidth(get(sheetname), colIndex=c(1:ncol(data)), colWidth=11)
          setColumnWidth(get(sheetname), colIndex=2, colWidth=19)
        }
        
      }
      
      saveWorkbook(get(wb),filex)
    })
  
  output$downloadData1 <- downloadHandler(
    filename = function() {
      paste("performance based rankings", Sys.Date(), ".xlsx", sep="")
    },
    content = function(filex1) {
      
      
      wb<-paste("performance based rankings",Sys.time())
      assign(wb,createWorkbook(type="xlsx"))
      
      ################################### formating ##############################################################
      
      # Excel Formatings
      
      CellStyle(get(wb), dataFormat=NULL, alignment=NULL,
                border=NULL, fill=NULL, font=NULL)
      
      # Define some cell styles
      #++++++++++++++++++++
      # Title and sub title styles
      TITLE_STYLE <- CellStyle(get(wb))+ Font(get(wb),  heightInPoints=16, 
                                              color="blue", isBold=TRUE, underline=1)
      SUB_TITLE_STYLE <- CellStyle(get(wb)) + 
        Font(get(wb),  heightInPoints=14, 
             isItalic=TRUE, isBold=TRUE,color="orange")
      
      SUB_TITLE_STYLE1 <- CellStyle(get(wb)) + 
        Font(get(wb),  heightInPoints=14, 
             isItalic=TRUE, isBold=FALSE)
      # Styles for the data table row/column names
      TABLE_ROWNAMES_STYLE <- CellStyle(get(wb)) + Font(get(wb), isBold=TRUE)
      TABLE_COLNAMES_STYLE <- CellStyle(get(wb)) + Font(get(wb), isBold=TRUE) +
        Alignment(wrapText=TRUE, horizontal="ALIGN_CENTER") +
        Border(color="black", position=c("TOP", "BOTTOM"), 
               pen=c("BORDER_THIN", "BORDER_THICK")) 
      
      #++++++++++++++++++++++++
      ########################################################################################
      
      Perfomance_based_rankings <- createSheet(get(wb), sheetName = "Perfomance based rankings")
      
      
      # Add title
      xlsx.addTitle(Perfomance_based_rankings, rowIndex=1, title="Perfomance based rankings",
                    titleStyle = TITLE_STYLE)
      
      performance_rankings<-perfomancebasedranking()
      
      for(dat in names(performance_rankings)){
        
        data<-performance_rankings[[dat]]
        posit<-which(names(performance_rankings)==dat)
        colu<-c(1,seq(5,(5*length(names(performance_rankings))-1),5))
        
        # Add sub title
        
        xlsx.assert_create_orappend(Perfomance_based_rankings, rowIndex=2, colIndex =  colu[posit],
                                    title=paste(strsplit(dat," ")[[1]][1],"group"),
                                    titleStyle = SUB_TITLE_STYLE1)
        
        # Add a table
        addDataFrame(data, Perfomance_based_rankings, startRow=3, startColumn=colu[posit], 
                     colnamesStyle = TABLE_COLNAMES_STYLE,
                     rownamesStyle = TABLE_ROWNAMES_STYLE)
        
      }
      saveWorkbook(get(wb),filex1)
      
      
      
    })
  
  
  
  output$downloadData2 <- downloadHandler(
    filename = function() {
      paste("Correlation based rankings", Sys.Date(), ".xlsx", sep="")
    },
    content = function(filex2) {
      
      
      wb<-paste("Correlation based rankings",Sys.time())
      assign(wb,createWorkbook(type="xlsx"))
      
      ################################### formating ##############################################################
      
      # Excel Formatings
      
      CellStyle(get(wb), dataFormat=NULL, alignment=NULL,
                border=NULL, fill=NULL, font=NULL)
      
      # Define some cell styles
      #++++++++++++++++++++
      # Title and sub title styles
      TITLE_STYLE <- CellStyle(get(wb))+ Font(get(wb),  heightInPoints=16, 
                                              color="blue", isBold=TRUE, underline=1)
      SUB_TITLE_STYLE <- CellStyle(get(wb)) + 
        Font(get(wb),  heightInPoints=14, 
             isItalic=TRUE, isBold=TRUE,color="orange")
      
      SUB_TITLE_STYLE1 <- CellStyle(get(wb)) + 
        Font(get(wb),  heightInPoints=14, 
             isItalic=TRUE, isBold=FALSE)
      # Styles for the data table row/column names
      TABLE_ROWNAMES_STYLE <- CellStyle(get(wb)) + Font(get(wb), isBold=TRUE)
      TABLE_COLNAMES_STYLE <- CellStyle(get(wb)) + Font(get(wb), isBold=TRUE) +
        Alignment(wrapText=TRUE, horizontal="ALIGN_CENTER") +
        Border(color="black", position=c("TOP", "BOTTOM"), 
               pen=c("BORDER_THIN", "BORDER_THICK")) 
      
      #++++++++++++++++++++++++
      ########################################################################################
      
      
      correlation_based_rankings <- createSheet(get(wb), sheetName = "Correlation based rankings")
      
      # Add title
      xlsx.addTitle(correlation_based_rankings, rowIndex=1, title="Correlation based rankings",
                    titleStyle = TITLE_STYLE)
      
      corr_rankings<-correlationbasedrankings()
      
      for(dat in names(corr_rankings)){
        data<-corr_rankings[[dat]]
        posit<-which(names(corr_rankings)==dat)
        colu<-c(1,seq(5,(5*length(names(corr_rankings))-1),5))
        
        
        # Add sub title
        
        xlsx.assert_create_orappend(correlation_based_rankings, rowIndex=2, colIndex =  colu[posit],
                                    title=paste(strsplit(dat," ")[[1]][1],"group"),
                                    titleStyle = SUB_TITLE_STYLE1)
        
        # Add a table
        addDataFrame(data, correlation_based_rankings, startRow=3, startColumn=colu[posit], 
                     colnamesStyle = TABLE_COLNAMES_STYLE,
                     rownamesStyle = TABLE_ROWNAMES_STYLE)
        
      }
      
      
      saveWorkbook(get(wb),filex2)
      
    })
  
  output$downloadData3 <- downloadHandler(
    filename = function() {
      paste("performance based graphs", Sys.Date(), ".xlsx", sep="")
    },
    content = function(filex3) {
      
      wb<-paste("performance based graphs",Sys.time())
      assign(wb,createWorkbook(type="xlsx"))
      
      ################################### formating ##############################################################
      
      # Excel Formatings
      
      CellStyle(get(wb), dataFormat=NULL, alignment=NULL,
                border=NULL, fill=NULL, font=NULL)
      
      # Define some cell styles
      #++++++++++++++++++++
      # Title and sub title styles
      TITLE_STYLE <- CellStyle(get(wb))+ Font(get(wb),  heightInPoints=16, 
                                              color="blue", isBold=TRUE, underline=1)
      SUB_TITLE_STYLE <- CellStyle(get(wb)) + 
        Font(get(wb),  heightInPoints=14, 
             isItalic=TRUE, isBold=TRUE,color="orange")
      
      SUB_TITLE_STYLE1 <- CellStyle(get(wb)) + 
        Font(get(wb),  heightInPoints=14, 
             isItalic=TRUE, isBold=FALSE)
      
      # Styles for the data table row/column names
      TABLE_ROWNAMES_STYLE <- CellStyle(get(wb)) + Font(get(wb), isBold=TRUE)
      TABLE_COLNAMES_STYLE <- CellStyle(get(wb)) + Font(get(wb), isBold=TRUE) +
        Alignment(wrapText=TRUE, horizontal="ALIGN_CENTER") +
        Border(color="black", position=c("TOP", "BOTTOM"), 
               pen=c("BORDER_THIN", "BORDER_THICK")) 
      
      #++++++++++++++++++++++++
      ########################################################################################
      
      # performance based graphs
      nFast<-input$nFast
      nSlow<-input$nSlow
      roll<-input$roll
      macds<-macds()
      ppgraphs<-ppgraphss()
      moving_avg_data<-moving_avg_data()
      moving_avg_graphs<-moving_avg_graphs()
      Topbottomquintiles_data<-Topbottomquintiles_data()
      Topbottomquintiles_graphs<-Topbottomquintiles_graphs()
      
      
      Perfomance_based_graphs <- createSheet(get(wb), sheetName = "Perfomance based graphs")
      
      # price oscillator
      # Add title
      xlsx.addTitle(Perfomance_based_graphs, rowIndex=1, title="Percentage price oscillators",
                    titleStyle = TITLE_STYLE)
      # Add sub title
      xlsx.addTitle(Perfomance_based_graphs, rowIndex=2, 
                    title="percentage price oscillators calculated as:",
                    titleStyle = SUB_TITLE_STYLE)
      # Add sub title
      xlsx.addTitle(Perfomance_based_graphs, rowIndex=3, 
                    title=paste("100*(EMA",nFast," - EMA",nSlow,")/EMA",nSlow,sep=""),
                    titleStyle = SUB_TITLE_STYLE)
      # Add sub title
      xlsx.addTitle(Perfomance_based_graphs, rowIndex=4, 
                    title= paste("where",paste("EMA",nFast,sep=""),"and",paste("EMA",nSlow,sep=""),"are",nFast,"day",
                                 "and",nSlow,"day","exponential moving averages"),
                    titleStyle = SUB_TITLE_STYLE)
      
      
      startrow<-6
      
      for(dat in names(macds)){
        data<-macds[[dat]]
        graph<-ppgraphs[[dat]]
        posit<-which(names(macds)==dat)
        colu<-c(1,seq((1+ncol(data)+12),((1+ncol(data)+12)*length(names(macds))-1),(1+ncol(data)+12)))
        
        
        xlsx.assert_create_orappend(Perfomance_based_graphs, rowIndex=startrow, colIndex =  colu[posit],
                                    title=paste(strsplit(dat," ")[[1]][1],"group"),
                                    titleStyle = SUB_TITLE_STYLE1)
        
        
        
        addDataFrame(data,Perfomance_based_graphs, startRow=startrow+1, startColumn=colu[posit], 
                     colnamesStyle = TABLE_COLNAMES_STYLE,
                     rownamesStyle = TABLE_ROWNAMES_STYLE)
        
        # add graph
        
        ggsave(paste(dat,".png",sep=""),plot=graph,device = "png",height=6,width=6)
        addPicture(paste(dat,".png",sep=""), Perfomance_based_graphs, scale = 1, startRow = startrow+2,
                   startColumn = (colu[posit])+2+ncol(data))
        unlink(paste(dat,".png",sep=""))
        
        endrow<- startrow+(macds%>%map(nrow)%>%bind_rows()%>%max())
      }
      
      # Price Rolling Averages
      
      # Add title
      
      sectionstart<-endrow+1
      
      xlsx.addTitle(Perfomance_based_graphs, rowIndex=sectionstart+1, title="Price Rolling Averages",
                    titleStyle = TITLE_STYLE)
      # Add sub title
      xlsx.addTitle(Perfomance_based_graphs, rowIndex=sectionstart+2, 
                    title=paste(roll,"day Price rolling averages for each ranking cluster"),
                    titleStyle = SUB_TITLE_STYLE)
      
      startrow<-sectionstart+3
      
      for(dat in names(moving_avg_data)){
        data<-moving_avg_data[[dat]]
        graph<-moving_avg_graphs[[dat]]
        posit<-which(names(moving_avg_data)==dat)
        colu<-c(1,seq((1+ncol(data)+12),((1+ncol(data)+12)*length(names(moving_avg_data))-1),(1+ncol(data)+12)))
        
        
        
        xlsx.assert_create_orappend(Perfomance_based_graphs, rowIndex=startrow, colIndex =  colu[posit],
                                    title=paste(strsplit(dat," ")[[1]][1],"group"),
                                    titleStyle = SUB_TITLE_STYLE1)
        
        addDataFrame(data,Perfomance_based_graphs, startRow=startrow+1, startColumn=colu[posit], 
                     colnamesStyle = TABLE_COLNAMES_STYLE,
                     rownamesStyle = TABLE_ROWNAMES_STYLE)
        
        # add graph
        
        ggsave(paste(dat,".png",sep=""),plot=graph,device = "png",height=6,width=6)
        addPicture(paste(dat,".png",sep=""), Perfomance_based_graphs, scale = 1, startRow = startrow+2,
                   startColumn = (colu[posit])+2+ncol(data))
        unlink(paste(dat,".png",sep=""))
        endrow<- startrow+(moving_avg_data%>%map(nrow)%>%bind_rows()%>%max())
      }
      
      
      # Top and bottom quintile
      
      sectionstart<-endrow+1
      # Add title
      xlsx.addTitle(Perfomance_based_graphs, rowIndex=sectionstart+1, title="Top and Bottom quintile performance days",
                    titleStyle = TITLE_STYLE)
      # Add sub title
      xlsx.addTitle(Perfomance_based_graphs, rowIndex=sectionstart+2, 
                    title=paste("Based on",roll,"price rolling averages"),
                    titleStyle = SUB_TITLE_STYLE)
      
      startrow<-sectionstart+3
      
      for(dat in names(Topbottomquintiles_data)){
        data<-Topbottomquintiles_data[[dat]]
        graph<-Topbottomquintiles_graphs[[dat]]
        posit<-which(names(Topbottomquintiles_data)==dat)
        colu<-c(1,seq((1+ncol(data)+12),((1+ncol(data)+12)*length(names(Topbottomquintiles_data))-1),(1+ncol(data)+12)))
        
        
        
        xlsx.assert_create_orappend(Perfomance_based_graphs, rowIndex=startrow, colIndex =  colu[posit],
                                    title=paste(strsplit(dat," ")[[1]][1],"group"),
                                    titleStyle = SUB_TITLE_STYLE1)
        
        addDataFrame(data,Perfomance_based_graphs, startRow=startrow+1, startColumn=colu[posit], 
                     colnamesStyle = TABLE_COLNAMES_STYLE,
                     rownamesStyle = TABLE_ROWNAMES_STYLE)
        
        # add graph
        
        ggsave(paste(dat,".png",sep=""),plot=graph,device = "png",height=6,width=6)
        addPicture(paste(dat,".png",sep=""), Perfomance_based_graphs, scale = 1, startRow = startrow+2,
                   startColumn = (colu[posit])+2+ncol(data))
        unlink(paste(dat,".png",sep=""))
        endrow<- startrow+(Topbottomquintiles_data%>%map(nrow)%>%bind_rows()%>%max())
      }
      saveWorkbook(get(wb),filex3)
    })
  
  
  output$downloadData4 <- downloadHandler(
    filename = function() {
      paste("correlation based graphs", Sys.Date(), ".xlsx", sep="")
    },
    content = function(filex4) {
      
      
      wb<-paste("correlation based graphs",Sys.time())
      assign(wb,createWorkbook(type="xlsx"))
      
      ################################### formating ##############################################################
      
      # Excel Formatings
      
      CellStyle(get(wb), dataFormat=NULL, alignment=NULL,
                border=NULL, fill=NULL, font=NULL)
      
      # Define some cell styles
      #++++++++++++++++++++
      # Title and sub title styles
      TITLE_STYLE <- CellStyle(get(wb))+ Font(get(wb),  heightInPoints=16, 
                                              color="blue", isBold=TRUE, underline=1)
      SUB_TITLE_STYLE <- CellStyle(get(wb)) + 
        Font(get(wb),  heightInPoints=14, 
             isItalic=TRUE, isBold=TRUE,color="orange")
      
      SUB_TITLE_STYLE1 <- CellStyle(get(wb)) + 
        Font(get(wb),  heightInPoints=14, 
             isItalic=TRUE, isBold=FALSE)
      # Styles for the data table row/column names
      TABLE_ROWNAMES_STYLE <- CellStyle(get(wb)) + Font(get(wb), isBold=TRUE)
      TABLE_COLNAMES_STYLE <- CellStyle(get(wb)) + Font(get(wb), isBold=TRUE) +
        Alignment(wrapText=TRUE, horizontal="ALIGN_CENTER") +
        Border(color="black", position=c("TOP", "BOTTOM"), 
               pen=c("BORDER_THIN", "BORDER_THICK")) 
      
      #++++++++++++++++++++++++
      ########################################################################################
      
      # corr based graphs
      roll<-input$roll1
      corr_rankings<-correlationbasedrankings()
      
      average_corr<-average_corr()
      average_corr_graphs<-average_corr_graphs()
      average_corr_graphs2<-average_corr_graphs2()
      average_corr_graphs3<-average_corr_graphs3()
      
      Correlation_based_graphs <- createSheet(get(wb), sheetName = "Correlation based graphs")
      
      
      # Add title
      xlsx.addTitle(Correlation_based_graphs, rowIndex=1, title="Average correlations for each ranking cluster",
                    titleStyle = TITLE_STYLE)
      # Add sub title
      xlsx.addTitle(Correlation_based_graphs, rowIndex=2, 
                    title=paste("These Average correlations are based on",roll, "day rolling correlations,while considering fisher Z transformation"),
                    titleStyle = SUB_TITLE_STYLE)
      
      
      startrow<-3
      
      for(dat in names(average_corr)){
        data<-average_corr[[dat]]
        graph<-average_corr_graphs[[dat]]
        posit<-which(names(average_corr)==dat)
        colu<-c(1,seq((1+ncol(data)+12),((1+ncol(data)+12)*length(names(corr_rankings))-1),(1+ncol(data)+12)))
        
        
        
        xlsx.assert_create_orappend(Correlation_based_graphs, rowIndex=startrow, colIndex =  colu[posit],
                                    title=paste(strsplit(dat," ")[[1]][1],"group"),
                                    titleStyle = SUB_TITLE_STYLE1)
        
        addDataFrame(data,Correlation_based_graphs, startRow=startrow+1, startColumn=colu[posit], 
                     colnamesStyle = TABLE_COLNAMES_STYLE,
                     rownamesStyle = TABLE_ROWNAMES_STYLE)
        
        # add graph
        
        ggsave(paste(dat,".png",sep=""),plot=graph,device = "png",height=6,width=6)
        addPicture(paste(dat,".png",sep=""), Correlation_based_graphs, scale = 1, startRow = startrow+2,
                   startColumn = (colu[posit])+2+ncol(data))
        unlink(paste(dat,".png",sep=""))
        endrow<- startrow+(average_corr%>%map(nrow)%>%bind_rows()%>%max())
      }
      
      #average based
      
      sectionstart<-endrow+1
      
      # Add title
      xlsx.addTitle(Correlation_based_graphs, rowIndex=sectionstart, title="Average correlations for each ranking cluster",
                    titleStyle = TITLE_STYLE)
      # Add sub title
      xlsx.addTitle(Correlation_based_graphs, rowIndex=sectionstart+1, 
                    title=paste("These Average correlations are based on",roll, "day rolling correlations,while considering fisher Z transformation"),
                    titleStyle = SUB_TITLE_STYLE)
      # Add title
      xlsx.addTitle(Correlation_based_graphs, rowIndex=sectionstart+2, title="Line graph for average correlations for each ranking cluster",
                    titleStyle = TITLE_STYLE)
      
      startrow<-sectionstart+3
      
      for(dat in names(average_corr)){
        data<-average_corr[[dat]]
        graph<-average_corr_graphs2[[dat]]
        posit<-which(names(average_corr)==dat)
        colu<-c(1,seq((1+ncol(data)+12),((1+ncol(data)+12)*length(names(corr_rankings))-1),(1+ncol(data)+12)))
        
        
        
        xlsx.assert_create_orappend(Correlation_based_graphs, rowIndex=startrow, colIndex =  colu[posit],
                                    title=paste(strsplit(dat," ")[[1]][1],"group"),
                                    titleStyle = SUB_TITLE_STYLE1)
        
        addDataFrame(data,Correlation_based_graphs, startRow=startrow+1, startColumn=colu[posit], 
                     colnamesStyle = TABLE_COLNAMES_STYLE,
                     rownamesStyle = TABLE_ROWNAMES_STYLE)
        
        # add graph
        
        ggsave(paste(dat,".png",sep=""),plot=graph,device = "png",height=6,width=6)
        addPicture(paste(dat,".png",sep=""), Correlation_based_graphs, scale = 1, startRow = startrow+2,
                   startColumn = (colu[posit])+2+ncol(data))
        unlink(paste(dat,".png",sep=""))
        endrow<- startrow+(average_corr%>%map(nrow)%>%bind_rows()%>%max())
      }
      
      
      sectionstart<-endrow+1
      
      # Add title
      
      xlsx.addTitle(Correlation_based_graphs, rowIndex=sectionstart, title="Average correlations for each ranking cluster",
                    titleStyle = TITLE_STYLE)
      # Add sub title
      xlsx.addTitle(Correlation_based_graphs, rowIndex=sectionstart+1, 
                    title=paste("These Average correlations are based on",roll, "day rolling correlations,while considering fisher Z transformation"),
                    titleStyle = SUB_TITLE_STYLE)
      # Add title
      xlsx.addTitle(Correlation_based_graphs, rowIndex=sectionstart+2, title="Top and Bottom Quintile correlation clusters",
                    titleStyle = TITLE_STYLE)
      
      
      startrow<-sectionstart+3
      
      for(dat in names(average_corr)){
        data<-average_corr[[dat]]
        graph<-average_corr_graphs3[[dat]]
        posit<-which(names(average_corr)==dat)
        colu<-c(1,seq((1+ncol(data)+12),((1+ncol(data)+12)*length(names(corr_rankings))-1),(1+ncol(data)+12)))
        
        
        
        xlsx.assert_create_orappend(Correlation_based_graphs, rowIndex=startrow, colIndex =  colu[posit],
                                    title=paste(strsplit(dat," ")[[1]][1],"group"),
                                    titleStyle = SUB_TITLE_STYLE1)
        
        addDataFrame(data,Correlation_based_graphs, startRow=startrow+1, startColumn=colu[posit], 
                     colnamesStyle = TABLE_COLNAMES_STYLE,
                     rownamesStyle = TABLE_ROWNAMES_STYLE)
        
        # add graph
        
        ggsave(paste(dat,".png",sep=""),plot=graph,device = "png",height=6,width=6)
        addPicture(paste(dat,".png",sep=""), Correlation_based_graphs, scale = 1, startRow = startrow+2,
                   startColumn = (colu[posit])+2+ncol(data))
        unlink(paste(dat,".png",sep=""))
        endrow<- startrow+(average_corr%>%map(nrow)%>%bind_rows()%>%max())
      }
      
      saveWorkbook(get(wb),filex4)
      
    })
  
  output$downloadData5 <- downloadHandler(
    filename = function() {
      paste("correlation report", Sys.Date(), ".xlsx", sep="")
    },
    content = function(filex5) {
      
      
      wb<-paste("correlation report",Sys.time())
      assign(wb,createWorkbook(type="xlsx"))
      
      ################################### formating ##############################################################
      
      # Excel Formatings
      
      CellStyle(get(wb), dataFormat=NULL, alignment=NULL,
                border=NULL, fill=NULL, font=NULL)
      
      # Define some cell styles
      #++++++++++++++++++++
      # Title and sub title styles
      TITLE_STYLE <- CellStyle(get(wb))+ Font(get(wb),  heightInPoints=16, 
                                              color="blue", isBold=TRUE, underline=1)
      SUB_TITLE_STYLE <- CellStyle(get(wb)) + 
        Font(get(wb),  heightInPoints=14, 
             isItalic=TRUE, isBold=TRUE,color="orange")
      
      SUB_TITLE_STYLE1 <- CellStyle(get(wb)) + 
        Font(get(wb),  heightInPoints=14, 
             isItalic=TRUE, isBold=FALSE)
      # Styles for the data table row/column names
      TABLE_ROWNAMES_STYLE <- CellStyle(get(wb)) + Font(get(wb), isBold=TRUE)
      TABLE_COLNAMES_STYLE <- CellStyle(get(wb)) + Font(get(wb), isBold=TRUE) +
        Alignment(wrapText=TRUE, horizontal="ALIGN_CENTER") +
        Border(color="black", position=c("TOP", "BOTTOM"), 
               pen=c("BORDER_THIN", "BORDER_THICK")) 
      
      #++++++++++++++++++++++++
      ########################################################################################
      asset_group_data<-datainput()
      
      
      for(dat in names(asset_group_data)){
        data<-asset_group_data[[dat]]
        sheetname<-paste(strsplit(dat," ")[[1]][1],"Data",sep="_")
        sheetlabel<-paste(strsplit(dat," ")[[1]][1],"Data")
        
        assign(sheetname,createSheet(get(wb), sheetName = paste(sheetlabel)))
        
        # Add title
        xlsx.addTitle(get(sheetname), rowIndex=1, title=paste(sheetlabel),
                      titleStyle = TITLE_STYLE)
        # Add sub title
        xlsx.addTitle(get(sheetname), rowIndex=2, 
                      title=paste("Dataset for",strsplit(dat," ")[[1]][1],"constituents"),
                      titleStyle = SUB_TITLE_STYLE1)
        
        
        # Add a table
        addDataFrame(data,get(sheetname), startRow=3, startColumn=1, 
                     colnamesStyle = TABLE_COLNAMES_STYLE,
                     rownamesStyle = TABLE_ROWNAMES_STYLE)
        
        # Change column width
        setColumnWidth(get(sheetname), colIndex=c(1:ncol(data)), colWidth=11)
        setColumnWidth(get(sheetname), colIndex=2, colWidth=19)
      }
      
      
      
      all_returns<-returns()
      for(dat in names(all_returns)){
        data<-all_returns[[dat]]
        
        sheetname<-paste(strsplit(dat," ")[[1]][1],"returns",sep="_")
        sheetlabel<-paste(strsplit(dat," ")[[1]][1],"returns")
        
        assign(sheetname,createSheet(get(wb), sheetName = paste(sheetlabel)))
        
        # Add title
        xlsx.addTitle(get(sheetname), rowIndex=1, title=paste(sheetlabel),
                      titleStyle = TITLE_STYLE)
        # Add sub title
        xlsx.addTitle(get(sheetname), rowIndex=2, 
                      title=paste("Calculated",strsplit(dat," ")[[1]][1], "daily percentage returns"),
                      titleStyle = SUB_TITLE_STYLE1)
        
        # Add a table
        addDataFrame(data, get(sheetname), startRow=3, startColumn=1, 
                     colnamesStyle = TABLE_COLNAMES_STYLE,
                     rownamesStyle = TABLE_ROWNAMES_STYLE)
        # Change column width
        setColumnWidth(get(sheetname), colIndex=c(1:ncol(data)), colWidth=11)
        setColumnWidth(get(sheetname), colIndex=2, colWidth=19)
      }
      
      
      Perfomance_based_rankings <- createSheet(get(wb), sheetName = "Perfomance based rankings")
      
      
      # Add title
      xlsx.addTitle(Perfomance_based_rankings, rowIndex=1, title="Perfomance based rankings",
                    titleStyle = TITLE_STYLE)
      
      performance_rankings<-perfomancebasedranking()
      for(dat in names(performance_rankings)){
        data<-performance_rankings[[dat]]
        posit<-which(names(performance_rankings)==dat)
        colu<-c(1,seq(5,(5*length(names(performance_rankings))-1),5))
        
        # Add sub title
        
        xlsx.assert_create_orappend(Perfomance_based_rankings, rowIndex=2, colIndex =  colu[posit],
                                    title=paste(strsplit(dat," ")[[1]][1],"group"),
                                    titleStyle = SUB_TITLE_STYLE1)
        
        # Add a table
        addDataFrame(data, Perfomance_based_rankings, startRow=3, startColumn=colu[posit], 
                     colnamesStyle = TABLE_COLNAMES_STYLE,
                     rownamesStyle = TABLE_ROWNAMES_STYLE)
        
      }
      
      # performance based graphs
      # performance based graphs
      nFast<-input$nFast
      nSlow<-input$nSlow
      roll<-input$roll
      macds<-macds()
      ppgraphs<-ppgraphss()
      moving_avg_data<-moving_avg_data()
      moving_avg_graphs<-moving_avg_graphs()
      Topbottomquintiles_data<-Topbottomquintiles_data()
      Topbottomquintiles_graphs<-Topbottomquintiles_graphs()
      
      Perfomance_based_graphs <- createSheet(get(wb), sheetName = "Perfomance based graphs")
      
      # price oscillator
      # Add title
      xlsx.addTitle(Perfomance_based_graphs, rowIndex=1, title="Percentage price oscillators",
                    titleStyle = TITLE_STYLE)
      # Add sub title
      xlsx.addTitle(Perfomance_based_graphs, rowIndex=2, 
                    title="percentage price oscillators calculated as:",
                    titleStyle = SUB_TITLE_STYLE)
      # Add sub title
      xlsx.addTitle(Perfomance_based_graphs, rowIndex=3, 
                    title=paste("100*(EMA",nFast," - EMA",nSlow,")/EMA",nSlow,sep=""),
                    titleStyle = SUB_TITLE_STYLE)
      # Add sub title
      xlsx.addTitle(Perfomance_based_graphs, rowIndex=4, 
                    title= paste("where",paste("EMA",nFast,sep=""),"and",paste("EMA",nSlow,sep=""),"are",nFast,"day",
                                 "and",nSlow,"day","exponential moving averages"),
                    titleStyle = SUB_TITLE_STYLE)
      
      
      startrow<-6
      for(dat in names(macds)){
        data<-macds[[dat]]
        graph<-ppgraphs[[dat]]
        posit<-which(names(macds)==dat)
        r2to<-1+ncol(data)+12
        colu<-c(1,seq(r2to,(r2to*length(names(macds))-1),r2to))
        
        
        xlsx.assert_create_orappend(Perfomance_based_graphs, rowIndex=6, colIndex =  colu[posit],
                                    title=paste(strsplit(dat," ")[[1]][1],"group"),
                                    titleStyle = SUB_TITLE_STYLE1)
        
        
        
        addDataFrame(data,Perfomance_based_graphs, startRow=7, startColumn=colu[posit], 
                     colnamesStyle = TABLE_COLNAMES_STYLE,
                     rownamesStyle = TABLE_ROWNAMES_STYLE)
        
        # add graph
        
        ggsave(paste(dat,".png",sep=""),plot=graph,device = "png",height=6,width=6)
        addPicture(paste(dat,".png",sep=""), Perfomance_based_graphs, scale = 1, startRow = 8,
                   startColumn = (colu[posit])+2+ncol(data))
        unlink(paste(dat,".png",sep=""))
        
        endrow<- startrow+(macds%>%map(nrow)%>%bind_rows()%>%max())
      }
      
      
      # Price Rolling Averages
      sectionstart<- endrow+1
      # Add title
      xlsx.addTitle(Perfomance_based_graphs, rowIndex=sectionstart+1, title="Price Rolling Averages",
                    titleStyle = TITLE_STYLE)
      # Add sub title
      xlsx.addTitle(Perfomance_based_graphs, rowIndex=sectionstart+3, 
                    title=paste(roll,"day Price rolling averages for each ranking cluster"),
                    titleStyle = SUB_TITLE_STYLE)
      
      startrow<-sectionstart+4
      
      for(dat in names(moving_avg_data)){
        data<-moving_avg_data[[dat]]
        graph<-moving_avg_graphs[[dat]]
        posit<-which(names(moving_avg_data)==dat)
        colu<-c(1,seq((1+ncol(data)+12),((1+ncol(data)+12)*length(names(moving_avg_data))-1),(1+ncol(data)+12)))
        
        
        
        xlsx.assert_create_orappend(Perfomance_based_graphs, rowIndex=startrow, colIndex =  colu[posit],
                                    title=paste(strsplit(dat," ")[[1]][1],"group"),
                                    titleStyle = SUB_TITLE_STYLE1)
        
        addDataFrame(data,Perfomance_based_graphs, startRow=startrow+1, startColumn=colu[posit], 
                     colnamesStyle = TABLE_COLNAMES_STYLE,
                     rownamesStyle = TABLE_ROWNAMES_STYLE)
        
        # add graph
        
        ggsave(paste(dat,".png",sep=""),plot=graph,device = "png",height=6,width=6)
        addPicture(paste(dat,".png",sep=""), Perfomance_based_graphs, scale = 1, startRow = startrow+2,
                   startColumn = (colu[posit])+2+ncol(data))
        unlink(paste(dat,".png",sep=""))
        endrow<- startrow+(moving_avg_data%>%map(nrow)%>%bind_rows()%>%max())
      }
      
      # Top and bottom quintile
      # Add title
      sectionstart<-endrow+1
      xlsx.addTitle(Perfomance_based_graphs, rowIndex=sectionstart+1, title="Top and Bottom quintile performance days",
                    titleStyle = TITLE_STYLE)
      # Add sub title
      xlsx.addTitle(Perfomance_based_graphs, rowIndex=sectionstart+3, 
                    title=paste("Based on",roll,"price rolling averages"),
                    titleStyle = SUB_TITLE_STYLE)
      
      startrow<-sectionstart+4
      
      for(dat in names(Topbottomquintiles_data)){
        data<-Topbottomquintiles_data[[dat]]
        graph<-Topbottomquintiles_graphs[[dat]]
        posit<-which(names(Topbottomquintiles_data)==dat)
        colu<-c(1,seq((1+ncol(data)+12),((1+ncol(data)+12)*length(names(Topbottomquintiles_data))-1),(1+ncol(data)+12)))
        
        
        
        xlsx.assert_create_orappend(Perfomance_based_graphs, rowIndex=startrow, colIndex =  colu[posit],
                                    title=paste(strsplit(dat," ")[[1]][1],"group"),
                                    titleStyle = SUB_TITLE_STYLE1)
        
        addDataFrame(data,Perfomance_based_graphs, startRow=startrow+1, startColumn=colu[posit], 
                     colnamesStyle = TABLE_COLNAMES_STYLE,
                     rownamesStyle = TABLE_ROWNAMES_STYLE)
        
        # add graph
        
        ggsave(paste(dat,".png",sep=""),plot=graph,device = "png",height=6,width=6)
        addPicture(paste(dat,".png",sep=""), Perfomance_based_graphs, scale = 1, startRow = startrow+2,
                   startColumn = (colu[posit])+2+ncol(data))
        unlink(paste(dat,".png",sep=""))
        endrow<- startrow+(Topbottomquintiles_data%>%map(nrow)%>%bind_rows()%>%max())
      }
      
      
      correlation_based_rankings <- createSheet(get(wb), sheetName = "Correlation based rankings")
      
      # Add title
      xlsx.addTitle(correlation_based_rankings, rowIndex=1, title="Correlation based rankings",
                    titleStyle = TITLE_STYLE)
      
      corr_rankings<-correlationbasedrankings()
      
      for(dat in names(corr_rankings)){
        data<-corr_rankings[[dat]]
        posit<-which(names(corr_rankings)==dat)
        colu<-c(1,seq(5,(5*length(names(corr_rankings))-1),5))
        
        
        # Add sub title
        
        xlsx.assert_create_orappend(correlation_based_rankings, rowIndex=2, colIndex =  colu[posit],
                                    title=paste(strsplit(dat," ")[[1]][1],"group"),
                                    titleStyle = SUB_TITLE_STYLE1)
        
        # Add a table
        addDataFrame(data, correlation_based_rankings, startRow=3, startColumn=colu[posit], 
                     colnamesStyle = TABLE_COLNAMES_STYLE,
                     rownamesStyle = TABLE_ROWNAMES_STYLE)
        
      }
      
      
      
      # corr based graphs
      roll<-input$roll1
      corr_rankings<-correlationbasedrankings()
      
      average_corr<-average_corr()
      average_corr_graphs<-average_corr_graphs()
      average_corr_graphs2<-average_corr_graphs2()
      average_corr_graphs3<-average_corr_graphs3()
      
      Correlation_based_graphs <- createSheet(get(wb), sheetName = "Correlation based graphs")
      
      
      # Add title
      xlsx.addTitle(Correlation_based_graphs, rowIndex=1, title="Average correlations for each ranking cluster",
                    titleStyle = TITLE_STYLE)
      # Add sub title
      xlsx.addTitle(Correlation_based_graphs, rowIndex=2, 
                    title=paste("These Average correlations are based on",roll, "day rolling correlations,while considering fisher Z transformation"),
                    titleStyle = SUB_TITLE_STYLE)
      
      
      startrow<-3
      
      for(dat in names(average_corr)){
        data<-average_corr[[dat]]
        graph<-average_corr_graphs[[dat]]
        posit<-which(names(average_corr)==dat)
        colu<-c(1,seq((1+ncol(data)+12),((1+ncol(data)+12)*length(names(corr_rankings))-1),(1+ncol(data)+12)))
        
        
        
        xlsx.assert_create_orappend(Correlation_based_graphs, rowIndex=startrow, colIndex =  colu[posit],
                                    title=paste(strsplit(dat," ")[[1]][1],"group"),
                                    titleStyle = SUB_TITLE_STYLE1)
        
        addDataFrame(data,Correlation_based_graphs, startRow=startrow+1, startColumn=colu[posit], 
                     colnamesStyle = TABLE_COLNAMES_STYLE,
                     rownamesStyle = TABLE_ROWNAMES_STYLE)
        
        # add graph
        
        ggsave(paste(dat,".png",sep=""),plot=graph,device = "png",height=6,width=6)
        addPicture(paste(dat,".png",sep=""), Correlation_based_graphs, scale = 1, startRow = startrow+2,
                   startColumn = (colu[posit])+2+ncol(data))
        unlink(paste(dat,".png",sep=""))
        endrow<- startrow+(average_corr%>%map(nrow)%>%bind_rows()%>%max())
      }
      
      #average based
      
      sectionstart<-endrow+1
      
      # Add title
      xlsx.addTitle(Correlation_based_graphs, rowIndex=sectionstart, title="Average correlations for each ranking cluster",
                    titleStyle = TITLE_STYLE)
      # Add sub title
      xlsx.addTitle(Correlation_based_graphs, rowIndex=sectionstart+1, 
                    title=paste("These Average correlations are based on",roll, "day rolling correlations,while considering fisher Z transformation"),
                    titleStyle = SUB_TITLE_STYLE)
      # Add title
      xlsx.addTitle(Correlation_based_graphs, rowIndex=sectionstart+2, title="Line graph for average correlations for each ranking cluster",
                    titleStyle = TITLE_STYLE)
      
      startrow<-sectionstart+3
      
      for(dat in names(average_corr)){
        data<-average_corr[[dat]]
        graph<-average_corr_graphs2[[dat]]
        posit<-which(names(average_corr)==dat)
        colu<-c(1,seq((1+ncol(data)+12),((1+ncol(data)+12)*length(names(corr_rankings))-1),(1+ncol(data)+12)))
        
        
        
        xlsx.assert_create_orappend(Correlation_based_graphs, rowIndex=startrow, colIndex =  colu[posit],
                                    title=paste(strsplit(dat," ")[[1]][1],"group"),
                                    titleStyle = SUB_TITLE_STYLE1)
        
        addDataFrame(data,Correlation_based_graphs, startRow=startrow+1, startColumn=colu[posit], 
                     colnamesStyle = TABLE_COLNAMES_STYLE,
                     rownamesStyle = TABLE_ROWNAMES_STYLE)
        
        # add graph
        
        ggsave(paste(dat,".png",sep=""),plot=graph,device = "png",height=6,width=6)
        addPicture(paste(dat,".png",sep=""), Correlation_based_graphs, scale = 1, startRow = startrow+2,
                   startColumn = (colu[posit])+2+ncol(data))
        unlink(paste(dat,".png",sep=""))
        endrow<- startrow+(average_corr%>%map(nrow)%>%bind_rows()%>%max())
      }
      
      
      sectionstart<-endrow+1
      
      # Add title
      
      xlsx.addTitle(Correlation_based_graphs, rowIndex=sectionstart, title="Average correlations for each ranking cluster",
                    titleStyle = TITLE_STYLE)
      # Add sub title
      xlsx.addTitle(Correlation_based_graphs, rowIndex=sectionstart+1, 
                    title=paste("These Average correlations are based on",roll, "day rolling correlations,while considering fisher Z transformation"),
                    titleStyle = SUB_TITLE_STYLE)
      # Add title
      xlsx.addTitle(Correlation_based_graphs, rowIndex=sectionstart+2, title="Top and Bottom Quintile correlation clusters",
                    titleStyle = TITLE_STYLE)
      
      
      startrow<-sectionstart+3
      
      for(dat in names(average_corr)){
        data<-average_corr[[dat]]
        graph<-average_corr_graphs3[[dat]]
        posit<-which(names(average_corr)==dat)
        colu<-c(1,seq((1+ncol(data)+12),((1+ncol(data)+12)*length(names(corr_rankings))-1),(1+ncol(data)+12)))
        
        
        
        xlsx.assert_create_orappend(Correlation_based_graphs, rowIndex=startrow, colIndex =  colu[posit],
                                    title=paste(strsplit(dat," ")[[1]][1],"group"),
                                    titleStyle = SUB_TITLE_STYLE1)
        
        addDataFrame(data,Correlation_based_graphs, startRow=startrow+1, startColumn=colu[posit], 
                     colnamesStyle = TABLE_COLNAMES_STYLE,
                     rownamesStyle = TABLE_ROWNAMES_STYLE)
        
        # add graph
        
        ggsave(paste(dat,".png",sep=""),plot=graph,device = "png",height=6,width=6)
        addPicture(paste(dat,".png",sep=""), Correlation_based_graphs, scale = 1, startRow = startrow+2,
                   startColumn = (colu[posit])+2+ncol(data))
        unlink(paste(dat,".png",sep=""))
        endrow<- startrow+(average_corr%>%map(nrow)%>%bind_rows()%>%max())      }
      saveWorkbook(get(wb),filex5)
    })
}

