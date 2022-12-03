install.packages(c("tsoutliers", "seasonal", "parallel", "data.table", "beepr","tidyverse","seasonalview","haven","reshape2","ggfortify","stringr","TSstudio","parallel","tsbox", "rlang","vctrs","shiny","readxl", "writexl","svDialogs","dplyr","magrittr","openxlsx"))) 

library("dplyr")
library("tidyverse")
library("svDialogs")
library("tsoutliers")
library("seasonal")
library("parallel")
library("magrittr")
library("writexl")
library("shiny")
library("TSstudio")
library("tsbox")
library("data.table")
library("beepr")
library("seasonalview")
library("haven")
library("reshape2")
library("ggfortify")
library("stringr")
library("rlang")
library("vctrs")
library("shiny")
library("readxl")
library("openxlsx")

df_directory <- file.path("Files_path") 

files_names <- list.files(path = df_directory, pattern = "*.txt")

setwd(df_directory)

temp = files_names

dataframe1 <-data.frame()

for (i in 1:length(temp)){
  
  data <- read.csv(temp[i], header =FALSE,sep=",")
  
  dataframe1 <- rbind(dataframe1,data)
  
  dataframe2 <- dataframe1[,c(1,2,3)]
  
  colnames(dataframe2)[1]<-"Serie"
  colnames(dataframe2)[2]<-"Date"
  colnames(dataframe2)[3]<-"value"
  
  dataframe3 <- reshape(dataframe2, v.names = "value", timevar = "Date", idvar="Serie", direction = "wide") 
  
  names(dataframe3) <- substring(names(dataframe3), 7)
  
  colnames(dataframe3)[1]<-"Serie"
  
}

detalhes_series <- read.csv ("File_path\\File_name.txt", sep =";")

dataframe3[is.na(dataframe3)] <- 0 

dataframe4 <- dataframe3[rowSums(dataframe3[,-1])>0,]

df <- merge(detail_series,dataframe4,by="Serie")

dir_out <- "Output_directory_path\\"

version <- dlgInput("version (vn) - ex. v1", Sys.info()["user"])$res

date_start <- dlgInput("Date in which range starts (AAAA.t) - ex. 2021.2", Sys.info()["user"])$res

date_end <- dlgInput("Date in which range ends (AAAA.t) - ex. 2021.3", Sys.info()["user"])$res

df <- add_column(df, "2021:2" = '', .after = 30)

df[,31] <- df [,28]

df[,28] <- NULL

colnames(df)[c(30,30)] <- "2021:2"

cols <- sapply(df, is.numeric)

outliers_neg <- df[rowSums(df[cols] < 0) > 0, ]

outliers_neg$DesigSerie <- NULL

colnames(outliers_neg) <- sub(":",".",colnames(outliers_neg))

outliers_neg2 <- melt(outliers_neg, id.vars = "Serie", na.rm=TRUE) 

outliers_neg2 <- outliers_neg2 %>% rename('Period'='variable')

outliers_neg2$Periodo <- sub("\\.","0",outliers_neg2$Periodo)

outliers_neg2 <- outliers_neg2[order(outliers_neg2$Periodo),]

date_start2 <- sub("\\.","0",date_start)

date_end2 <- sub("\\.","0",date_end)

outliers_neg3 <-  outliers_neg2 %>% filter(Period>= date_start2 & Period<= date_end2)

outliers_neg4 <- outliers_neg3[outliers_neg3$value < 0, ]

write.table(outliers_neg4, paste(dir_out, 'Outliers_neg_', date_start, '_', date_end, '_', version, '.csv', sep = ''),append=TRUE,sep=';',row.names=FALSE) 

outliers <- function(y){
  
  y <- df
  
  row.names(y) <- NULL
  
  colnames(y)[1]<-'Series'
  
  dates <- sort(colnames(df[,3:ncol(df)]))
  
  year_start <- as.numeric(substr(dates[1],1,4))
  
  year_end <- as.numeric(substr(tail(dates,n=1),1,4)) 
  
  quarter_start <- as.numeric(substr(dates[1],6,6)) 
  
  quarter_end <- as.numeric(substr(tail(dates,n=1),6,6))
  
  i<-1
  
  a <- 1
  
  b <- 1
  
  for (i in 1:nrow(y)){ 
    
    vectname <- as.numeric(y[i,3:ncol(y)])
    
    vectser <- as.numeric(y[i,1])
    
    vectdesig <- as.character(y[i,2])
    
    vectmetaltern <- "method alternative"
    
    timeseries <- ts(vectname, start=c(year_start,quarter_start), end=c(year_end,quarter_end), frequency=4)
    
    m <- try(seas(x=timeseries, outlier=''), silent=TRUE)
    
    relat <- ts_reshape(ts_pc(timeseries), type='long', frequency = 4) %>% rename(var.relat = value)
    
    outputrel <- try(relat %>% mutate(serie=vectser, period = paste(year, quarter,sep="."), value = timeseries[1:length(timeseries)]) %>% select(serie,period,var.relat, value) %>% replace_na(list(var.relat = 0)))
    
    dif <- diff(outputrel$value)
    
    outputrel ["Dif"] <- c(0,dif)
    
    if(class(m)=='try-error'){
      
      output <- try(outputrel %>% mutate(desig=vectdesig, method=vectmetaltern) %>% select(method,serie,desig,period,var.relat,Dif) %>% replace_na(list(var.relat = 0)))
      
      quantiles <- output %>% summarize(Q1 = quantile(var.relat, probs = c(0.25), na.rm = TRUE), Q2 = quantile(var.relat, probs = c(0.5), na.rm = TRUE), Q3 = quantile(var.relat, probs = c(0.75), na.rm = TRUE), IQR = Q3-Q1)
      
      outliers_alternative <- output %>% filter(var.relat > (quantile(var.relat, 0.75, na.rm = TRUE) + 3*IQR(var.relat, na.rm = TRUE)) | var.relat < (quantile(var.relat, 0.25, na.rm = TRUE) - 3*IQR(var.relat, na.rm = TRUE))) 
      
      outliers_alternative <- outliers_alternative %>% filter(var.relat != '0') %>% filter (period >= date_start & period <= date_end)
      
      colnames(outliers_alternative) <- c("method","Serie","Desig","Period","VarPerc","Dif")
      
      if (a==1){
        
        colnames = TRUE
        
      }else{
        
        colnames=FALSE
        
      }
      
      write_csv2(outliers_alternative, paste(dir_out, 'method_alternative_', date_start, '_', date_end, '_', version, '.csv', sep = ''),append=TRUE,col_names=colnames) 
      
      a <- a+1
      
    }else{ 
      
      m_estim <- summary(m)
      
      m_coef <- m_estim$coefficients[,c(1,3), drop=FALSE] 
      
      m_outl <- as.data.frame(m_coef[which(substr(rownames(m_coef), 1,2) %in% c("AO", "LS")),,drop=FALSE])
      
      m_decomp <- ifelse(m_estim$transform == "log", "mult", "add")
      
      m_outl <- m_outl[which(substr(rownames(m_outl),3,6) %in% year_start:year_end),,drop=FALSE]
      
      outlier_count <- length(m_outl[,1])
      
      if (outlier_count == 0) {
        
        output <- "No outliers detected by TRAMO."
        
      } else { 
        
        period <- substr(rownames(m_outl),3,10) 
        
        type <- paste(substr(rownames(m_outl),1,2), m_decomp, sep = "_")
        
        score <- round(unname(m_outl[,2]), 1)
        
        estimate <- unname(m_outl[,1])
        
        vectmettramo <- "Method TRAMO"   
        
        output <- data.frame(period, type, score)
        
        output <- output %>%
          mutate(serie = vectser, method=vectmettramo, desig=vectdesig, severity = case_when(abs(score) > 7 ~ 'Severe',
                                                                                             abs(score) <=7 & abs(score) >= 5 ~ 'Median',
                                                                                             abs(score) < 5 & abs(score) >= 3.5 ~ 'Low')) %>% select(method, serie,period,type,score,severity,desig) %>% filter (period >= date_start & period <= date_end)
        
        colnames(output) <- c("method", "Serie", "Period", "Tiype", "Score", "Severity","Desig")
        
        colnames(outputrel) <- c("Serie","Period","Var%","Value","Dif")
        
        output2 <- merge(outputrel,output,by=c("Serie","Period"))
        
        output2 <- output2 [,c(6,1,10,2,3,5,7,9)]
        
        colnames(output2) <- c("method","Serie","Desig","Period","VarPerc","Dif","Type","Severity")
        
        if (b==1){
          
          colnames = TRUE
          
        }else{
          
          colnames=FALSE
          
        }
        
        write_csv2(output2, paste(dir_out, 'method_TRAMO_', date_start, '_', date_end, '_', version, '.csv', sep = ''),append=TRUE,col_names=colnames)
        
        b <- b+1
        
      }
      
    }
  }
}

outliers(df)

method_alternative <- read.csv (paste(dir_out,'method_alternative_', date_start, '_', date_end, '_', version, '.csv',sep=""),header =TRUE, sep =";")

method_TRAMO <- read.csv (paste(dir_out,'method_TRAMO_', date_start, '_', date_end, '_', version, '.csv',sep=""),header =TRUE, sep =";")

Negatives <- read.csv (paste(dir_out,'Outliers_neg_', date_start, '_', date_end, '_', version, '.csv',sep=""),header =TRUE, sep =";")

Justifications <- read.xlsx ("File_path\\File_name.xlsx")

Negatives$Period <- sub("^(.{5})0", "\\1.", Negatives$Period)

substr(Nagetives$Period,5,5) <- "."

Nagetives <- add_column(Nagetives, method = 'Values Nagetives', .before = 1)

Nagetives <- merge (Nagetives,df,by="Serie")

Nagetives <- Nagetives %>% rename('Desig'='DesigSerie')

Nagetives$value <- NULL

Nagetives <- add_column(Nagetives, Period = '', .after = 4)

Nagetives$Period.1 <- Nagetives$Period

Nagetives$Period <- NULL

Nagetives <- Nagetives %>% rename('Period'='Period.1')

Nagetives <- add_column(Nagetives, VarPerc = '', .after = 3)

Nagetives <- add_column(Nagetives, Dif = '', .after = 4)

Nagetives <- add_column(Nagetives, Type = '', .after = 5)

Nagetives <- add_column(Nagetives, Severity = '', .after = 6)

method_alternative['Type'] <- ''

method_alternative['Severity'] <- ''

method_TRAMO$Tipo <-  substr(method_TRAMO$Type,1,nchar(method_TRAMO$Tipo)-4)

method_TRAMO$Tipo <- ifelse(method_TRAMO$Type=='AO','Temporary shock','Series brak')

method_TRAMO$VarPerc <- scan(text=method_TRAMO$VarPerc, dec=",", sep=".")

method_TRAMO$Dif <- scan(text=method_TRAMO$Dif, dec=",", sep=".")

method_alternative$VarPerc <- scan(text=method_alternative$VarPerc, dec=",", sep=".")

method_alternative$Dif <- scan(text=method_alternative$Dif, dec=",", sep=".")

method_alternative <- merge (method_alternative,df,by="Serie")

method_TRAMO <- merge (method_TRAMO,df,by="Serie")

method_alternative <- method_alternative[-9]

method_TRAMO <- method_TRAMO[-9]

method_alternative_filt <- method_alternative %>% filter(Dif>10|Dif<(-10))

method_TRAMO_filt <- method_TRAMO %>% filter(Dif>10|Dif<(-10))

Outliers <- rbind(method_alternative_filt,method_TRAMO_filt,Nagetives)

Outliers <- Outliers %>% mutate_if(is.numeric, round, digits=3)

Outliers <- Outliers %>% mutate(Dif = substr(Dif, 1, 6))

Outliers <- Outliers %>% mutate(VarPerc= substr(VarPerc, 1, 6))

Outliers <- merge(Outliers,Justifications,by="Serie",all.x=TRUE)

Outliers <- add_column(Outliers, Justifications = '', .after = 8)

Outliers[,9] <- Outliers$Justifications

Outliers$Justifications <- NULL

Outliers <- Outliers %>% rename('Justifications'='Justifications.1')

wb <- createWorkbook()

addWorksheet(wb, "Outliers_")

sht = "Outliers_"

writeDataTable(wb, 1, Outliers, startRow = 1, startCol = 1, tableStyle = "TableStyleLight1", withFilter = FALSE)

sty1 = createStyle(numFmt="0")

addStyle(wb, sht, sty1, rows=2:(nrow(Outliers)+1), cols=5)
addStyle(wb, sht, sty1, rows=2:(nrow(Outliers)+1), cols=6)
addStyle(wb, sht, sty1, rows=2:(nrow(Outliers)+1), cols=1)

saveWorkbook(wb, paste(dir_out, 'Outliers_', date_start, '_', date_end, '_', version, ".xlsx", sep = ''), overwrite = TRUE)