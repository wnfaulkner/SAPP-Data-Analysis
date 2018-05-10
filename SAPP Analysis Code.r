#########################################################
##### 	SAPP DATA ANALYSIS                        	#####
#########################################################

#install.packages('devtools')
#install.packages("httr")
#install.packages("readr")
#install.packages("data.table")
#install.packages("dplyr")
#install.packages("googlesheets")
#install.packages("stringr")
#install.packages("ReporteRs")
#install.packages("xlsx")
#install.packages("chron")

### INITIAL SETUP ###
options(java.parameters = "- Xmx1024m")

rm(list=ls()) #Remove lists

start_time <- Sys.time()

#library(devtools)
library(magrittr)
library(googlesheets)
library(dplyr)
library(ReporteRs)
library(ggplot2)
library(stringr)
library(reshape2)
library(openxlsx)
library(chron)

### LOAD & CLEAN DATA ###

  # Main data
    #Excel
      #Set working directory
        #M900
          wd <- "C:/Users/WNF/Google Drive/1. FLUX CONTRACTS - CURRENT/2016-09 Missouri Education/3. Missouri Education - GDRIVE/2018-03 SAPP Analysis"
        
        #Thinkpad
          #wd <- "G:/My Drive/1. FLUX CONTRACTS - CURRENT/2016-09 Missouri Education/3. Missouri Education - GDRIVE/2018-03 SAPP Analysis/"
        
        setwd(wd)
      
      #READ MOST RECENTLY MODIFIED EXCEL FILE INTO A WORKBOOK
        survey.data.xlsx <- list.files()[grepl(".xlsx",list.files()) & !grepl("analysis",list.files()) & !grepl(".gsheet",list.files())]
        current.survey.file <- survey.data.xlsx[order(survey.data.xlsx %>% file.mtime, decreasing =  TRUE)][1]
        
        wb <- loadWorkbook(current.survey.file)
        wb.sheetnames <- wb$sheet_names
        sapp.wb.sheetnames <- wb.sheetnames[!grepl("buildings|users",wb.sheetnames)]
        
        
        sapp.ls <- list()
        
      #IMPORT USER TABLE
        #! right now have to add user.is.test variable manually in Excel. Could be made auto later with knowledge of which emails have to be selected for deletion
        users.df <- read.xlsx(xlsxFile = current.survey.file,
                              sheet = "users")
        users.df <- users.df[users.df$user.is.test == 0,!grepl("user.is.test",names(users.df))]
        users.df$tot.num.responses <- ""
        users.df$pct.same <- ""
        
      #IMPORT BUILDING TABLE  
        buildings.df <- read.xlsx(xlsxFile = current.survey.file, sheet = "buildings")
        
      #IMPORT SAPP DATA AND STACK INTO SINGLE DATA FRAME
        #i=1
        for(i in 1:length(sapp.wb.sheetnames)){                              #START LOOP BY SAPP/SHEET
          
          sapp.df.i <- read.xlsx( xlsxFile = current.survey.file,            #Read sheet into data frame
                                sheet = sapp.wb.sheetnames[i]
                                )
          sapp.df.i <- sapp.df.i[!is.na(sapp.df.i$email),!grepl("updated_at", names(sapp.df.i))]   #Remove rows with no user email and the 'updated_at' variable
          sapp.df.i <- sapp.df.i[,c(names(sapp.df.i)[grepl("email|created_at",names(sapp.df.i))],
                                    names(sapp.df.i)[!grepl("email|created_at",names(sapp.df.i))]
                                )]
        
          #Append worksheet name to variable names - Response ID and all sheet questions (excludes 'email' and 'created_at')
            sapp.df.i$id <- paste(sapp.wb.sheetnames[i], sapp.df.i$id, sep = "_")
            names(sapp.df.i) <- c(
                                    names(sapp.df.i)[grepl("email|created_at",names(sapp.df.i))],
                                    paste(sapp.wb.sheetnames[i], 
                                          names(sapp.df.i)[!grepl("email|created_at",names(sapp.df.i))], 
                                          sep = "_")
                                  )
            
          #Add column for SAPP Sheet Name
            sapp.df.i$sapp <- sapp.wb.sheetnames[i]
          
          #Create Date & Time Variables (formed off of 'created_at' variable and dropped 'updated at' variable because was the same as 'created_at in all but 2 cases - only checked 'str' sheet)
            sapp.df.i$date <- substr(sapp.df.i$created_at, 1, 10) %>% as.Date()
            sapp.df.i$year <- substr(sapp.df.i$created_at, 1, 10) %>% as.Date() %>% format(., "%Y")
            sapp.df.i$month <- substr(sapp.df.i$created_at, 1, 10) %>% as.Date() %>% format(., "%m")
            sapp.df.i$day <- substr(sapp.df.i$created_at, 1, 10) %>% as.Date() %>% format(., "%d")
            sapp.df.i$time <- substr(sapp.df.i$created_at, 12, nchar(sapp.df.i$created_at)) %>% chron(times = .)
            
            #! could add time.of.day recoding times as morning, afternoon, night, etc
            #print(c(sapp.wb.sheetnames[i],dim(sapp.df.i)))
          
          #Store results in list
            sapp.ls[[i]] <- sapp.df.i                                     #Store data frame in list (for cleaning)
            #sapp.df.name.i <- paste(sapp.wb.sheetnames[i],".df",sep = "") #Create name for new data frame
            #assign(sapp.df.name.i, sapp.df.i)                             #Assign name to new data frame
        
        } # END LOOP BY SAPP/SHEET
        
        sapp.df <- sapp.ls %>% Reduce(function(x,y) full_join(x,y, all = TRUE),.)
        sapp.df <-  cbind(
                      sapp.df[,!grepl(paste(sapp.wb.sheetnames,collapse = "|"),names(sapp.df))],
                      sapp.df[grepl(paste(sapp.wb.sheetnames,collapse = "|"),names(sapp.df))]
                    )
       
        sapp.ans.colnums.v <-  grep(paste(sapp.wb.sheetnames,collapse = "|"),names(sapp.df)) 
        non.sapp.colnums.v <- c(1:length(names(sapp.df))) %>% setdiff(.,sapp.ans.cols.v)
        non.sapp.colnames.v <- c(1:length(names(sapp.df))) %>% setdiff(.,sapp.ans.cols.v)
        
        #col.alphabetized.order <- sapp.df[grepl(paste(sapp.wb.sheetnames,collapse = "|"),names(sapp.df))] %>% names %>% order
        #sapp.df <- sapp.df[,c(1,3,2,4,5,6,7,col.alphabetized.order + 7)]
        #sapp.df[grepl(paste(sapp.wb.sheetnames,collapse = "|"),names(sapp.df))] <- sapp.df[grepl(paste(sapp.wb.sheetnames,collapse = "|"),names(sapp.df))] %>%
        #                                                                              .[,col.alphabetized.order]                           
                            
        #c(grepl("_",names(sapp.df),order(names(sapp.df))]
  
        
      #CALCULATE PERCENT ANSWERS THE SAME (& OTHER USEFUL USER STATISTICS)
        
        for(j in 1:dim(users.df)[1]){
          user.email.j <- users.df$email[j]
          sapp.responses.df.j <- sapp.df[sapp.df$email == user.email.j,]
          
          users.df[j,]$tot.num.responses <- dim(sapp.responses.df.j)[1]
          
          #print(c(j,user.email.j,dim(user.responses.df.j)[1]))
          if((sapp.responses.df.j %>% dim(.))[1] < 3){next()}else{}
          sapp.responses.df.j <- sapp.responses.df.j[,apply(sapp.responses.df.j,2, function (x){any(!is.na(x))})] #Filter out columns that are all NA (no measured values)
          apply(sapp.responses.df.j,1,function(x){
                                                    mean(x, na.rm = TRUE))
                                                  })
          
          #apply(sapp.responses.df.j[,70:75],2,
          
          #apply(sapp.responses.df.j[,sapp.ans.cols.v],2,function(x) mean(x, na.rm = TRUE))  
          
          
          
        }
        
      #EXPORT FINAL AS EXCEL
        #Create unique folder for output
          output.dir <- paste(wd,"/R script outputs/",
                              "Output_",
                              gsub(":",".",Sys.time()), sep = "")
          dir.create(output.dir, recursive = TRUE)
        
        #Write .xlsx file with three sheets
          setwd(output.dir)
          
          sapp.datasets <- list("users" = users.df, "buildings" = buildings.df, "sapp data" = sapp.df)
          output.file.name <- paste("sapp.data.r.output_",
                                    gsub(":","-",Sys.time()),
                                    ".xlsx",
                                    sep = "")
          write.xlsx(sapp.datasets, file = output.file.name)
          
          setwd(wd)
        
