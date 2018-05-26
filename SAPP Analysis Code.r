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
                              sheet = wb.sheetnames[grepl("users",wb.sheetnames)]) 
        
        #Restrict Dataset
          #For MMD districts only
            mmd.districts.201718 <- c("Cameron R-I","Mound City R-II","Linn Co. R-I","Sheldon R-VIII","Green Forest R-II","Malden R-I",
                                      "Strain-Japan R-XVI","Hayti R-II","Belton 124","University Academy","Sedalia 200","Meramec Valley R-III","Rolla 31",
                                      "Farmington R-VII","Poplar Bluff R-I","Willow Springs R-IV","McDonald Co. R-I","Raytown C-2","Ft. Zumwalt R-II",
                                      "Pattonville R-III")
            users.df$mmd.2017_18 <- users.df$district %in% mmd.districts.201718
          
          #For test users  
            test.user.emails <- c("test.user.emails <- admin@test.com","alex8022@gmail.com","angela.riner-mooney@dese.mo.gov","arden.day@nau.edu","barb.gilpin@dese.mo.gov",
            "cadre3@missouri.pd.org","carla.williams@raytownschools.org","dana.desmond@dese.mo.gov","dev@test.com","district@test.com","jeff.maui4me@gmail.com",
            "jlfreel@hotmail.com","judy.wartick@moed-sail.org","judy.wartick@moedu-sail.org","kohzadi@ucmo.edu","leader@test.com","pcarte@truman.edu","ronda.jenson@moedu-sail.org",
            "ronda@belton.net","ronda@poplarbluff.net","rondadistrict@dev.text","rondaeugenebleader@dev.test","rondalakeroadbleader@dev.test","rondaleader@test.dev",
            "rondamcdonaldcohighbleader@dev.test","rondaoakgrovebleader@test.dev","rondaonealbleader@test.dev","rondapbbernieeleader@dev.test","rpdc@test.com",
            "rpdc1@missouripd.org","rpdc2@missouripd.org","rpdc3@missouripd.org","rpdc4@missouripd.org","rpdc5@missouripd.org","rpdc6@missouripd.org","rpdc7@missouripd.org",
            "rpdc8@missouripd.org","pdc9@missouripd.org","sapp_admin@test.com","sarahmartentest@dev.test","teacher@test.com","teacher3@test.com","teacher4@test.com",
            "testeral@sedalia200.org","thea.scott@dese.mo.gov")
            
            users.df$test.user <- users.df$email %in% test.user.emails 
        
        #New variable placeholders
        users.df$tot.num.responses <- ""
        users.df$pct.dif <- ""
        
      #IMPORT BUILDING TABLE  
        buildings.df <- read.xlsx(xlsxFile = current.survey.file, sheet = wb.sheetnames[grepl("building",wb.sheetnames)])
        
      #IMPORT SAPP DATA AND STACK INTO SINGLE DATA FRAME
        i=1
        #for(i in 1:length(sapp.wb.sheetnames)){                              #START LOOP BY SAPP/SHEET
          
          sapp.df.i <- read.xlsx( xlsxFile = current.survey.file,            #Read sheet into data frame
                                sheet = sapp.wb.sheetnames[i]
                                )
          #sapp.df.i <- sapp.df.i[!is.na(sapp.df.i$email),!grepl("updated_at", names(sapp.df.i))]   #Remove rows with no user email and the 'updated_at' variable
          #sapp.df.i <- sapp.df.i[,c(names(sapp.df.i)[grepl("email|created_at",names(sapp.df.i))],
          #                         names(sapp.df.i)[!grepl("email|created_at",names(sapp.df.i))]
          #                      )]
        
          #Append worksheet name to variable names - Response ID and all sheet questions (excludes 'email' and 'created_at')
            sapp.profile.names.v <- c("acl","cdt","cfa","es","feedback","lead","metacog","rt","str")
            sapp.df.id <- sapp.profile.names.v[!is.na(pmatch(sapp.profile.names.v, sapp.wb.sheetnames[i]))]
            sapp.idvar <- names(sapp.df.i)[grep("id",names(sapp.df.i))] %>% .[!grepl("_",.)] 
            names(sapp.df.i)[names(sapp.df.i) == sapp.idvar] <- paste(sapp.df.id,"_",names(sapp.df.i)[grep("id",names(sapp.df.i))] %>% .[!grepl("_",.)], sep = "")
            
            sapp.df.i$id <- sapp.df.id
            #sapp.df.i$id <- paste(sapp.wb.sheetnames[i], sapp.df.i$id, sep = "_")
            names(sapp.df.i) <- c(
                                    names(sapp.df.i)[grepl("id|created_at",names(sapp.df.i))],
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
        non.sapp.colnums.v <- c(1:length(names(sapp.df))) %>% setdiff(.,sapp.ans.colnums.v)
        non.sapp.colnames.v <- names(sapp.df)[c(1:length(names(sapp.df))) %>% setdiff(.,sapp.ans.colnums.v)]
        
        #col.alphabetized.order <- sapp.df[grepl(paste(sapp.wb.sheetnames,collapse = "|"),names(sapp.df))] %>% names %>% order
        #sapp.df <- sapp.df[,c(1,3,2,4,5,6,7,col.alphabetized.order + 7)]
        #sapp.df[grepl(paste(sapp.wb.sheetnames,collapse = "|"),names(sapp.df))] <- sapp.df[grepl(paste(sapp.wb.sheetnames,collapse = "|"),names(sapp.df))] %>%
        #                                                                              .[,col.alphabetized.order]                           
                            
        #c(grepl("_",names(sapp.df),order(names(sapp.df))]
  
        
      #CALCULATE PERCENT ANSWERS THE SAME (& OTHER USEFUL USER STATISTICS)
        
        for(j in 1:dim(users.df)[1]){   # START OF LOOP BY USER
          
          user.email.j <- users.df$email[j]
          sapp.responses.df.j <- sapp.df[sapp.df$email == user.email.j,]
          
          users.df[j,]$tot.num.responses <- dim(sapp.responses.df.j)[1]
          
          if((sapp.responses.df.j %>% dim(.))[1] < 3){next()}else{}
          
          sapp.responses.df.j <- sapp.responses.df.j[order(sapp.responses.df.j$created_at),apply(sapp.responses.df.j,2, function (x){any(!is.na(x))})] #Filter out columns that are all NA (no measured values)
          
          response.colnames.v.j <- setdiff(names(sapp.responses.df.j),non.sapp.colnames.v) %>% .[!grepl("_id|_at|Overall.Sum",.)]
          
          pct.dif.responses.ls <- list()
          
          for(k in 1:length(response.colnames.v.j)){  # START OF LOOP BY RESPONSE VARIABLE
            responses.v.k <- sapp.responses.df.j[,names(sapp.responses.df.j) == response.colnames.v.j[k]]
            responses.v.k <- responses.v.k[!is.na(responses.v.k)]
            dif.responses.v.k <- (responses.v.k[2:length(responses.v.k)] - responses.v.k[1:(length(responses.v.k)-1)])
            pct.dif.responses.ls[[k]] <- length(dif.responses.v.k[dif.responses.v.k != 0])/length(dif.responses.v.k) #Calculation of statistic: percentage of answers which were different from answer that came before it in time by that user
          } # END OF LOOP BY RESPONSE VARIABLE
        
          users.df$pct.dif[j] <- pct.dif.responses.ls %>% unlist %>% mean(.)*100 #Calculation of final statistic: mean of percentage of answers that were different across all variables that the user responded to
        
        } #END OF LOOP BY USER
        
        users.df$tot.num.responses <- users.df$tot.num.responses %>% as.numeric(.)
        users.df$pct.dif <- users.df$pct.dif %>% as.numeric
        
        users.df$pct.dif %>% summary
        
      #EXPORT FINAL AS EXCEL
        #Create unique folder for output
          output.dir <- paste(wd,
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
          Sys.setenv("R_ZIPCMD" = "C:/Rtools/bin/zip.exe") #Avoids error about zipping command using openxlsx (https://github.com/awalker89/openxlsx/issues/111)
          
          write.xlsx(sapp.datasets, file = output.file.name)
          
          setwd(wd)
          
          
          
          
          
          
          
          
          
###############################################################################################################################################################
# EXTRA CODE
          #print(c(j,user.email.j,dim(user.responses.df.j)[1]))
          
          #sapp.responses.df.j <- sapp.responses.df.j[,setdiff(names(sapp.responses.df.j),non.sapp.colnames.v) %>% .[!grepl("_id",.)]]
          #apply(sapp.responses.df.j,1,function(x){mean(x, na.rm = TRUE)})
          
          #apply(,2,function(x){mean(x, na.rm = TRUE)})
          
          #apply(sapp.responses.df.j[,70:75],2,
          
          #apply(sa
