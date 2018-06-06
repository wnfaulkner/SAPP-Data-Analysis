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
          working.dir <- "C:/Users/WNF/Google Drive/1. FLUX CONTRACTS - CURRENT/2016-09 Missouri Education/3. Missouri Education - GDRIVE/2018-03 SAPP Analysis"
          
        #Thinkpad
          #working.dir <- "G:/My Drive/1. FLUX CONTRACTS - CURRENT/2016-09 Missouri Education/3. Missouri Education - GDRIVE/2018-03 SAPP Analysis/"
        
          
        source.dir <- paste(working.dir,"/Source Data/",sep="")
        
        setwd(source.dir)
      
      #READ MOST RECENTLY MODIFIED EXCEL FILE INTO A WORKBOOK
        
        survey.data.xlsx <- list.files()[grepl(".xlsx",list.files()) & !grepl("analysis",list.files()) & !grepl(".gsheet",list.files())]
        current.survey.file <- survey.data.xlsx[order(survey.data.xlsx %>% file.mtime, decreasing =  TRUE)][1]
        
        wb <- loadWorkbook(current.survey.file)
        wb.sheetnames <- wb$sheet_names
        sapp.wb.sheetnames <- wb.sheetnames[!grepl("buildings|users",wb.sheetnames)]
        
        
        sapp.ls <- list()
        
      #IMPORT USER TABLE
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
        buildings.df$building_code <- paste(buildings.df$cd_code,"_",buildings.df$school_codes,sep="")
        buildings.df <- buildings.df[,!grepl("cd_code|school_codes",names(buildings.df))] %>%
                          .[,c(length(names(.)),1:(length(names(.))-1))]
        
      #IMPORT SAPP DATA AND STACK INTO SINGLE DATA FRAME
        
        sapp.profile.names.v <- c("acl","cdt","cfa","es","feedback","lead","metacog","rt","str")
        
        progress.bar.i <- txtProgressBar(min = 0, max = 100, style = 3)
        maxrow <- length(sapp.wb.sheetnames)
        
        #i=1
        for(i in 1:length(sapp.wb.sheetnames)){                              #START LOOP BY SAPP/SHEET
          
          sapp.name.i <- sapp.profile.names.v[!is.na(pmatch(sapp.profile.names.v, sapp.wb.sheetnames[i]))]
          if(length(sapp.name.i) < 1){next()}else{}
          
          #Load sheet into data frame
            sapp.df.i <- read.xlsx( xlsxFile = current.survey.file,            #Read sheet into data frame
                                sheet = sapp.wb.sheetnames[i]
                                )
         
          #Append worksheet name to variable names - Response ID and all sheet questions (excludes 'email' and 'created_at')
        
            sapp.idvar <- names(sapp.df.i)[grep("id",names(sapp.df.i))] %>% .[!grepl("_",.)] 
            names(sapp.df.i)[names(sapp.df.i) == sapp.idvar] <- paste(sapp.name.i,"_",names(sapp.df.i)[grep("id",names(sapp.df.i))] %>% .[!grepl("_",.)], sep = "") #Append SAPP profile name to 'id' variable
            names(sapp.df.i)[grepl("q",names(sapp.df.i))] <- names(sapp.df.i)[grepl("q",names(sapp.df.i))] %>% paste(sapp.name.i,"_",.,sep="")
          
          #Add column for SAPP Sheet Name
            sapp.df.i$sapp <- sapp.name.i
            
          #Re-order columns
            sapp.df.i <- sapp.df.i[, c(which(!grepl("q",names(sapp.df.i))), which(grepl("q",names(sapp.df.i))))]

          #Store results in list
            sapp.df.i$created_at <- sapp.df.i$created_at %>% as.character #prevents error when joining/stacking after loop
            sapp.df.i$updated_at <- sapp.df.i$updated_at %>% as.character #prevents error when joining/stacking after loop
            sapp.ls[[i]] <- sapp.df.i                                     #Store data frame in list (for cleaning)
        
          setTxtProgressBar(progress.bar.i, 100*i/maxrow)
            
        } # END LOOP BY SAPP/SHEET
        
        sapp.df <- sapp.ls %>% Reduce(function(x,y) full_join(x,y, all = TRUE),.) #Stack list outputs from loop
        
        sapp.df$response_id <- 1:dim(sapp.df)[1] #Create an overall id unique for each response
        
        #Create Date & Time Variables (formed off of 'created_at' variable and dropped 'updated at' variable because was the same as 'created_at in all but 2 cases - only checked 'str' sheet)
        sapp.df$created_datetime <- convertToDateTime(sapp.df$created_at)
        created_datetime.na.v <- sapp.df$created_datetime %>% is.na %>% which   #this and line below necessary because some input created_at are already in correct format
        sapp.df$created_datetime[created_datetime.na.v] <- sapp.df$created_at[created_datetime.na.v]
        
        sapp.df$created_year <- sapp.df$created_datetime %>% format(., "%Y")
        sapp.df$month <-  sapp.df$created_datetime %>% format(., "%m")
        sapp.df$day <-  sapp.df$created_datetime %>% format(., "%d")
        sapp.df$time <- substr(sapp.df$created_datetime, 12, nchar(sapp.df$created_datetime %>% as.character)) %>% chron(times = .)
        
        
        sapp.df <-  cbind(
                      sapp.df[,!grepl(paste(paste(sapp.profile.names.v,"_",sep=""),collapse = "|"),names(sapp.df))],
                      sapp.df[grepl(paste(paste(sapp.profile.names.v,"_",sep=""),collapse = "|"),names(sapp.df))]
                    )
       
        sapp.ans.colnums.v <-  grep(paste(sapp.profile.names.v,collapse = "|"),names(sapp.df)) 
        non.sapp.colnums.v <- c(1:length(names(sapp.df))) %>% setdiff(.,sapp.ans.colnums.v)
        non.sapp.colnames.v <- names(sapp.df)[c(1:length(names(sapp.df))) %>% setdiff(.,sapp.ans.colnums.v)]
        
      
      #CALCULATE PERCENT ANSWERS THE SAME (& OTHER USEFUL USER STATISTICS)
        
        sapp.responses.duplicated.ls <- list()
        
        progress.bar.j <- txtProgressBar(min = 0, max = 100, style = 3)
        maxrow <- dim(users.df)[1]
        
        for(j in 1:dim(users.df)[1]){   # START OF LOOP BY USER
          
          user_id.j <- users.df$user_id[j]
          sapp.responses.df.j <- sapp.df[sapp.df$user_id == user_id.j,]
         
          users.df$tot.num.responses[j] <- dim(sapp.responses.df.j)[1]
          
          sapp.responses.df.j <- sapp.responses.df.j[order(sapp.responses.df.j$created_at),apply(sapp.responses.df.j,2, function (x){any(!is.na(x))})] #Filter out columns that are all NA (no measured values)
          
          if(dim(sapp.responses.df.j)[1] < 3){next()}else{}
          
          response.colnames.v.j <- setdiff(names(sapp.responses.df.j),non.sapp.colnames.v) %>% .[!grepl("_id|_at|Overall.Sum",.)]
          
          # Producing graph of duplicated responses and time between them
            sapp.responses.duplicated.df.j <- sapp.responses.df.j[,names(sapp.responses.df.j) %in% response.colnames.v.j] %>%  # Filter out non-unique responses
                                            duplicated(.) %>% 
                                            sapp.responses.df.j[.,]
            if(dim(sapp.responses.duplicated.df.j)[1] < 1){}else{
              sapp.responses.duplicated.df.j <- sapp.responses.duplicated.df.j[with(sapp.responses.duplicated.df.j, order(sapp, created_datetime)),]
              
              sapp.responses.duplicated.df.j$duplicated_timedif <- c(NA,
                                                                     difftime(sapp.responses.duplicated.df.j$created_datetime[2:length(sapp.responses.duplicated.df.j$created_datetime)],
                                                                              sapp.responses.duplicated.df.j$created_datetime[1:length(sapp.responses.duplicated.df.j$created_datetime)-1],
                                                                              units = "hours")
                                                                    )
              sapp.responses.duplicated.df.j$same_sapp <- c(NA,sapp.responses.duplicated.df.j$sapp[2:length(sapp.responses.duplicated.df.j$sapp)] == sapp.responses.duplicated.df.j$sapp[1:length(sapp.responses.duplicated.df.j$sapp)-1])
              sapp.responses.duplicated.df.j$duplicated_timedif[sapp.responses.duplicated.df.j$same_sapp == FALSE] <- NA
              sapp.responses.duplicated.ls[[j]] <- sapp.responses.duplicated.df.j[,names(sapp.responses.duplicated.df.j) %in% c("response_id","duplicated_timedif")]
              
              #sapp.df <- full_join(sapp.df, sapp.responses.duplicated.df.j, by = "response_id")
            }
            
          # Calculating pct.dif variable to guage if users answering differently each time
            sapp.responses.unique.df.j <- sapp.responses.df.j[,names(sapp.responses.df.j) %in% response.colnames.v.j] %>%  # Filter out non-unique responses
                                            duplicated(.) %>% 
                                            `!` %>% 
                                            sapp.responses.df.j[.,]
            
            sapp.table.df.j <- sapp.responses.unique.df.j$sapp %>% table %>% as.data.frame # Table of number of uniqe responses by SAPP
            sapp.compare.v.j <- sapp.table.df.j[,1] %>% .[sapp.table.df.j[,2] > 2] %>% as.character(.) # SAPP names where user responded more than twice to that profile
            
            if(length(sapp.compare.v.j) < 1){next()}else{} # Skip pct.dif calculation if user responded no SAPPs in non-duplicated way more than twice
            
            sapp.responses.compare.df.j <- sapp.responses.unique.df.j[sapp.responses.unique.df.j$sapp %in% sapp.compare.v.j,] # Final user data frame with only non-unique rows for 
            
            pct.dif.responses.ls <- list()
            
            for(k in 1:length(response.colnames.v.j)){  # START OF LOOP BY RESPONSE VARIABLE
              
              responses.v.k <- sapp.responses.unique.df.j[,names(sapp.responses.unique.df.j) == response.colnames.v.j[k]]
              responses.v.k <- responses.v.k[!is.na(responses.v.k)]
              dif.responses.v.k <- (responses.v.k[2:length(responses.v.k)] - responses.v.k[1:(length(responses.v.k)-1)])
              pct.dif.responses.ls[[k]] <- length(dif.responses.v.k[dif.responses.v.k != 0])/length(dif.responses.v.k) #Calculation of statistic: percentage of answers which were different from answer that came before it in time by that user
            
            } # END OF LOOP BY RESPONSE VARIABLE
          
            users.df$pct.dif[j] <- pct.dif.responses.ls %>% unlist %>% mean(.)*100 #Calculation of final statistic: mean of percentage of answers that were different across all variables that the user responded to
            
          #print(c(j,length(sapp.compare.v.j),users.df$pct.dif[j]))
          
          setTxtProgressBar(progress.bar.j, 100*j/maxrow)
          
        } #END OF LOOP BY USER
        
        users.df$tot.num.responses <- users.df$tot.num.responses %>% as.numeric(.)
        users.df$pct.dif <- users.df$pct.dif %>% as.numeric
        
        users.df$pct.dif %>% summary
        
      #EXPORT FINAL AS EXCEL
        #Create unique folder for output
          output.dir <- paste(working.dir,
                            "/Output_",
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
          
          setwd(working.dir)
          
          

