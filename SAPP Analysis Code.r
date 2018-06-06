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
    
  #FINISHING UP CREATION OF TABLES BEFORE EXPORT
    
    #Joining variable for time between duplicate responses to response table (sapp.df)
      sapp.responses.duplicated.ls <- sapp.responses.duplicated.ls[-which(sapply(sapp.responses.duplicated.ls, is.null))] # remove NULL elements of list
      sapp.responses.duplicated.df <- do.call(rbind, sapp.responses.duplicated.ls)  # convert list to data frame
      sapp.df <- full_join(sapp.df, sapp.responses.duplicated.df, by = "response_id", all = TRUE) # join resulting data frame to sapp.df to give it extra timediff column
      
      #Graph density of time difference between duplicate responses
        density(sapp.df$duplicated_timedif[!is.na(sapp.df$duplicated_timedif)] , kernel = "gaussian") %>% plot(., main = "Distribution of time between duplicate responses (sec)")
      
      #Create frequency to determine shape of distribution
        #timediff.cuts <- seq(min(sapp.df$duplicated_timedif),max(sapp.df$duplicated_timedif), 10)
      
    #Changining variable classes for certain columns
      users.df$tot.num.responses <- users.df$tot.num.responses %>% as.numeric(.)
      users.df$pct.dif <- users.df$pct.dif %>% as.numeric
      #users.df$pct.dif %>% summary

      
      
      
      
      
### REPORTING CALCULATIONS & EXPORT TO POWERPOINT ###
  
  reporting.districts <- setdiff(users.df$district %>% unique, c(NA, "Test District"))
  m <- 1
  #for(m in 1:length(reporting.districts)){ # START OF LOOP BY DISTRICT
    
    district.name.m <- reporting.districts[m]
    users.df.m <- users.df[users.df$district == district.name.m,]
    buildings.df.m <- buildings.df[buildings.df]
    
    #Copy template file into target directory & rename with individual report name 
    if(m == 1){
      template.file <- "C:/Users/WNF/Google Drive/1. FLUX CONTRACTS - CURRENT/2016-09 Missouri Education/3. Missouri Education - GDRIVE/2017-09 CWIS automation for MMD/Report Template/CWIS Template.pptx"
      target.dir <- paste("C:/Users/WNF/Google Drive/1. FLUX CONTRACTS - CURRENT/2016-09 Missouri Education/3. Missouri Education - GDRIVE/2018-03 SAPP Analysis/R script outputs/",
                          "Output_",
                          gsub(":",".",Sys.time()), sep = "")
      dir.create(target.dir)
    }
    
    #file.copy(template.file,target.dir)
    target.file.m <- paste( target.dir,
                            "/",
                            "SAPP Report_",
                            district.name.m,
                            ".pptx", sep="") 
    file.copy(template.file, target.file.m)
        
    #Powerpoint Formatting Setup
      pptx.m <- pptx(template = target.file.m)
      
      options("ReporteRs-fontsize" = 20)
      options("ReporteRs-default-font" = "Calibri")
      
      layouts = slide.layouts(pptx.m)
      layouts
      for(k in layouts ){
        slide.layouts(pptx.m, k )
        title(sub = k )
      }
      
    #Useful colors
      titlegreen <- rgb(118,153,48, maxColorValue=255)
      notesgrey <- rgb(131,130,105, maxColorValue=255)
      graphlabelsgrey <- "#5a6b63"
      graphgridlinesgrey <- "#e6e6e6"
      purpleshade <- "#d0abd6"
      purpleheader <- "#3d2242"
      purplegraphshade <- "#402339"
      backgroundgreen <- "#94c132"
      subtextgreen <- "#929e78"
      #notesgray <- rgb(131,130,105, maxColorValue=255)
      
    #Text formatting
      title.format <- textProperties(color = titlegreen, font.size = 48, font.weight = "bold")
      title.format.small <- textProperties(color = titlegreen, font.size = 40, font.weight = "bold")
      subtitle.format <- textProperties(color = notesgrey, font.size = 28, font.weight = "bold")
      section.title.format <- textProperties(color = "white", font.size = 48, font.weight = "bold")
      notes.format <- textProperties(color = notesgrey, font.size = 14)
    
    ### SLIDE ## Cover Slide
    {
      pptx.m <- addSlide( pptx.m, slide.layout = 'Title Slide', bookmark = 1)
      
      #Text itself - as "piece of text" (pot) objects
      title.m <- pot(
        "Self Assessment Practice Profile",
        title.format
      )
  
      districttitle.m <- pot(
        paste("DISTRICT: ",district.name.m %>% toupper),
        subtitle.format
      )
      
      #Write Text into Slide
      pptx.m <- addParagraph(pptx.m, #Title
                             title.m, 
                             height = 1.92,
                             width = 7.76,
                             offx = 1.27,
                             offy = 2.78,
                             par.properties=parProperties(text.align="left", padding=0)
      )
      
      pptx.m <- addParagraph(pptx.m, #District
                             districttitle.m, 
                             height = 0.67,
                             width = 7.76,
                             offx = 1.27,
                             offy = 4.65,
                             par.properties=parProperties(text.align="left", padding=0)
      )
      writeDoc(pptx.m, file = target.file.m) #test Slide 1 build
    }  

    ## SLIDE ## USERS BY SCHOOL
    {  
      # CALCULATIONS
      
      
      # PPT SLIDE CREATION
        pptx.m <- addSlide( pptx.m, slide.layout = 'S2')
        
        #Title
          s3.title <- pot("Overall Scale Performance",title.format)
          pptx.m <- addParagraph(pptx.m, 
                                 s3.title, 
                                 height = 0.89,
                                 width = 8.47,
                                 offx = 0.83,
                                 offy = 0.85,
                                 par.properties=parProperties(text.align="left", padding=0)
          )
          
        #Viz
          s3.outputs.df$category <- factor(s3.outputs.df$category, levels = s3.outputs.df$category)
          
          s3.graph <- ggplot(data = s3.outputs.df, aes(x = rev(category))) + 
            geom_bar(
              aes(y = score_school_avg), 
              stat="identity", 
              fill = c(rep(purplegraphshade, length(s3.outputs.df$score_school_avg)-1), titlegreen),  
              #      cbind(s3.outputs.df$category %>% as.character(.),.) %>%
              #       as.data.frame(.),
              #s3.outputs.df[!is.na(s3.outputs.df$score_school_avg),],
              #by = c("V1","category")
              #)
              
              width = 0.8) + 
            geom_errorbar(
              mapping = aes(ymin = score_state_avg-.015, ymax = score_state_avg-.015), 
              color = "black", 
              width = 0.9,
              size = 1.2,
              alpha = 0.2) +
            geom_errorbar(
              mapping = aes(ymin = score_state_avg, ymax = score_state_avg), 
              color = "#fae029", 
              width = 0.9,
              size = 1.2) +
            ylim(0,5) +
            labs(x = "", y = "") +
            scale_x_discrete(labels = c("Effective Teaching and Learning",
                                        "Common Formative Assessment",
                                        "Data-based Decision-making",
                                        "Leadership",
                                        "Professional Development") %>% rev
            ) +
            geom_text(aes(                                                          #data labels inside base of columns
              y = 0.4, 
              label = s3.outputs.df$score_school_avg %>% round(.,1) %>% format(., nsmall = 1)
            ), 
            size = 4,
            color = "white") + 
            theme(panel.background = element_blank(),
                  panel.grid.major.y = element_blank(),
                  panel.grid.major.x = element_line(color = graphgridlinesgrey),
                  axis.text.x = element_text(size = 0, color = "white"),
                  axis.text.y = element_text(size = 15, color = graphlabelsgrey),
                  axis.ticks = element_blank()
            ) +     
            coord_flip() 
          
          s3.graph
          
          pptx.m <- addPlot(pptx.m,
                            fun = print,
                            x = s3.graph,
                            height = 2.85,
                            width = 6.39,
                            offx = 1.8,
                            offy = 2.16)
          
        #Notes
          s3.notes <- pot("Notes: (a) A score of 4 represents a response of 'most of the time' for Effective Teaching and Learning Practices, Common Formative Assessment, and Data-based Decision-making scales, and 'agree' for the Leadership, and Professional Development scale. (b) The 2017-18 state average is marked with the yellow bar.",
                          notes.format)
          
          pptx.m <- addParagraph(pptx.m,
                                 s3.notes,
                                 height = 1.39,
                                 width = 8.47,
                                 offx = 0.83,
                                 offy = 5.28,
                                 par.properties = parProperties(text.align ="left", padding = 0)
          )
          
          s3.notes.2 <- pot("The 2017-18 state average is marked with the yellow bar.",
                            textProperties(color = "black", font.size = 14))
          
          pptx.m <- addParagraph(pptx.m,
                                 s3.notes.2,
                                 height = 0.71,
                                 width = 8.47,
                                 offx = 2.74,
                                 offy = 6.85,
                                 par.properties = parProperties(text.align ="left", padding = 0)
          )
          
        #Page number
          pptx.m <- addPageNumber(pptx.m)
        
        #Write slide
          writeDoc(pptx.m, file = target.file.m) #test Slide build up to this point
    }  
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
        
  #}  # END OF LOOP BY DISTRICT
  

      
      
      
      
      
            
### EXPORT FINAL Data AS EXCEL ###
{ 
  #Create unique folder for output
   
  
  #Write .xlsx file with three sheets
    setwd(target.dir)
    
    sapp.datasets <- list("users" = users.df, "buildings" = buildings.df, "sapp data" = sapp.df)
    output.file.name <- paste("sapp.data.r.output_",
                              gsub(":","-",Sys.time()),
                              ".xlsx",
                              sep = "")
    Sys.setenv("R_ZIPCMD" = "C:/Rtools/bin/zip.exe") #Avoids error about zipping command using openxlsx (https://github.com/awalker89/openxlsx/issues/111)
    
    write.xlsx(sapp.datasets, file = output.file.name)
   
    setwd(working.dir)
}          
          

