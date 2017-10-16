#########################################################
##### 	R >< GOOGLE SHEETS INTERACITON SCRATCHPAD	#####
#########################################################

#install.packages('devtools')
#install.packages("httr")
#install.packages("readr")
#install.packages("data.table")
#install.packages("dplyr")
#install.packages("googlesheets")
#install.packages("ReporteRs")

### INITIAL SETUP ###
rm(list=ls()) #Remove lists

#library(devtools)
library(magrittr)
library(googlesheets)
library(dplyr)
library(ReporteRs)

### LOAD DATA ###

#trigger.ss <- 	gs_key("1ONbbzVy5N0yNj-41LrGiWo3Y5na9AmUdaCnfbLORMao",verbose = TRUE) 
#trigger.df <- 	gs_read(trigger.ss, ws = "Form Responses 1", range = NULL, literal = TRUE) %>% as.data.frame()

  # School & district name selections
  school.names <- "North Middle School" #!

  # Main data
  cwis.ss <- 	gs_key("1ONbbzVy5N0yNj-41LrGiWo3Y5na9AmUdaCnfbLORMao",verbose = TRUE)
  cwis.df <- 	gs_read(cwis.ss, ws = 1, range = NULL, literal = TRUE) %>% as.data.frame()

### DATA CLEANING & PREP ###

  # Remove trailing column and rows with extra labels
  dat.startrow <- which(cwis.df[,1] %>% substr(.,1,1) == "{")+1
  dat.remove.colnums <- which(cwis.df %>% names %>% substr(.,1,1) == "X")+1
  
  dat.df.1 <- cwis.df[dat.startrow:length(cwis.df[,1]),                            # all rows after row where first cell begins with "{"
                    setdiff(1:length(names(cwis.df)),dat.remove.colnums)]        # all columns except those whose names begin with "X"
  
  names(dat.df.1) <- names(cwis.df)[1:dim(dat.df.1)[2]]

  # Variable renaming for useful variables
  names(dat.df.1)[names(dat.df.1) == "Q27_2"] <- "school.name"
  names(dat.df.1)[names(dat.df.1) == "Q27_1"] <- "district.name"
  names(dat.df.1)[names(dat.df.1) == "Q6"] <- "role"
  names(dat.df.1) <- names(dat.df.1) %>% gsub("_","\\.",.)
  # Exclude responses with 'NA' for school name
  dat.df <- dat.df.1[!is.na(dat.df.1$school.name),]
   
  # Lookup Tables
  
    #Variable names & questions
    vars.df <- cbind(names(cwis.df),cwis.df[1,] %>% t) %>% as.data.frame
    names(vars.df) <- c("var","question.full")
    
    #Answer Options
    ans.opt.always.df <-  cbind(
                            c(1:5),
                            c("Always","Most of the time","About half the time","Sometimes","Never")
                          ) %>% as.data.frame
    names(ans.opt.always.df) <- c("ans.num","ans.text")
    ans.opt.always.df$ans.full <- paste(ans.opt.always.df$ans.num, ans.opt.always.df$ans.text, sep = ". ")
    
    slide.names.v <- c("Participation Details")
  
### PRODUCING TABLES & CHARTS ###

  i <- 1 # for testing loop
  
  #for(i in l: length(school.names)){   #START OF LOOP BY SCHOOL
   
    # Create data frame for this loop - restrict to responses from school name i
    school.name.i <- school.names[i]
    dat.df.i <- dat.df[dat.df$school.name == school.name.i,] %>% tbl_df
    
    #S2 Table for slide 2 "Participation Details"
      s2.mtx <- table(dat.df.i$role) %>% as.matrix
      s2.df <- cbind(row.names(s2.mtx),s2.mtx[,1]) %>% as.data.frame
      rownames(s2.df) <- c("Role","Num Responses")
    
    #S3 Table for slide 3 "Overall Scale Performance"
    
  
    #S6 Table for slide 6 "ETLP Scale Performance"
      s6.headers.v <- paste(c(1:8),
                          c("learning targets",
                            "students assess",
                            "students identify",
                            "feedback to targets",
                            "student to student feedback",
                            "students state criteria",
                            "instruction state standards",
                            "student reviews cfa"),sep = ". "
                          )
      
      s6.varnames.v <- names(dat.df.i)[substr(names(dat.df.i),1,5)=="2.Q4."][c(1:7,10)]
      
      dat.df.i.s6 <- dat.df.i[,names(dat.df.i) %in% s6.varnames.v]
      
      s6.ls <- list()
      
      for(j in 1:length(s6.varnames.v)){
        s6.varname.j <- s6.varnames.v[j]
        s6.df.j <- table(dat.df.i.s6[,names(dat.df.i.s6)==s6.varname.j]) %>% as.matrix %>% as.data.frame
        s6.df.j$ans.opt <- row.names(s6.df.j)
        names(s6.df.j) <- c(s6.varname.j, "ans.opt")
        s6.ls[[j]] <- s6.df.j
      }
      
      s6.df <- Reduce(function(df1,df2) full_join(df1, df2,by = "ans.opt"), s6.ls)
      s6.df <- s6.df[,c(which(names(s6.df)=="ans.opt"),1,3:length(names(s6.df)))] # re-order columns
      s6.df[is.na(s6.df)] <- 0
      s6.df <- left_join(s6.df, ans.opt.always.df, by = c("ans.opt"="ans.text"))
      s6.df <- s6.df[order(s6.df$ans.num),
                     c(which(grepl("ans.full",names(s6.df))),which(!grepl("ans",names(s6.df))))]

  #} # END OF LOOP I BY SCHOOL 
              
### EXPORTING RESULTS TO POWERPOINT ###
  
  j <- 1
  #for(j in 1:length(school.names)){
  
  if(j == 1){
    template.file <- "C:/Users/WNF/Google Drive/1. FLUX CONTRACTS - CURRENT/2016-09 Missouri Education/3. Missouri Education - GDRIVE/2017-09 CWIS automation/2017-18 Results/CWIS_School Report Template.pptx"
    target.file <- "C:/Users/WNF/Desktop/CWIS Automation Testing/Template.pptx"
    file.copy(template.file,target.file)
  }
  
  #Copy template to new file for editing
  target.name.j <- paste( "C:/Users/WNF/Desktop/CWIS Automation Testing/",
                          "CWIS Report_",school.name.i,
                          "_",
                          gsub(":",".",Sys.time()),".pptx", sep="") 
  
  file.rename(target.file, target.name.j)
  
  
  
  
  
  file.from.dir <- "C:/Users/WNF/Desktop/"
  file.to.dir <- "C:/Users/WNF/Desktop/"
  
  #Edit powerpoint template
  
 





### EXPORTING RESULTS TO GOOGLE SHEETS ###
  
  output.ss <- gs_new(title = "Cleaned Data", ws_title = "Cleaned Data")
  gs_edit_cells(output.ss, ws = 1, input = dat.df, verbose = TRUE)
      
      
      
  cwis.ss %>% gs_new("Cleaned Data", input = dat.df)#, verbose = TRUE)
 
  gs_edit_cells(dat, ws='sheetname', input=colnames(result), byrow=TRUE, anchor="A1")
  gs_edit_cells(dat, ws='sheetname', input = result, anchor="A2", col_names=FALSE, trim=TRUE)
  
  } #END OF LOOP J BY SCHOOL
