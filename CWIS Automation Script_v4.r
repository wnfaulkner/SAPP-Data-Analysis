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
library(ggplot2)
library(stringr)

### LOAD DATA ###

  # School & district name selections
  school.names <- "North Middle School" #!

  # Main data
  cwis.ss <- 	gs_key("1ONbbzVy5N0yNj-41LrGiWo3Y5na9AmUdaCnfbLORMao",verbose = TRUE)
  cwis.df <- 	gs_read(cwis.ss, ws = 1, range = NULL, literal = TRUE) %>% as.data.frame()
  
  # Variable helper table
  embed.names.df <- 	gs_read(cwis.ss, ws = 2, range = NULL, literal = TRUE) %>% as.data.frame()

### DATA CLEANING & PREP ###

  # Remove trailing column and rows with extra labels
    dat.startrow <- which(cwis.df[,1] %>% substr(.,1,1) == "{")+1
    dat.remove.colnums <- which(cwis.df %>% names %>% substr(.,1,1) == "X")+1
    
    dat.df.1 <- cwis.df[dat.startrow:length(cwis.df[,1]),                          # all rows after row where first cell begins with "{"
                      setdiff(1:length(names(cwis.df)),dat.remove.colnums)]        # all columns except those whose names begin with "X"
    
    names(dat.df.1) <- names(cwis.df)[1:dim(dat.df.1)[2]]
  
  # Re-stack & collapse columns that are split up because of survey branching 
    split.names.ls <- strsplit(names(dat.df.1),"_")
    #split.names.elements.v <- split.names.ls %>% .[sapply(., length)==1] %>% which %>% unlist
    
    base.varnames.v <-  split.names.ls %>% .[sapply(., length)==1] %>% unlist
    
    ans.opt.varnames.v <- split.names.ls %>% 
                            .[sapply(., length)==2] %>% 
                              sapply(., function(x){paste(x, collapse = "_")}) 
    
    branch.ans.opt.varnames.v <- split.names.ls %>% 
                                 .[sapply(., length)==3] %>% 
                                    sapply(., function(x){paste(x, collapse = "_")})   
    branch.q.names.v <- strsplit(branch.ans.opt.varnames.v, "_") %>% 
                                  unlist %>% 
                                    .[grep("Q",.)] %>% 
                                      unique 
    
    q.ls <- list()
    varname.match.ls <- list()
    #h <- 4 #for testing loop
    
    for(h in 1:length(branch.q.names.v)){     ### START OF LOOP BY QUESTION; only for questions with branched variables
      
      q.name.h <- branch.q.names.v[h]
      varnames.h <- names(dat.df.1)[grep(q.name.h, names(dat.df.1))]
     
        q.ans.options.h <- paste(varnames.h %>% strsplit(.,"_") %>% lapply(., `[[`, 2) %>% unique %>% unlist,
                                varnames.h %>% strsplit(.,"_") %>% lapply(., `[[`, 3) %>% unique %>% unlist,
                                    sep = "_")
        varname.match.ls[[h]] <- q.ans.options.h
        
        collapsed.h.ls <- list()
        
        for(g in 1:length(q.ans.options.h)){    ### START OF LOOP BY ANSWER OPTION

          check.varnames.h <- str_sub(names(dat.df.1), start = -nchar(q.ans.options.h[g]))
          uncollapsed.df <- grep(q.ans.options.h[g],check.varnames.h) %>% c(8,.) %>% dat.df.1[,.]
          uncollapsed.ls <- apply(uncollapsed.df,1, function(x) x[!is.na(x)]) 
          collapsed.df <- do.call(rbind, lapply(uncollapsed.ls, `[`, 1:max(sapply(uncollapsed.ls, length)))) %>% as.data.frame
          collapsed.df[,1] <- collapsed.df[,1] %>% as.character
          collapsed.df[,2] <- collapsed.df[,2] %>% as.character
          names(collapsed.df) <- c("ResponseId",q.ans.options.h[g])        
          collapsed.h.ls[[g]] <- collapsed.df
        
        } ### END OF LOOP BY ANSWER OPTION
      
        q.dat.df <- collapsed.h.ls %>% Reduce(function(x, y) full_join(x,y, all = TRUE), .)
        q.ls[[h + 2]] <- q.dat.df
   
      } ### END OF LOOP BY QUESTION
      
    #Re-merge with non-branched variables
      q.ls[[1]] <- dat.df.1[names(dat.df.1) %in% base.varnames.v]
      q.ls[[2]] <- dat.df.1[names(dat.df.1) %in% c("ResponseId",ans.opt.varnames.v)]
      dat.df.2 <- q.ls %>% Reduce(function(x, y) full_join(x,y, all = TRUE), .) 

   
  # Lookup Tables
  
    #Variable names & questions, adjusting for collapsed columns
      vars.df <- names(dat.df.2) %>% as.data.frame(., stringsAsFactors = FALSE)
      names(vars.df) <- "question.id"
      vars.df$question.full <- ""
      q.ans.options.v <- varname.match.ls %>% do.call(c, .)
      varmatch.df <- paste("1_", q.ans.options.v, sep = "") %>% cbind(q.ans.options.v, .) %>% as.data.frame
      
      #f=25 #LOOP TESTER
      for(f in 1:dim(vars.df)[1]){
        question.id.f <- vars.df$question.id[f]
        
        if(question.id.f %in% varmatch.df$q.ans.options.v){
          match.id.f <- varmatch.df$.[varmatch.df$q.ans.options.v == question.id.f] %>% as.character
        }else{
          match.id.f <- question.id.f
        }
        
        vars.df$question.full[f] <- cwis.df[1,] %>% .[names(cwis.df)==match.id.f] %>% as.character
      }
      
      #Remove repeated parts of questions
      question.full.remove.strings <- c(
        "Please ",
        "Please use the agreement scale to respond to each prompt representing your",
        "Please use the frequency scale to respond to each prompt representing your",
        " - Classroom Teacher - ",
        "\\[Field-2\\]",
        "\\[Field-3\\]"
      )
      vars.df$question.full <- gsub(paste(question.full.remove.strings, collapse = "|"),
                                    "",
                                    vars.df$question.full)
      
      #Add variable names from embedding tool 
      embed.names.df$dat.var.id <- gsub("q","Q",embed.names.df$dat.var.id)
      vars.df <-  left_join(vars.df, embed.names.df, by = c("question.id" = "dat.var.id"))
      
      
    #Answer Options
      ans.opt.always.df <-  cbind(
                              c(5:1),
                              c("Always","Most of the time","About half the time","Sometimes","Never")
                            ) %>% as.data.frame
      names(ans.opt.always.df) <- c("ans.num","ans.text")
      ans.opt.always.df$ans.full <- paste(ans.opt.always.df$ans.num, ans.opt.always.df$ans.text, sep = ". ")
      ans.opt.always.df[,1] <- ans.opt.always.df[,1] %>% as.character %>% as.numeric
      ans.opt.always.df[,2] <- ans.opt.always.df[,2] %>% as.character
      ans.opt.always.df[,3] <- ans.opt.always.df[,3] %>% as.character
      
      slide.names.v <- c("Participation Details")
      
      
  # Variable renaming for useful variables
    names(dat.df.2)[names(dat.df.2) == "Q27_2"] <- "school.name"
    names(dat.df.2)[names(dat.df.2) == "Q27_1"] <- "district.name"
    names(dat.df.2)[names(dat.df.2) == "Q6"] <- "role"
    #names(dat.df.1) <- names(dat.df.1) %>% gsub("_","\\.",.)
      
  # Add numeric variables for ordinal scales
    #dat.df.3 <- left_join(dat.df.2, ans.opt.always.df, by = c("" = ""))

  # Exclude responses with 'NA' for school name
    #ord.vars <- 
    dat.df <- dat.df.2[!is.na(dat.df.2$school.name),] # removes 481 rows in test data      
  
### PRODUCING DATA, TABLES, & CHARTS ###

  i <- 1 # for testing loop
  
  #for(i in l: length(school.names)){   #START OF LOOP BY SCHOOL
   
    # Create data frame for this loop - restrict to responses from school name i
      school.name.i <- school.names[i]
      dat.df.i <- dat.df[dat.df$school.name == school.name.i,] %>% tbl_df
      district.name.i <- dat.df.i$district.name %>% unique
    
    #S2 Table for slide 2 "Participation Details"
    {
      s2.mtx <- table(dat.df.i$role) %>% as.matrix
      s2.df <- cbind(row.names(s2.mtx),s2.mtx[,1]) %>% as.data.frame
      names(s2.df) <- c("Role","Num Responses")
      rownames(s2.df) <- c()
      s2.df$Role <- s2.df$Role %>% as.character #convert factor to character
      s2.df$`Num Responses` <- s2.df$`Num Responses` %>% as.character %>% as.numeric #convert factor to numeric
      s2.df <- rbind(s2.df, c("Total", sum(s2.df[,2] %>% as.character %>% as.numeric))) #add "Total" row
      
    }      
    #S3 Table for slide 3 "Overall Scale Performance"
    {
      #Define input variables for each column
      
      fig.categories <- c("etl","cfa","dbdm")
      fig.outputs.ls <- list()
                                
      #k <- 1 #for testing loop
      
      for(k in 1:length(fig.categories)){ #START OF BY CATEGORY LOOP, I.E. EACH COLUMN REPRESENTS A CATEGORY IN S3
        
        #Subset data
        fig.cat.k <- fig.categories[k]
        cat.calc.vars.k <- vars.df$question.id[vars.df$s3 %>% grepl(fig.cat.k,.)]
        cat.df.k <- dat.df.i[, names(dat.df.i) %in% cat.calc.vars.k]  #Dataset containing all input values for school i
        cat.df.k.state <- dat.df[, names(dat.df.i) %in% cat.calc.vars.k]    #Dataset containing all input values for state
        
        #Re-coding variables as numeric
        ordinal.to.numeric.f <- function(x){
          recode(x, Never=1, Sometimes = 2, `About half the time` = 3 , `Most of the time` = 4, Always = 5)
        }
        
        #Calculate average of re-coded variables
        avg.cat.k <- apply(cat.df.k, 2, ordinal.to.numeric.f) %>% mean(., na.rm = TRUE) #for school
        avg.cat.k.state <- apply(cat.df.k.state, 2, ordinal.to.numeric.f) %>% mean(., na.rm = TRUE) #for school
        
        #Store result along with category name
        fig.outputs.ls[[k]] <- c(fig.cat.k, avg.cat.k, avg.cat.k.state) %>% as.matrix
        
      
      }  #END OF BY CATEGORY LOOP  
     
      s3.outputs.df <- do.call(cbind, fig.outputs.ls ) %>% t %>% as.data.frame
      names(s3.outputs.df) <- c("category","score_school_avg","score_state_avg")
      s3.outputs.df$category <- s3.outputs.df$category %>% as.character
      s3.outputs.df$score_school_avg <- s3.outputs.df$score_school_avg %>% as.character %>% as.numeric %>% round(., digits = 2)
      s3.outputs.df$score_state_avg <- s3.outputs.df$score_state_avg %>% as.character %>% as.numeric %>% round(., digits = 2)
    }    
    #S6 Table for slide 6 "ETLP Scale Performance"
    {
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
      
      s6.varnames.v <- names(dat.df.i)[substr(names(dat.df.i),1,2)=="Q4"][c(1:7,10)]
      
      dat.df.i.s6 <- dat.df.i[,names(dat.df.i) %in% s6.varnames.v]
      
      s6.ls <- list()
      
      for(j in 1:length(s6.varnames.v)){
        s6.varname.j <- s6.varnames.v[j]
        s6.df.j <- table(dat.df.i.s6[,names(dat.df.i.s6)==s6.varname.j]) %>% as.matrix %>% as.data.frame
        s6.df.j$ans.opt <- row.names(s6.df.j)
        names(s6.df.j) <- c(s6.varname.j, "ans.opt")
        s6.ls[[j]] <- s6.df.j
      }
      
      s6.outputs.df <- Reduce(function(df1,df2) full_join(df1, df2,by = "ans.opt"), s6.ls)
      s6.outputs.df <- s6.outputs.df[,c(which(names(s6.outputs.df)=="ans.opt"),1,3:length(names(s6.outputs.df)))] # re-order columns
      s6.outputs.df <- full_join(s6.outputs.df, ans.opt.always.df, by = c("ans.opt"="ans.text"))
      s6.outputs.df[is.na(s6.outputs.df)] <- 0
      s6.outputs.df <- s6.outputs.df[order(s6.outputs.df$ans.num, decreasing = TRUE),
                     c(which(grepl("ans.full",names(s6.outputs.df))),which(!grepl("ans",names(s6.outputs.df))))]
      names(s6.outputs.df) <- c("Answer option","1. learning targets","2. students assess","3. students identify","4. feedback to targets","5. student to student feedback","6. students state criteria", "7. instruction state standards", "student reviews cfa")
    }
    #S7 Bar Chart "ETLP Scale Rates of Implementation"
    {
      s7.headers.v <- s6.headers.v
      s7.varnames.v <- s6.varnames.v
      s7.ls <- list()
      
      for(m in 1:length(s7.varnames.v)){
        #s7.varname.m <- s7.varnames.v[m]
        s7.m <- (100*(s6.outputs.df[1:2,m+1] %>% sum)/ (s6.outputs.df[1:5,m+1] %>% sum)) %>% round(., digits = 0)
        #s7.df.m$ans.opt <- row.names(s7.df.m)
        #names(s7.df.m) <- c(s7.varname.m, "ans.opt")
        s7.ls[[m]] <- s7.m
      }
      
      s7.outputs.df <- do.call(rbind, s7.ls) %>% cbind(s7.headers.v, .) %>% as.data.frame
      names(s7.outputs.df) <- c("category","score_school_rate")
      s7.outputs.df$category <- s7.outputs.df$category %>% as.character
      s7.outputs.df$score_school_rate <- s7.outputs.df$score_school_rate %>% as.character %>% as.numeric
    }
    #S9 Table "CFA Scale Performance"
    {
        s9.headers.v <- paste(c(1:4),
                              c("use cfa",
                                "all in cfa",
                                "student reviews cfa",
                                "cfa used to plan"),sep = ". "
        )
        
        s9.varnames.v <- vars.df$question.id[vars.df$s3 %>% grepl("cfa",.)]
        
        dat.df.i.s9 <- dat.df.i[,names(dat.df.i) %in% s9.varnames.v]
        
        s9.ls <- list()
        
        for(j in 1:length(s9.varnames.v)){
          s9.varname.j <- s9.varnames.v[j]
          s9.df.j <- table(dat.df.i.s9[,names(dat.df.i.s9)==s9.varname.j]) %>% as.matrix %>% as.data.frame
          s9.df.j$ans.opt <- row.names(s9.df.j)
          names(s9.df.j) <- c(s9.varname.j, "ans.opt")
          s9.ls[[j]] <- s9.df.j
        }
        
        s9.outputs.df <- Reduce(function(df1,df2) full_join(df1, df2,by = "ans.opt"), s9.ls)
        s9.outputs.df <- s9.outputs.df[,c(which(names(s9.outputs.df)=="ans.opt"),1,3:length(names(s9.outputs.df)))] # re-order columns
        s9.outputs.df <- full_join(s9.outputs.df, ans.opt.always.df, by = c("ans.opt"="ans.text"))
        s9.outputs.df[is.na(s9.outputs.df)] <- 0
        s9.outputs.df <- s9.outputs.df[order(s9.outputs.df$ans.num, decreasing = TRUE),
                       c(which(grepl("ans.full",names(s9.outputs.df))),which(!grepl("ans",names(s9.outputs.df))))]
        names(s9.outputs.df) <- c("Answer option",s9.headers.v)
    }
    #s10 Bar Chart "CFA Scale Rates of Implementation"
    {
      s10.headers.v <- s9.headers.v
      s10.varnames.v <- s9.varnames.v
      s10.ls <- list()
      
      for(m in 1:length(s10.varnames.v)){
        s10.m <- (100*(s9.outputs.df[1:2,m+1] %>% sum)/(s9.outputs.df[1:nrow(s9.outputs.df),m+1] %>% sum)) %>% round(., digits = 0)
        s10.ls[[m]] <- s10.m
      }
      
      s10.outputs.df <- do.call(rbind, s10.ls) %>% cbind(s10.headers.v, .) %>% as.data.frame
      names(s10.outputs.df) <- c("category","score_school_rate")
      s10.outputs.df$category <- s10.outputs.df$category %>% as.character
      s10.outputs.df$score_school_rate <- s10.outputs.df$score_school_rate %>% as.character %>% as.numeric
    }
    #S12 Table "DBDM Scale Performance"
    {
      s12.headers.v <- c("team reviews data",
                          "team positive",
                          "effective teaming practices",
                          "data determine practices",
                          "visual representations") %>%
                      paste(c(1:length(.)),.,sep = ". ")
      
      
      s12.varnames.v <- vars.df$question.id[vars.df$s3 %>% grepl("dbdm",.)]
      
      dat.df.i.s12 <- dat.df.i[,names(dat.df.i) %in% s12.varnames.v]
      
      s12.ls <- list()
      
      for(j in 1:length(s12.varnames.v)){
        s12.varname.j <- s12.varnames.v[j]
        s12.df.j <- table(dat.df.i.s12[,names(dat.df.i.s12)==s12.varname.j]) %>% as.matrix %>% as.data.frame
        s12.df.j$ans.opt <- row.names(s12.df.j)
        names(s12.df.j) <- c(s12.varname.j, "ans.opt")
        s12.ls[[j]] <- s12.df.j
      }
      
      s12.outputs.df <- Reduce(function(df1,df2) full_join(df1, df2,by = "ans.opt"), s12.ls)
      s12.outputs.df <- s12.outputs.df[,c(which(names(s12.outputs.df)=="ans.opt"),1,3:length(names(s12.outputs.df)))] # re-order columns
      s12.outputs.df <- full_join(s12.outputs.df, ans.opt.always.df, by = c("ans.opt"="ans.text"))
      s12.outputs.df[is.na(s12.outputs.df)] <- 0
      s12.outputs.df <- s12.outputs.df[order(s12.outputs.df$ans.num, decreasing = TRUE),
                                     c(which(grepl("ans.full",names(s12.outputs.df))),which(!grepl("ans",names(s12.outputs.df))))]
      names(s12.outputs.df) <- c("Answer option",s12.headers.v)
    } 
    #S13 Bar Chart "DBDM Scale Rates of Implementation"
    {
      s13.headers.v <- s12.headers.v
      s13.varnames.v <- s12.varnames.v
      s13.ls <- list()
      
      for(m in 1:length(s13.varnames.v)){
        s13.m <- (100*(s12.outputs.df[1:2,m+1] %>% sum)/(s12.outputs.df[1:nrow(s12.outputs.df),m+1] %>% sum)) %>% round(., digits = 0)
        s13.ls[[m]] <- s13.m
      }
      
      s13.outputs.df <- do.call(rbind, s13.ls) %>% cbind(s13.headers.v, .) %>% as.data.frame
      names(s13.outputs.df) <- c("category","score_school_rate")
      s13.outputs.df$category <- s13.outputs.df$category %>% as.character
      s13.outputs.df$score_school_rate <- s13.outputs.df$score_school_rate %>% as.character %>% as.numeric
    }
  
  #} # END OF LOOP I BY SCHOOL 
  
      
      
      
      
                  
### EXPORTING RESULTS TO POWERPOINT #############################################################
  
  j <- 1
  #for(j in 1:length(school.names)){    #START LOOP J BY SCHOOL
  
  #Copy template file into target directory & rename with individual report name 
    if(j == 1){
      template.file <- "C:/Users/WNF/Google Drive/1. FLUX CONTRACTS - CURRENT/2016-09 Missouri Education/3. Missouri Education - GDRIVE/2017-09 CWIS automation/Report Template/CWIS Template.pptx"
      target.dir <- "C:/Users/WNF/Google Drive/1. FLUX CONTRACTS - CURRENT/2016-09 Missouri Education/3. Missouri Education - GDRIVE/2017-09 CWIS automation/3. R script outputs/"
    }
    
    #file.copy(template.file,target.dir)
    target.file.j <- paste( target.dir,
                            "CWIS Report_",school.name.i,
                            "_",
                            gsub(":",".",Sys.time()),".pptx", sep="") 
    file.copy(template.file, target.file.j)
    
  #Powerpoint Formatting Setup
    pptx.j <- pptx(template = target.file.j)
    
    options("ReporteRs-fontsize" = 20)
    options("ReporteRs-default-font" = "Calibri")
    
    layouts = slide.layouts(pptx.j)
    layouts
    for(k in layouts ){
      slide.layouts(pptx.j, k )
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
      subtitle.format <- textProperties(color = notesgrey, font.size = 22, font.weight = "bold")
      section.title.format <- textProperties(color = "white", font.size = 48, font.weight = "bold")
      notes.format <- textProperties(color = notesgrey, font.size = 14)

  
  #Edit powerpoint template
    
    ## SLIDE 1 ##
    {
      pptx.j <- addSlide( pptx.j, slide.layout = 'Title Slide', bookmark = 1)
      
      #Text itself - as "piece of text" (pot) objects
        title.j <- pot(
                        "Collaborative Work Implementation Survey",
                        title.format
                      )
        subtitle.j <- pot(
                        paste("School Report: ",school.name.i),
                        subtitle.format
                      )
      
        districttitle.j <- pot(
                              paste("District: ",district.name.i),
                              subtitle.format
                            )
      
      #Write Text into Slide
        pptx.j <- addParagraph(pptx.j, 
                               title.j, 
                               height = 1.92,
                               width = 7.76,
                               offx = 1.27,
                               offy = 2.78,
                               par.properties=parProperties(text.align="left", padding=0)
                                )
        pptx.j <- addParagraph(pptx.j, 
                               subtitle.j, 
                               height = 0.67,
                               width = 7.76,
                               offx = 1.27,
                               offy = 4.7,
                               par.properties=parProperties(text.align="left", padding=0)
                               )
      
        pptx.j <- addParagraph(pptx.j, 
                               districttitle.j, 
                               height = 0.67,
                               width = 7.76,
                               offx = 1.27,
                               offy = 5.35,
                               par.properties=parProperties(text.align="left", padding=0)
                               )
      #writeDoc(pptx.j, file = target.file.j) #test Slide 1 build
    }  
    ## SLIDE 2 ##
    {
      pptx.j <- addSlide( pptx.j, slide.layout = 'S2')
  
     #Title
        s2.title <- pot("Participation Details",title.format)
        pptx.j <- addParagraph(pptx.j, 
                               s2.title, 
                               height = 0.89,
                               width = 8.47,
                               offx = 0.83,
                               offy = 0.85,
                               par.properties=parProperties(text.align="left", padding=0)
                              )
      
      #Viz
        s2.ft <- FlexTable(
            data = s2.df,
            header.columns = TRUE,
            add.rownames = FALSE,
            
            header.cell.props = cellProperties(background.color = purpleheader),
            header.text.props = textProperties(color = "white", font.size = 22, font.weight = "bold"),
            
            body.cell.props = cellProperties(background.color = "white")
          )
      
        s2.ft <- setFlexTableWidths(s2.ft, widths = c(4, 2))      
        s2.ft <- setZebraStyle(s2.ft, odd = purpleshade, even = "white" ) 
        
        pptx.j <- addFlexTable(pptx.j, 
                               s2.ft, 
                               height = 4.01,
                               width = 6.71,
                               offx = 1.65,
                               offy = 2.75,
                               par.properties=parProperties(text.align="left", padding=0)
                              )
      
      writeDoc(pptx.j, file = target.file.j) #test Slide build up to this point
    
    }    
    ## SLIDE 3 ##
    {  
      pptx.j <- addSlide( pptx.j, slide.layout = 'S2')
      
      #Title
        s3.title <- pot("Overall Scale Performance",title.format)
        pptx.j <- addParagraph(pptx.j, 
                               s3.title, 
                               height = 0.89,
                               width = 8.47,
                               offx = 0.83,
                               offy = 0.85,
                               par.properties=parProperties(text.align="left", padding=0)
        )
      
      #Viz
        s3.graph <- ggplot(data = s3.outputs.df, aes(x = category)) + 
                    geom_bar(
                        aes(y = score_school_avg), 
                        stat="identity", 
                        fill = purplegraphshade, 
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
                    scale_x_discrete(labels = c("Effective Teaching and Learning","Common Formative Assessment","Data-based Decision-making")) +
                    geom_text(aes(
                                y = 0.4, 
                                label = s3.outputs.df$score_school_avg,
                                ), 
                              size = 4,
                              color = "white") + 
                    theme(panel.background = element_blank(),
                          panel.grid.major.y = element_blank(),
                          panel.grid.major.x = element_line(color = graphgridlinesgrey),
                          axis.text.x = element_text(size = 15),
                          axis.text.y = element_text(size = 15, color = graphlabelsgrey)
                          ) +     
                    coord_flip() 
        
        
        
        pptx.j <- addPlot(pptx.j,
                          fun = print,
                          x = s3.graph,
                          height = 2.85,
                          width = 6.39,
                          offx = 1.8,
                          offy = 2.16)
        
      #Notes
        s3.notes <- pot("Notes: (a) A score of 4 represents a response of 'most of the time' for Effective Teaching and Learning Practices, Common Formative Assessment, and Data-based Decision-making scales, and 'agree' for the Leadership, and Professional Development scale. (b) The 2016-17 state average is marked with the yellow bar.",
                        notes.format)
        
        pptx.j <- addParagraph(pptx.j,
                               s3.notes,
                               height = 1.39,
                               width = 8.47,
                               offx = 0.83,
                               offy = 5.28,
                               par.properties = parProperties(text.align ="left", padding = 0)
        )
        
      writeDoc(pptx.j, file = target.file.j) #test Slide build up to this point
    }  
    ## SLIDE 4 ##
    {  
      pptx.j <- addSlide( pptx.j, slide.layout = 'S2')
      
      #Title
        s4.title <- pot("Individual Response Plot",title.format)
        pptx.j <- addParagraph(pptx.j, 
                               s4.title, 
                               height = 0.89,
                               width = 8.47,
                               offx = 0.83,
                               offy = 0.85,
                               par.properties=parProperties(text.align="left", padding=0)
        )
      
      #Viz
      
      #Notes
        s4.notes <- pot("A score of 3 represents a response of 'about half the time' for Effective Teaching and Learning, Common Formative Assessment, and Data-based Decision-making scales, while a 2 represents 'sometimes'. 
                        A 3 corresponds to 'neither agree or disagree' for the Leadership, and Professional Development scale while a 2 represents 'disagree.'",
                        notes.format)
        pptx.j <- addParagraph(pptx.j,
                               s4.notes,
                               height = 1.39,
                               width = 8.47,
                               offx = 0.83,
                               offy = 5.28,
                               par.properties = parProperties(text.align ="left", padding = 0)
        )
        
      writeDoc(pptx.j, file = target.file.j) #test Slide build up to this point
    }  
    ## SLIDE 5 ##
    {  
      pptx.j <- addSlide( pptx.j, slide.layout = 'Section Header')
      
      #S5 Title
        s5.title <- pot("Diving Deeper Into Scales: EFFECTIVE TEACHING AND LEARNING PRACTICES",
                        section.title.format)
        pptx.j <- addParagraph(pptx.j, 
                               s5.title, 
                               height = 2.63,
                               width = 7.76,
                               offx = 1.25,
                               offy = 2.48,
                               par.properties=parProperties(text.align="left", padding=0)
                               )
      writeDoc(pptx.j, file = target.file.j) #test Slide build up to this point
    }
    ## Slide 6 ##
    {
      pptx.j <- addSlide( pptx.j, slide.layout = 'S2')
      
      #Title
        s6.title <- pot("ETLP Scale Performance",title.format)
        pptx.j <- addParagraph(pptx.j, 
                               s6.title, 
                               height = 0.89,
                               width = 8.47,
                               offx = 0.83,
                               offy = 0.18,
                               par.properties=parProperties(text.align="left", padding=0)
        )
      
      #Viz
        s6.ft <- FlexTable(
          data = s6.outputs.df,
          header.columns = TRUE,
          add.rownames = FALSE,
          
          header.cell.props = cellProperties(background.color = purpleheader),
          header.text.props = textProperties(color = "white", font.size = 14, font.weight = "bold"),
          
          body.cell.props = cellProperties(background.color = "white"), 
          body.text.props = textProperties(font.size = 12, font.weight = "bold", color = "#333333")
        )
        
        s6.ft <- setFlexTableWidths(s6.ft, widths = c(1, rep(0.8,ncol(s6.outputs.df)-1)))      
        s6.ft <- setZebraStyle(s6.ft, odd = purpleshade, even = "white" ) 
        s6.ft[,2:(ncol(s6.outputs.df))] <- parProperties(text.align="center")
        
        pptx.j <- addFlexTable(pptx.j, 
                               s6.ft, 
                               height = 1.72,
                               width = 7.77,
                               offx = 0.83,
                               offy = 1.2,
                               par.properties=parProperties(text.align="left", padding=0.2)
        )
      
      #Notes
        s6.notes <- pot("Prompt Text:
                        1. (learning targets) The students in my classroom, including students with disabilities, write/state learning targets using 'I can' or 'I know' statements.
                        2. (students assess) The students in my classroom, including students with disabilities, assess their progress by using evidence of student work (rubrics or portfolios).
                        3. (students identify) The students in my classroom, including students with disabilities, identify what they should do next in their learning based on self-assessment of their progress.
                        4. (feedback to targets) Students in my classroom, including students with disabilities, receive feedback on their progress toward their learning targets.
                        5. (student to student feedback) Student-to-student feedback, focused on improving learning, occurs during instruction.
                        6. (students state criteria) Students in my classroom state the success criteria for achieving their learning target.
                        7. (instruction state standards) The instruction of teachers in my building intentionally addresses the state standards for my grade/subject.
                        8. (student reviews cfa) Each student reviews his/her results of common formative assessments with a teacher.",
                        textProperties(color = notesgrey, font.size = 10))
        
        pptx.j <- addParagraph(pptx.j,
                               s6.notes,
                               height = 3.3,
                               width = 9.57,
                               offx = 0.22,
                               offy = 4.25,
                               par.properties = parProperties(text.align ="left", padding = 0)
        )
        
      writeDoc(pptx.j, file = target.file.j) #test Slide build up to this point
    }
    ## SLIDE 7 ##
    {  
      pptx.j <- addSlide( pptx.j, slide.layout = 'S2')
      
      #Title
        s7.title <- pot("ETLP Scale Rates of Implementation",title.format.small)
        pptx.j <- addParagraph(pptx.j, 
                               s7.title, 
                               height = 0.89,
                               width = 8.47,
                               offx = 0.83,
                               offy = 0.85,
                               par.properties=parProperties(text.align="left", padding=0)
        )
        
      #Viz
        s7.graph <- ggplot(data = s7.outputs.df, aes(x = s7.outputs.df$category[order(s7.outputs.df$category, decreasing = TRUE)])) + 
          geom_bar(
            aes(y = score_school_rate), 
            stat="identity", 
            fill = backgroundgreen, 
            width = 0.8) + 
          ylim(0,100) +
          labs(x = "", y = "% Implementing 'Always' or 'Most of the time'") +
          scale_x_discrete(labels = s7.headers.v[order(s7.headers.v, decreasing = TRUE)]) +
          geom_text(
            aes(y = score_school_rate + 8, label = paste(s7.outputs.df$score_school_rate, "%", sep = "")), 
            size = 4,
            color = graphlabelsgrey) + 
          theme(panel.background = element_blank(),
                panel.grid.major.y = element_blank(),
                panel.grid.major.x = element_line(color = graphgridlinesgrey),
                axis.text.x = element_text(size = 15, color = graphlabelsgrey),
                axis.text.y = element_text(size = 15, color = graphlabelsgrey)
          ) +     
          coord_flip() 

        pptx.j <- addPlot(pptx.j,
                          fun = print,
                          x = s7.graph,
                          height = 3.25,
                          width = 8,
                          offx = 1,
                          offy = 2)
      
      #Notes
        s7.notes <- pot("Documented above are the percent of school respondents who answered that they use the following practices either “most of the time” or “always.” Please see page 6 for a list of complete prompts.",
                        notes.format)
        pptx.j <- addParagraph(pptx.j,
                               s7.notes,
                               height = 1.39,
                               width = 8.47,
                               offx = 0.83,
                               offy = 5.28,
                               par.properties = parProperties(text.align ="left", padding = 0)
        )
        
        writeDoc(pptx.j, file = target.file.j) #test Slide build up to this point
    }
    ## SLIDE 8 ##
    {  
      pptx.j <- addSlide( pptx.j, slide.layout = 'Section Header')
      
      #s8 Title
        s8.title <- pot("Diving Deeper Into Scales: COMMON FORMATIVE ASSESSMENT",
                        section.title.format)
        pptx.j <- addParagraph(pptx.j, 
                               s8.title, 
                               height = 2.63,
                               width = 7.76,
                               offx = 1.25,
                               offy = 2.48,
                               par.properties=parProperties(text.align="left", padding=0)
        )
        writeDoc(pptx.j, file = target.file.j) #test Slide build up to this point
    }
    ## Slide 9 ##
    {
      pptx.j <- addSlide( pptx.j, slide.layout = 'S2')
      
      #Title
      s9.title <- pot("CFA Scale Performance",title.format)
      pptx.j <- addParagraph(pptx.j, 
                             s9.title, 
                             height = 0.89,
                             width = 8.47,
                             offx = 0.83,
                             offy = 0.18,
                             par.properties=parProperties(text.align="left", padding=0)
      )
      
      #Viz
      s9.ft <- FlexTable(
        data = s9.outputs.df,
        header.columns = TRUE,
        add.rownames = FALSE,
        
        header.cell.props = cellProperties(background.color = purpleheader),
        header.text.props = textProperties(color = "white", font.size = 16, font.weight = "bold"),
        
        body.cell.props = cellProperties(background.color = "white"), 
        body.text.props = textProperties(font.size = 15, font.weight = "bold", color = "#333333")
      )
      
      s9.ft[,] <- borderProperties(color = "white")
      s9.ft <- setFlexTableWidths(s9.ft, widths = c(2, rep(1.5,ncol(s9.outputs.df)-1)))      
      s9.ft <- setZebraStyle(s9.ft, odd = purpleshade, even = "white" ) 
      s9.ft[,2:(ncol(s9.outputs.df))] <- parProperties(text.align="center")
      
      pptx.j <- addFlexTable(pptx.j, 
                             s9.ft, 
                             height = 2.75,
                             width = 8,
                             offx = 0.83,
                             offy = 2.0,
                             par.properties = parProperties(text.align="left", padding = 0)
      )
      
      #Notes
      s9.notes <- pot("Prompt Text:
                      1. (use cfa) I use common formative assessments aligned to the Missouri Learning Standards.
                      2. (all in cfa) All students in my classroom participate in common formative assessments, including students with disabilities.
                      3. (student reviews cfa) Each student reviews his/her results of common formative assessments with a teacher.
                      4. (cfa used to plan) I use the results from common formative assessment to plan for re-teaching and/or future instruction.",
                      textProperties(color = notesgrey, font.size = 10)
      )
      
      pptx.j <- addParagraph(pptx.j,
                             s9.notes,
                             height = 3.3,
                             width = 9.57,
                             offx = 0.22,
                             offy = 4.25,
                             par.properties = parProperties(text.align ="left", padding = 0)
      )
      
      writeDoc(pptx.j, file = target.file.j) #test Slide build up to this point
    }
    ## SLIDE 10 ##
    {  
      pptx.j <- addSlide( pptx.j, slide.layout = 'S2')
      
      #Title
      s10.title <- pot("CFA Scale Rates of Implementation",title.format.small)
      pptx.j <- addParagraph(pptx.j, 
                             s10.title, 
                             height = 0.89,
                             width = 8.47,
                             offx = 0.83,
                             offy = 0.85,
                             par.properties=parProperties(text.align="left", padding=0)
      )
      
      #Viz
      s10.graph <- ggplot(data = s10.outputs.df, aes(x = s10.outputs.df$category[order(s10.outputs.df$category, decreasing = TRUE)])) + 
        geom_bar(
          aes(y = score_school_rate), 
          stat="identity", 
          fill = backgroundgreen, 
          width = 0.8) + 
        ylim(0,100) +
        labs(x = "", y = "% Implementing 'Always' or 'Most of the time'") +                 #axis labels
        scale_x_discrete(labels = s10.headers.v[order(s10.headers.v, decreasing = TRUE)]) +
        geom_text(                                                                          #data labels
          aes(y = score_school_rate + 8, label = paste(s10.outputs.df$score_school_rate, "%", sep = "")), 
          size = 4,
          color = graphlabelsgrey) + 
        theme(panel.background = element_blank(),
              panel.grid.major.y = element_blank(),
              panel.grid.major.x = element_line(color = graphgridlinesgrey),
              axis.text.x = element_text(size = 15, color = graphlabelsgrey),
              axis.text.y = element_text(size = 15, color = graphlabelsgrey)
        ) +     
        coord_flip() 
      
      pptx.j <- addPlot(pptx.j,
                        fun = print,
                        x = s10.graph,
                        height = 3.25,
                        width = 8,
                        offx = 1,
                        offy = 2)
      
      #Notes
      s10.notes <- pot("Documented above are the percent of school respondents who answered that they use the following practices either “most of the time” or “always.” Please see page 9 for a list of complete prompts.",
                      notes.format)
      pptx.j <- addParagraph(pptx.j,
                             s10.notes,
                             height = 1.39,
                             width = 8.47,
                             offx = 0.83,
                             offy = 5.28,
                             par.properties = parProperties(text.align ="left", padding = 0)
      )
      
      writeDoc(pptx.j, file = target.file.j) #test Slide build up to this point
    }
    ## SLIDE 11 ##
    {  
      pptx.j <- addSlide( pptx.j, slide.layout = 'Section Header')
      
      #s11 Title
      s11.title <- pot("Diving Deeper Into Scales: DATA-BASED DECISION-MAKING",
                      section.title.format)
      pptx.j <- addParagraph(pptx.j, 
                             s11.title, 
                             height = 2.63,
                             width = 7.76,
                             offx = 1.25,
                             offy = 2.48,
                             par.properties=parProperties(text.align="left", padding=0)
      )
      writeDoc(pptx.j, file = target.file.j) #test Slide build up to this point
    }
    ## Slide 12 ##
    {
      pptx.j <- addSlide( pptx.j, slide.layout = 'S2')
      
      #Title
      s12.title <- pot("CFA Scale Performance",title.format)
      pptx.j <- addParagraph(pptx.j, 
                             s12.title, 
                             height = 0.89,
                             width = 8.47,
                             offx = 0.83,
                             offy = 0.18,
                             par.properties=parProperties(text.align="left", padding=0)
      )
      
      #Viz
      s12.ft <- FlexTable(
        data = s12.outputs.df,
        header.columns = TRUE,
        add.rownames = FALSE,
        
        header.cell.props = cellProperties(background.color = purpleheader),
        header.text.props = textProperties(color = "white", font.size = 16, font.weight = "bold"),
        
        body.cell.props = cellProperties(background.color = "white"), 
        body.text.props = textProperties(font.size = 15, font.weight = "bold", color = "#333333")
      )
      
      s12.ft[,] <- borderProperties(color = "white")
      s12.ft <- setFlexTableWidths(s12.ft, widths = c(2, rep(1.5,ncol(s12.outputs.df)-1)))      
      s12.ft <- setZebraStyle(s12.ft, odd = purpleshade, even = "white" ) 
      s12.ft[,2:(ncol(s12.outputs.df))] <- parProperties(text.align="center")
      
      pptx.j <- addFlexTable(pptx.j, 
                             s12.ft, 
                             height = 2.75,
                             width = 8,
                             offx = 0.22,
                             offy = 2.0,
                             par.properties = parProperties(text.align="left", padding = 0)
      )
      
      #Notes
      s12.notes <- pot("Prompt Text:
                      1. (use cfa) I use common formative assessments aligned to the Missouri Learning Standards.
                      2. (all in cfa) All students in my classroom participate in common formative assessments, including students with disabilities.
                      3. (student reviews cfa) Each student reviews his/her results of common formative assessments with a teacher.
                      4. (cfa used to plan) I use the results from common formative assessment to plan for re-teaching and/or future instruction.",
                      textProperties(color = notesgrey, font.size = 10)
      )
      
      pptx.j <- addParagraph(pptx.j,
                             s12.notes,
                             height = 3.3,
                             width = 9.57,
                             offx = 0.22,
                             offy = 4.25,
                             par.properties = parProperties(text.align ="left", padding = 0)
      )
      
      writeDoc(pptx.j, file = target.file.j) #test Slide build up to this point
    }  
    ## SLIDE 13 ##
    {  
      pptx.j <- addSlide( pptx.j, slide.layout = 'S2')
      
      #Title
      s13.title <- pot("DBDM Scale Rates of Implementation",title.format.small)
      pptx.j <- addParagraph(pptx.j, 
                             s13.title, 
                             height = 0.89,
                             width = 8.47,
                             offx = 0.83,
                             offy = 0.4,
                             par.properties=parProperties(text.align="left", padding=0)
      )
      
      #Viz
      s13.graph <- ggplot(data = s13.outputs.df, aes(x = s13.outputs.df$category[order(s13.outputs.df$category, decreasing = TRUE)])) + 
        geom_bar(
          aes(y = score_school_rate), 
          stat="identity", 
          fill = backgroundgreen, 
          width = 0.8) + 
        ylim(0,100) +
        labs(x = "", y = "% Implementing 'Always' or 'Most of the time'") +                 #axis labels
        scale_x_discrete(labels = s13.headers.v[order(s13.headers.v, decreasing = TRUE)]) +
        geom_text(                                                                          #data labels
          aes(y = score_school_rate + 8, label = paste(s13.outputs.df$score_school_rate, "%", sep = "")), 
          size = 4,
          color = graphlabelsgrey) + 
        theme(panel.background = element_blank(),
              panel.grid.major.y = element_blank(),
              panel.grid.major.x = element_line(color = graphgridlinesgrey),
              axis.text.x = element_text(size = 15, color = graphlabelsgrey),
              axis.text.y = element_text(size = 15, color = graphlabelsgrey)
        ) +     
        coord_flip() 
      
      pptx.j <- addPlot(pptx.j,
                        fun = print,
                        x = s13.graph,
                        height = 3.25,
                        width = 8,
                        offx = 1,
                        offy = 2)
      
      #Notes
      s13.notes <- pot("Documented above are the percent of school respondents who answered that they use the following practices either “most of the time” or “always.” Please see page 9 for a list of complete prompts.",
                       notes.format)
      pptx.j <- addParagraph(pptx.j,
                             s13.notes,
                             height = 1.39,
                             width = 8.47,
                             offx = 0.83,
                             offy = 5.4,
                             par.properties = parProperties(text.align ="left", padding = 0)
      )
      
      writeDoc(pptx.j, file = target.file.j) #test Slide build up to this point
    }
    ## SLIDE 14 ##
    {  
      pptx.j <- addSlide( pptx.j, slide.layout = 'Section Header')
      
      #s14 Title
      s14.title <- pot("Diving Deeper Into Scales: LEADERSHIP",
                       section.title.format)
      pptx.j <- addParagraph(pptx.j, 
                             s14.title, 
                             height = 2.63,
                             width = 7.76,
                             offx = 1.25,
                             offy = 2.48,
                             par.properties=parProperties(text.align="left", padding=0)
      )
      writeDoc(pptx.j, file = target.file.j) #test Slide build up to this point
    }
    
    ## SLIDE 17 ##
    {  
      pptx.j <- addSlide( pptx.j, slide.layout = 'Section Header')
      
      #s17 Title
      s17.title <- pot("Diving Deeper Into Scales: PROFESSIONAL DEVELOPMENT",
                       section.title.format)
      pptx.j <- addParagraph(pptx.j, 
                             s17.title, 
                             height = 2.63,
                             width = 7.76,
                             offx = 1.25,
                             offy = 2.48,
                             par.properties=parProperties(text.align="left", padding=0)
      )
      writeDoc(pptx.j, file = target.file.j) #test Slide build up to this point
    }
      
      
      
      
      

### EXPORTING RESULTS TO GOOGLE SHEETS ###
  
  #output.ss <- gs_new(title = "Cleaned Data", ws_title = "Cleaned Data")
  #gs_edit_cells(output.ss, ws = 1, input = dat.df, verbose = TRUE)
      
      
      
  #cwis.ss %>% gs_new("Cleaned Data", input = dat.df)#, verbose = TRUE)
 
  #gs_edit_cells(dat, ws='sheetname', input=colnames(result), byrow=TRUE, anchor="A1")
  #gs_edit_cells(dat, ws='sheetname', input = result, anchor="A2", col_names=FALSE, trim=TRUE)
  
  #} #END OF LOOP J BY SCHOOL
