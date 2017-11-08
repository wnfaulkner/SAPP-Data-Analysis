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
    ans.opt.always.df[,1] <- ans.opt.always.df[,1] %>% as.character %>% as.numeric
    ans.opt.always.df[,2] <- ans.opt.always.df[,2] %>% as.character
    ans.opt.always.df[,3] <- ans.opt.always.df[,3] %>% as.character
    
    slide.names.v <- c("Participation Details")
  
### PRODUCING DATA, TABLES, & CHARTS ###

  i <- 1 # for testing loop
  
  #for(i in l: length(school.names)){   #START OF LOOP BY SCHOOL
   
    # Create data frame for this loop - restrict to responses from school name i
      school.name.i <- school.names[i]
      dat.df.i <- dat.df[dat.df$school.name == school.name.i,] %>% tbl_df
      district.name.i <- dat.df.i$district.name %>% unique
    
    #S2 Table for slide 2 "Participation Details"
      s2.mtx <- table(dat.df.i$role) %>% as.matrix
      s2.df <- cbind(row.names(s2.mtx),s2.mtx[,1]) %>% as.data.frame
      names(s2.df) <- c("Role","Num Responses")
      rownames(s2.df) <- c()
      s2.df$Role <- s2.df$Role %>% as.character #convert factor to character
      s2.df$`Num Responses` <- s2.df$`Num Responses` %>% as.character %>% as.numeric #convert factor to numeric
      s2.df <- rbind(s2.df, c("Total", sum(s2.df[,2] %>% as.character %>% as.numeric))) #add "Total" row
      
          
    #S3 Table for slide 3 "Overall Scale Performance"
    
      #Define variables to average for each 
      etl.vars <- 
  
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
      names(s6.df) <- c("Answer option","1. learning targets","2. students assess","3. students identify","4. feedback to targets","5. student to student feedback","6. students state criteria", "7. instruction state standards", "student reviews cfa")
        
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
    purpleshade <- "#d0abd6"
    purpleheader <- "#3d2242"
    backgroundgreen <- "#94c132"
    subtextgreen <- "#929e78"
    #notesgray <- rgb(131,130,105, maxColorValue=255)
    
    #Text formatting
      title.format <- textProperties(color = titlegreen, font.size = 54, font.weight = "bold")  
      subtitle.format <- textProperties(color = notesgrey, font.size = 22, font.weight = "bold")
      section.title.format <- textProperties(color = "white", font.size = 48, font.weight = "bold")
      notes.format <- textProperties(color = notesgrey, font.size = 14)
  
  #Edit powerpoint template

    ## SLIDE 1 ##
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
      
      
    ## SLIDE 2 ##
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
      
      #writeDoc(pptx.j, file = target.file.j) #test Slide 2 build
    
        
    ## SLIDE 3 ##
      
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
      
      #Notes
      s3.notes <- pot("Note: A score of 4 represents a response of "most of the time" for Effective Teaching and Learning Practices, Common Formative Assessment, and Data-based Decision-making scales, and "agree" for the Leadership, and Professional Development scale.",
                      notes.format)
      pptx.j <- addParagraph(pptx.j,
                             s3.notes,
                             height = 1.39,
                             width = 8.47,
                             offx = 0.83,
                             offy = 5.28,
                             par.properties = parProperties(text.align ="left", padding = 0)
      )
      
      
    ## SLIDE 4 ##
      
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
      s4.notes <- pot("A score of 3 represents a response of "about half the time" for Effective Teaching and Learning, Common Formative Assessment, and Data-based Decision-making scales, while a 2 represents "sometimes". 
                      A 3 corresponds to "neither agree or disagree" for the Leadership, and Professional Development scale while a 2 represents "disagree."",
                      notes.format)
      pptx.j <- addParagraph(pptx.j,
                             s4.notes,
                             height = 1.39,
                             width = 8.47,
                             offx = 0.83,
                             offy = 5.28,
                             par.properties = parProperties(text.align ="left", padding = 0)
      )
      

    ## SLIDE 5 ##
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
      writeDoc(pptx.j, file = target.file.j)
    
    ## Slide 6 ##
    
      pptx.j <- addSlide( pptx.j, slide.layout = 'S2')
      
      #Title
      s6.title <- pot("ETLP Scale Performance",title.format)
      pptx.j <- addParagraph(pptx.j, 
                             s6.title, 
                             height = 0.89,
                             width = 8.47,
                             offx = 0.83,
                             offy = 0.85,
                             par.properties=parProperties(text.align="left", padding=0)
      )
      
      #Viz
      s6.ft <- FlexTable(
        data = s6.df,
        header.columns = TRUE,
        add.rownames = FALSE,
        
        header.cell.props = cellProperties(background.color = purpleheader),
        header.text.props = textProperties(color = "white", font.size = 14, font.weight = "bold"),
        
        body.cell.props = cellProperties(background.color = "white"), 
        body.text.props = textProperties(font.size = 12, font.weight = "bold")
      )
      
      s6.ft <- setFlexTableWidths(s6.ft, widths = c(1, rep(0.8,ncol(s6.df)-1)))      
      s6.ft <- setZebraStyle(s6.ft, odd = purpleshade, even = "white" ) 
      
      pptx.j <- addFlexTable(pptx.j, 
                             s6.ft, 
                             height = 1.72,
                             width = 7.77,
                             offx = 1.1,
                             offy = 1.29,
                             par.properties=parProperties(text.align="left", padding=0)
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
                      notes.format)
      pptx.j <- addParagraph(pptx.j,
                             s6.notes,
                             height = 3.37,
                             width = 9.57,
                             offx = 0.22,
                             offy = 3.25,
                             par.properties = parProperties(text.align ="left", padding = 0)
      )
      
    
    
 




### EXPORTING RESULTS TO GOOGLE SHEETS ###
  
  #output.ss <- gs_new(title = "Cleaned Data", ws_title = "Cleaned Data")
  #gs_edit_cells(output.ss, ws = 1, input = dat.df, verbose = TRUE)
      
      
      
  #cwis.ss %>% gs_new("Cleaned Data", input = dat.df)#, verbose = TRUE)
 
  #gs_edit_cells(dat, ws='sheetname', input=colnames(result), byrow=TRUE, anchor="A1")
  #gs_edit_cells(dat, ws='sheetname', input = result, anchor="A2", col_names=FALSE, trim=TRUE)
  
  #} #END OF LOOP J BY SCHOOL
