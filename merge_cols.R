


#### CREATE A TOTAL WORKING CODE FROM ORIGINAL DOCUMENTATION
dat_raw <- read.csv("C:/Users/Jason Altman/Dropbox (TerraLuna)/Missouri CW (SPDG)/Missouri Model Districts/Evaluation Methods/Quantitative/CWIS/Survey Data/Count of School_District Participants.csv", 
                 header=TRUE, stringsAsFactors = FALSE, strip.white = TRUE, na.strings = c("NA", ""))


#delete first two rows
dat = tail(dat_raw, -2)

#add blank columns to fill values for Q3 and Q4
cols_name = c("q3_1","q3_2","q3_3","q3_4","q3_5","q3_6","q3_7","q3_8", "q4_1","q4_2","q4_3","q4_4","q4_5","q4_6","q4_7","q4_8", "q4_9", "q4_10", "q4_11")
dat[,cols_name] <- NA

#zero all NA's for q3 and q4 per survey data for sum checks ### NOTE: All columns name that begin with a number in R are converted to begin with an "X"
which( colnames(dat)=="X2_Q10_1" )
match("X1_Q10_1",names(dat)) #25
match("X8_Q10_8",names(dat)) #88
match("X1_Q4_1",names(dat))  #89
match("X8_Q4_11",names(dat)) #176

#x[c("a", "b")][is.na(x[c("a", "b")])] <- 0; change all values to numeric so that conditions can be met 
require(data.table)
for (col in 25:88) set(dat, which(is.na(dat[[col]])), col, 0)
dat[,25:88] = ifelse(dat[,25:88] == "Strongly agree", 5, 
                     ifelse(dat[,25:88] == "Agree", 4,
                            ifelse(dat[,25:88] == "Neither agree or disagree", 3,
                                   ifelse(dat[,25:88] == "Disagree", 2,
                                          ifelse(dat[,25:88] == "Strongly disagree", 1, 0)))))

for (col in 89:176) set(dat, which(is.na(dat[[col]])), col, 0)
dat[,89:176] = ifelse(dat[,89:176] == "Always", 5, 
                      ifelse(dat[,89:176] == "Most of the time", 4,
                             ifelse(dat[,89:176] == "About half the time", 3,
                                    ifelse(dat[,89:176] == "Sometimes", 2,
                                           ifelse(dat[,89:176] == "Never", 1, 0)))))




#check if each segement within q3 is non-zero, if so, copy to cols 183:190
dat[,183] <- ifelse(rowSums(dat[25:32]) > 0, dat[,25],
                    ifelse(rowSums(dat[33:40]) > 0, dat[,33],
                           ifelse(rowSums(dat[41:48]) > 0, dat[,41],
                                  ifelse(rowSums(dat[49:56]) > 0, dat[,49],
                                         ifelse(rowSums(dat[57:64]) > 0, dat[,57],
                                                ifelse(rowSums(dat[65:72]) > 0, dat[,65],
                                                       ifelse(rowSums(dat[73:80]) > 0, dat[,73],
                                                              ifelse(rowSums(dat[81:88]) > 0, dat[,81],0))))))))

# dat[,184] <- ifelse(rowSums(dat[33:40]) > 0, dat[,34],
#                     ifelse(rowSums(dat[41:48]) > 0, dat[,42],
#                            ifelse(rowSums(dat[49:56]) > 0, dat[,50],
#                                   ifelse(rowSums(dat[57:64]) > 0, dat[,58],
#                                          ifelse(rowSums(dat[65:72]) > 0, dat[,66],
#                                                 ifelse(rowSums(dat[73:80]) > 0, dat[,74],
#                                                        ifelse(rowSums(dat[81:88]) > 0, dat[,82],0)))))))

for (i in c(1:7)) {
  dat[,183 + i] <- ifelse(rowSums(dat[25:32]) > 0, dat[,25 +i],
                          ifelse(rowSums(dat[33:40]) > 0, dat[,33 + i],
                                 ifelse(rowSums(dat[41:48]) > 0, dat[,41 + i],
                                        ifelse(rowSums(dat[49:56]) > 0, dat[,49 + i],
                                               ifelse(rowSums(dat[57:64]) > 0, dat[,57 + i],
                                                      ifelse(rowSums(dat[65:72]) > 0, dat[,65 + i],
                                                             ifelse(rowSums(dat[73:80]) > 0, dat[,73 + i],
                                                                    ifelse(rowSums(dat[81:88]) > 0, dat[,81 + i],0))))))))
}

#j = 11
for (i in c(0:10)) {
  dat[,191 + i] <- ifelse(rowSums(dat[89:99]) > 0, dat[,89 + i],
                          ifelse(rowSums(dat[100:110]) > 0, dat[,100 + i],
                                 ifelse(rowSums(dat[111:121]) > 0, dat[,111 + i],
                                        ifelse(rowSums(dat[122:132]) > 0, dat[,122 + i],
                                               ifelse(rowSums(dat[133:143]) > 0, dat[,133 + i],
                                                      ifelse(rowSums(dat[144:154]) > 0, dat[,144 + i],
                                                             ifelse(rowSums(dat[155:165]) > 0, dat[,155 + i],
                                                                    ifelse(rowSums(dat[166:176]) > 0, dat[,166 + i],0))))))))
  
}



### Convert numbers back to words
dat[,183:190] = ifelse(dat[,183:190] ==5, "Strongly agree",  
                       ifelse(dat[,183:190] == 4, "Agree", 
                              ifelse(dat[,183:190] == 3, "Neither agree or disagree",
                                     ifelse(dat[,183:190] == 2, "Disagree",
                                            ifelse(dat[,183:190] == 1, "Strongly disagree", NA)))))


dat[,191:201] = ifelse(dat[,191:201] == 5, "Always", 
                       ifelse(dat[,191:201] == 4, "Most of the time",
                              ifelse(dat[,191:201] == 3, "About half the time",
                                     ifelse(dat[,191:201] == 2, "Sometimes", 
                                            ifelse(dat[,191:201] == 1, "Never", NA)))))

write.csv(dat, "C:/Users/Jason Altman/Documents/R/4.csv", row.names = TRUE)
