###Duplicate task counts
setwd("C:/Users/35148/Documents/R")
library(plyr)
library(dplyr)

bpm3 <- read.csv("bpmmay16nov16.csv")
bpm1 <- read.csv("bpmdec16jun17.csv")
bpm2 <- read.csv("bpmjul17dec17.csv")
bpm <- rbind(bpm3,bpm1,bpm2)

bpmsm <- bpm[,c(5,8,15)]
colnames(bpmsm) <- c('a','b','c')

f <- ddply(bpmsm,.(a,b,c),nrow)
colnames(f) <- c('Task','Date','Casenum','Count')
f$Date <- as.Date(f$Date,"%m/%d/%Y")
f$Month <- format(f$Date,"%B-%Y")
f <- subset(f,!is.na(f$Casenum))

g <- ddply(f,.(Month,Count),nrow)
h <- subset(g,Count != 1)
colnames(h) <- c('Month','Tasks Created','Count')

write.csv(h,"DupTasks Jun-Nov 2016.csv",row.names = FALSE)

##i <- subset(f, Count > 1)
##write.csv(i,"DupTasksDetail Jul-Dec 2017.csv",row.names = FALSE)

############################Count of Passive Review tasks
p <- subset(bpmsm, a == 'Passive Review Response')
q <- subset(bpmsm, a == 'Passive Review Response - Manual')
r <- rbind(p,q)
t <- ddply(r,.(a,b),nrow)
t$b <- as.Date(t$b,"%m/%d/%Y")

write.csv(t,"Passive Reviews May16-Dec17.csv",row.names = FALSE)

summary(t$V1)
y <- subset(t, b >= as.Date('2017-01-01'))
summary(y$V1)     

###################A&R Reviews
setwd("C:/Users/35148/Documents/R/MAXIMUS/AR Detail")
library(openxlsx)

n <- list.files()
b <- data.frame()

for (i in n) {
  a <- read.xlsx(i)  
  b <- rbind(b,a)
  rm(a)
}

c <- b[,c(8,10,13)]
c$RECEIVED_DATE <- as.Date(c$RECEIVED_DATE, origin="1899-12-30")
c <- subset(c,REQUEST_TYPE == 'REVIEW')
c <- subset(c,RECEIVED_DATE >= '2017-01-01')
c <- unique(c[ , 1:3 ])
d <- ddply(c,.(RECEIVED_DATE,REQUEST_TYPE),nrow)

e <- b[,c(8,10,13)]
e$RECEIVED_DATE <- as.Date(e$RECEIVED_DATE, origin="1899-12-30")
e <- subset(e,REQUEST_TYPE == 'APPLICATION')
e <- subset(e,RECEIVED_DATE >= '2017-01-01')
e <- unique(e[ , 1:3 ])

f <- ddply(e,.(RECEIVED_DATE,REQUEST_TYPE),nrow)

setwd("C:/Users/35148/Documents/R/MAXIMUS")
write.csv(d,"Reviews by Day using December AR Reports.csv")
write.csv(f,"Apps by Day using EOM AR Reports.csv")

################Capture all AR Detail reports in a month
setwd("P:\\Reports\\Archive\\2017 Daily Reports\\05 - May 2017")
library(openxlsx)
library(plyr)
library(dplyr)

n <- list.files()
n <- n[-4]
b <- as.data.frame(matrix(nrow=0,ncol=34))

for (i in n) {
  setwd(paste("P:\\Reports\\Archive\\2017 Daily Reports\\05 - May 2017\\",i,sep = ""))
  z <- getSheetNames("Application and Review Detail.xlsx")
  a <- read.xlsx("Application and Review Detail.xlsx",length(z))  
  b <- rbind(a,setNames(b,names(a)))
  rm(a)
  }

setwd("P:\\Reports\\Archive\\2017 Daily Reports\\11 - November 2017")

c <- unique(b[,1:27])
c <- b[,c(8,10,12,13,30,34)]  ##c(2,4,6,7,24,28)
c <- unique(c[,1:5])

c$RECEIVED_DATE <- as.Date(c$RECEIVED_DATE, origin="1899-12-30")
b$RUN_DATE <- as.Date(b$RUN_DATE, origin="1899-12-30")

d <- subset(c,RECEIVED_DATE >= '2017-01-01')
e <- ddply(d,.(RECEIVED_DATE,REQUEST_TYPE),nrow)
setwd("C:/Users/35148/Documents/R/MAXIMUS")
write.csv(e,"Docs by Day Using AR Reports 2017 all3.csv")


#########Capture all AR Detail reports from 1st day of month
setwd("P:\\Reports\\Archive\\2017 Daily Reports")
library(openxlsx)
library(plyr)
library(dplyr)

n <- list.files()
n <- n[-1]
q <- data.frame()

for (i in n) {
  setwd(paste("P:\\Reports\\Archive\\2017 Daily Reports\\",i,sep = ""))
  o <- list.files()[1]
  p <- getwd()
  setwd(paste(p,"/",o,sep=""))
    z <- getSheetNames("Application and Review Detail.xlsx")
    a <- read.xlsx("Application and Review Detail.xlsx",length(z))  
    q <- rbind(b,a) 
    
  rm(a)
}






