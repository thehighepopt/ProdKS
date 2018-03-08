library(plyr, quietly = TRUE)
library(dplyr, quietly = TRUE)
library(openxlsx, quietly = TRUE)

setwd("C:/Users/35148/Documents/R/MAXIMUS/ProdKS/Daily") #set path to folder with all (and only) prof files in it

##get all files in folder, stick them all together
n <- list.files()
b <- data.frame()

for (i in n) {
  a <- read.xlsx(i,sheet = 3, detectDates = TRUE)  
  b <- rbind(b,a)
  rm(a)
}
rm(n,i)

prod <-  b[,c(1:3,5:8,14:16,20,22,23,25,30)] #remove unneeded columns

##subset to only the items that are work, remove activities, get eligibility queue
tasks <- subset(prod,(Event == "Submitted") & (Timer > 0) & (Task.Queue == "Eligibility"))
      ##Test to see who has no tasks if you cut the Timer off at 5 or 2 minutes
      # t <- tasks %>% 
      #   group_by(Name,Functional.Area,Task.Type) %>%
      #   summarise(Taskcount = n())
      # w <- work %>% 
      #   group_by(Name,Functional.Area,Task.Type) %>%
      #   summarise(Taskcount = n())
      # tw <- merge(t,w,by = c("Name","Functional.Area","Task.Type"),all = TRUE, sort = TRUE)

work1 <- subset(tasks, Task.Type == 'Application' | Task.Type == 'Review')
work1 <- subset(work1, Timer > 300 & Timer < 20000)
work2 <- subset(tasks,Task.Type == 'Batch' | Task.Type == 'Research' | Task.Type == 'Case Maintenance')
work2 <- subset(work2, Timer > 120 & Timer < 20000)
work <- rbind(work1,work2)
rm(work1,work2)
rownames(work) <- 1:nrow(work)

##Reduce, arrange, add percentile, find values for percentile
perc <- work[,c(11,12,9)]
perc <- perc %>% 
  group_by(Functional.Area,Task.Type) %>%
     arrange(Timer,.by_group = TRUE) %>% 
         mutate(Percentile = cume_dist(Timer))

perc$Percentile <- round(perc$Percentile,3)
percs <- filter(perc,(Percentile >= .300 & Percentile < .31) | (Percentile >= .600 & Percentile < .61)
                | (Percentile >= .900 & Percentile < .91)) ##Need between because low volume tasks don't land evenly, may not be an issue with a month of data
percs$Percentile <- round(percs$Percentile,1)

percentile <- percs %>% 
  group_by(Functional.Area,Task.Type,Percentile) %>%
    summarise(HandleTime = median(Timer)) ##could use MIN here as well
colnames(percentile) <- c("Area" ,"Task.Type", "Percentile", "HandleTime")
percentile$HandleTime  <- round(percentile$HandleTime,0)
rm(perc,percs)

##write to file
setwd("C:/Users/35148/Documents/R/MAXIMUS/ProdKS")
current_date <- Sys.Date()
write.csv(percentile,paste("All_Feb_Percentile_Ranks_",current_date,".csv"),row.names = FALSE)
rm(current_date)

#############################################################################################
##Get AHT per task type per staff
staff  <- work[,c(1,11,12,9)]
AHT <- staff %>% 
  group_by(Name,Functional.Area,Task.Type) %>%
    summarise(AHT = mean(Timer), Count = n()) ##aht and count of tasks
colnames(AHT) <- c("Name", "Area" ,"Task.Type", "AHT","Count")
AHT$AHT  <- round(AHT$AHT,0)
AHT$Area <- as.character(AHT$Area)
AHT$Task.Type <- as.character(AHT$Task.Type)
AHT$Name <- as.character(AHT$Name)


##Apply Rating
Ratings <- AHT
Ratings$Rating <- NA  #add a column to store rating

for (i in 1:nrow(Ratings)) {
  s <- Ratings[i,]   
  
  p <- subset(percentile,Area == as.character(s[1,2]) & (Task.Type == as.character(s[1,3])))
  
  h <- as.numeric(ifelse(s[1,4] < p[1,4], 1, 
        ifelse(s[1,4] >= p[1,4] & s[1,4] < p[2,4], 2, 
              ifelse(s[1,4] >= p[2,4] & s[1,4] < p[3,4], 3,4))))
  Ratings$Rating[i] <- h
 
}  #loops through each aht per task per person and applies the rating based on percentiles

#Show count of ratings by value 
  r <-Ratings %>% 
    group_by(Rating) %>%
    summarise(Count = n())

  r <- mutate(r, grade_pct = Count / sum(Count))
  r$grade_pct <- round(r$grade_pct,4) * 100
  r

##Then get averages by staff
avgrating <- Ratings %>% 
  group_by(Name) %>%
  summarise(Avg_Rating = mean(Rating))
avgrating$Avg_Rating  <- round(avgrating$Avg_Rating,2)
 
##Apply the grad to avg Rating
grade <- data.frame(Grade = c("A","B","C","D","E"), avgRating = c(1.5,2.2,2.8,3.2,4)) ###adjust Rating if needed
performance <- avgrating
performance$Grade <- NA

  for (i in 1:nrow(performance)) {
    s <- performance[i,]   
    
    h <- as.character(ifelse(s[1,2] < grade[1,2], "A", 
        ifelse(s[1,2] >= grade[1,2] & s[1,2] < grade[2,2], "B", 
            ifelse(s[1,2] >= grade[2,2] & s[1,2] < grade[3,2], "C",
                 ifelse(s[1,2] >= grade[3,2] & s[1,2] < grade[4,2], "D","E")))))
    
    performance$Grade[i] <- h
  }

#Show count of grades by value 
  g <- performance %>% 
     group_by(Grade) %>%
     summarise(Count = n())

  g <- mutate(g, grade_pct = Count / sum(Count))
  g$grade_pct <- round(g$grade_pct,4) * 100
  g

##Write to file with multiple tabs
library(xlsx, quietly = TRUE)
setwd("C:/Users/35148/Documents/R/MAXIMUS/ProdKS/Performance")

xlsx.writeMultipleData <- function (file, ...) ##Create a function that will write
{
  require(xlsx, quietly = TRUE)
  objects <- list(...)
  fargs <- as.list(match.call(expand.dots = TRUE))
  objnames <- as.character(fargs)[-c(1, 2)]
  nobjects <- length(objects)
  for (i in 1:nobjects) {
    if (i == 1)
      write.xlsx(objects[[i]], file, sheetName = objnames[i])
    else write.xlsx(objects[[i]], file, sheetName = objnames[i],
                    append = TRUE)
  }
}

#run the function
xlsx.writeMultipleData("Performance Feb 2018.xlsx", as.data.frame(Ratings), as.data.frame(percentile),
                       as.data.frame(performance), as.data.frame(grade), as.data.frame(r), as.data.frame(g))


####merge everything together
library(bizdays)
o <-work %>% 
  group_by(Name) %>%
  summarise(Task.Count = n())
a <- work %>% 
  group_by(Name) %>%
  summarise(AHT = mean(Timer))
perfsup <- merge(unique(work[1:2]),performance1)
perfsup <- merge(perfsup,o)
perfsup <- merge(perfsup,a)
perfsup$AHT <- round(perfsup$AHT,0)
perfsup$TaskxDay <- round(perfsup$Task.Count / 19,1)


###Analyze short tasks####################
short <- subset(work, (Timer < 181) & ((Task.Type == 'Application') | (Task.Type == 'Review')) )
shorts <- short %>% 
  group_by(Name,Functional.Area,Task.Type) %>%
  summarise(AHT = mean(Timer), Count = n())
shorts$AHT  <- round(shorts$AHT,0)
all <- subset(work,  (Task.Type == 'Application') | (Task.Type == 'Review')) 
all <- all %>% 
  group_by(Name,Functional.Area,Task.Type) %>%
  summarise(AHT = mean(Timer), Count = n())
all$AHT  <- round(all$AHT,0)

bbb <- merge(shorts, all, by = c("Name","Functional.Area","Task.Type"),all = TRUE, sort = TRUE)

colnames(bbb) <- c("Name", "Area" ,"Task.Type", "AHT Short","Count Short","AHT All","Count All")
###################################



















