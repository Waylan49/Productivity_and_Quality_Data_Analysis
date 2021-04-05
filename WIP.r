library(readxl)
library(writexl)
library(dplyr)

setwd("C:/Users/Weilun_Chiu/Documents/3PAR/iReturn")
filename<-list.files()
iReturn_Data<-do.call(rbind, lapply(filename, read.csv, header=TRUE))

write.csv(iReturn_Data, "Current.csv")

mydata<-iReturn_Data %>% select(DISP.LOCATION, DATE.SENT, PART.SERNO)
mydata$DATE.SENT<-as.POSIXct(strptime(mydata$DATE.SENT, "%m/%d/%Y %H:%M"))
mydata1<-mydata %>% filter(DISP.LOCATION %in% c("RT23", "RT24"))
mydata2<-mydata1 %>% arrange(desc(DATE.SENT))
mydata3<-mydata2 %>% distinct(PART.SERNO, .keep_all = TRUE)
mydata3$NEW.SER<-substr(mydata3$PART.SERNO, 8, 14)



######################################################################################################################
#####mydata3 is the complete result we need for warranty check from Lewis's file, use write_xlsx to write it out######
###############################write_xlsx(mydata3, "lewis.xlsx")######################################################
######################################################################################################################

setwd("C:/Users/Weilun_Chiu/Documents/3PAR/SN")
SN<-read_xlsx("SN.xlsx")
SN$NEW.SN<-substr(SN$Serial.Number, 8, 14)

len<-dim(SN)[1]
Warranty<-rep("NA", len)

for(i in 1:len){
        loc<-which(SN$NEW.SN[i]==mydata3$NEW.SER)
        if(length(loc)>0){
                Warranty[i]<-mydata3[loc, 1]
        }
}

SN$Warranty<-Warranty
Miss<-which(SN$Warranty=="NA")


setwd("C:/Users/Weilun_Chiu/Documents/3PAR/Assy")
filename1<-list.files()
ThreePAR<-read.csv(filename1)

SN$ThreePar<-ifelse(SN$Warranty=="NA","Maybe", "Not_3PAR")

len2<-length(Miss)

for(i in Miss){
        loc<-which(SN$Serial.Number[i]==ThreePAR$Assy.)
        if(length(loc)>0){
              SN$ThreePar[i]<-ThreePAR$SPS.[loc]
              warrantyloc<-which(ThreePAR$SPS.[loc]==mydata3$PART.SERNO)
              SN$Warranty[i]<-mydata3$DISP.LOCATION[warrantyloc]
        }
}

IW<-which(SN$Warranty=="RT23")
SN$Warranty[IW]<-"Y"
OOW<-which(SN$Warranty=="RT24")
SN$Warranty[OOW]<-"N"

SN$Week<-substr(SN$Serial.Number, 10, 11)

setwd("C:/Users/Weilun_Chiu/Documents/3PAR/SN")
write_xlsx(SN, "SN_Checked.xlsx")



