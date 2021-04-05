#Lewis's i-Return Receiving File  -------------1
setwd("C:/Users/Weilun_Chiu/Documents/3PAR/iReturn")
filename<-list.files()
iReturn_Data<-do.call(rbind, lapply(filename, read.csv, header=TRUE))

myda<-iReturn_Data[, c("DISP.LOCATION", "DATE.SENT", "PART.SERNO")]
myda$DATE.SENT<-as.POSIXct(strptime(myda$DATE.SENT, "%m/%d/%Y %H:%M"))
myda<-myda %>% arrange(desc(DATE.SENT))
myda<-myda %>% distinct(PART.SERNO, .keep_all = TRUE)
mydaF<-data.frame(myda$PART.SERNO, myda$DISP.LOCATION)
mydaF$myda.PART.SERNO<-as.character(mydaF$myda.PART.SERNO)
mydaF$myda.DISP.LOCATION<-as.character(mydaF$myda.DISP.LOCATION)


#check<-function(x){
#  loc<-which(mydaF$myda.PART.SERNO==x)
#  warranty<-mydaF$myda.DISP.LOCATION[loc]
#  return(ifelse(warranty=="RT23", "IN", "OUT"))
# }

#Lewis's i-Return Receiving File  -------------1

#First match all SNs with Lewis's iReturn file to have Warranty Status
library(readxl)
library(writexl)
library(dplyr)
library(stringr)
library(tidyr)

## Remember to sort fault tag column in order to prevent the system automatically coerce the class of it to be logic

setwd("C:/Users/Weilun_Chiu/Desktop/QWE/CID_VI")
VI<-read_xlsx("VI.xlsx")
VI1<-filter(VI, Status=="CID")
colnames<-names(VI1)
colnames[2]<-"RMA"
names(VI1)<-colnames
VI2<-VI1 %>% filter(str_detect(RMA, "SVC"))
VI2_1<-VI2 %>% filter(!grepl("APD", RMA))
VI2_2<-VI2_1 %>% filter(!grepl("CZ", RMA))

SN<-VI2_2$S.N.
SN<-substr(SN, 6, 14)
len<-length(SN)

Result<-c(seq(1, len,1))

for(i in 1:len){
  num<-grep(SN[i], mydaF$myda.PART.SERNO)
  if(length(num)>0){
    Result[i]<-mydaF[num, 2]
  }else{
    Result[i]<-"NA"
  }
}

  
a<-grep("PCMBU", VI2_2$S.N.)
b<-grep("PDHWN", VI2_2$S.N.)
c<-grep("PDSET", VI2_2$S.N.)
d<-grep("PFLKQ", VI2_2$S.N.)
POSI<-sort(c(a,b,c,d))
len2<-length(POSI)

setwd("C:/Users/Weilun_Chiu/Documents/3PAR/Assy")
file<-list.files()
Assy<-read.csv(file)
Assy$Assy.<-as.character(Assy$Assy.)
Assy$SPS.<-as.character(Assy$SPS.)
SN<-VI2_2$S.N.


for(i in 1:len2){
  num<-which(SN[POSI[i]]==Assy$Assy.)
  NodeSN<-Assy[num,2]
  num2<-which(NodeSN==mydaF$myda.PART.SERNO)
  if(length(num2)>0){
    Result[POSI[i]]<-mydaF$myda.DISP.LOCATION[num2]
  }else{
    Result[POSI[i]]<-"NA"
  }
}

colnames[8]<-"Warranty"
names(VI2_2)<-colnames
VI2_2$Warranty<-Result
VI_TEMP<-filter(VI2_2, Warranty!="NA")
VI_TEMP$Warranty<-ifelse(VI_TEMP$Warranty=="RT23", "IN", "OUT")
names(VI_TEMP)[3]<-"PCA#"; names(VI_TEMP)[4]<-"SN#"; names(VI_TEMP)[10]<-"Date"; names(VI_TEMP)[5]<-"Comment"; names(VI_TEMP)[13]<-"Fault_tag"
VI_TEMP<-VI_TEMP %>% filter(Warranty=="IN")
VI_TEMP1<-VI_TEMP %>% select(Date, RMA, 'PCA#', 'SN#', Comment, Warranty, MROC, Fault_tag)
TEMP<-list(TEMP=VI_TEMP1)
setwd("C:/Users/Weilun_Chiu/Desktop/QWE/CID_VI")
write_xlsx(TEMP, "TEMP.xlsx")

#################################################################################### Modification CID Tab
#################################################################################### Modification CID Tab
#################################################################################### Modification CID Tab
#################################################################################### Modification CID Tab
#################################################################################### Modification CID Tab
#################################################################################### Modification CID Tab
#################################################################################### Modification CID Tab
#################################################################################### Modification CID Tab
####### Make sure no blank space, MROC 3, TO SPECIAL PACKING, etc.



VI_TEMP1<-read_xlsx("TEMP.xlsx")

Comment<-VI_TEMP1$Comment
logic<-rep(0, times=length(Comment))
list<-gregexpr(pattern="\\|\\|", Comment)
  for(i in 1:length(Comment)){
    if(list[[i]][1]!=-1){
      logic[i]<-ifelse(length(list[[i]])==1, 1, 2)
    }else{
      logic[i]==0
    }
  }

temp<-data.frame(Error=Comment, Len=logic, Found1=0, Found2=0, Found3=0)
loc1<-which(temp$Len==0)
for(i in 1:length(loc1)){
  temp$Found1[loc1[i]]<-strsplit(temp$Error[loc1[i]], ":")[[1]][1]
}

loc2<-which(temp$Len==1)
split1<-strsplit(temp$Error[loc2], " \\|\\| ")
for(i in 1:length(loc2)){
  temp$Found1[loc2[i]]<-strsplit(split1[[i]],":")[[1]][1]
  temp$Found2[loc2[i]]<-strsplit(split1[[i]],":")[[2]][1]
}

loc3<-which(temp$Len==2)
split2<-strsplit(temp$Error[loc3], " \\|\\| ")
for(i in 1:length(loc3)){
  temp$Found1[loc3[i]]<-strsplit(split2[[i]],":")[[1]][1]
  temp$Found2[loc3[i]]<-strsplit(split2[[i]],":")[[2]][1]
  temp$Found3[loc3[i]]<-strsplit(split2[[i]],":")[[3]][1]
}

CID_QTY<-dim(VI_TEMP1)[1]
VI_QTY<-dim(VI)[1]
CID_PER<-(CID_QTY/VI_QTY)
Count<-as.data.frame(t(data.frame(CID_QTY=CID_QTY, VI_QTY=VI_QTY, CID_PER=CID_PER)))


temp1<-temp %>% select(Found1, Found2, Found3)
temp2<-gather(temp1, "Found", "Desc.", 1:3)
Desc_count<-as.data.frame(table(temp2$Desc.))
Desc_count<-Desc_count %>% arrange(desc(Freq))

AC<-substr(VI_TEMP1$`PCA#`, 1, 6)
AC1<-as.data.frame(table(AC))
AC1<-AC1 %>% arrange(desc(Freq))
AC2<-AC1[1:3,]
AC2[,1]<-as.character(AC2[,1])
AC3<-rbind(AC2, c("Orthers", length(AC)-sum(AC2[,2])))
AC4<-rbind(AC3, c("Total", length(AC)))

Rank1<-as.character(AC1[1,1])
Rank2<-as.character(AC1[2,1])
Rank3<-as.character(AC1[3,1])

AD<-cbind(VI_TEMP1$`PCA#`, temp1)
colnames(AD)[1]<-"PCA#"
AD$`PCA#`<-substr(AD$`PCA#`, 1, 6)

AD_Rank1<-subset(AD, AD$`PCA#`==Rank1)
AD_Rank1<-gather(AD_Rank1, "Found", "Desc.", 2:4)
Desc_count_rank1<-as.data.frame(table(AD_Rank1$Desc.))
Desc_count_rank1<-Desc_count_rank1 %>% arrange(desc(Freq))

AD_Rank2<-subset(AD, AD$`PCA#`==Rank2)
AD_Rank2<-gather(AD_Rank2, "Found", "Desc.", 2:4)
Desc_count_rank2<-as.data.frame(table(AD_Rank2$Desc.))
Desc_count_rank2<-Desc_count_rank2 %*% arrange(desc(Freq))

AD_Rank3<-subset(AD, AD$`PCA#`==Rank3)
AD_Rank3<-gather(AD_Rank3, "Found", "Desc.", 2:4)
Desc_count_rank3<-as.data.frame(table(AD_Rank3$Desc.))
Desc_count_rank3<-Desc_count_rank3 %>% arrange(desc(Freq))
  
Final_Data<-list(Original_data=VI_TEMP, CID_IW=VI_TEMP1, Count=Count, Desc.count=Desc_count, Assembly_Count=AC4,
                 Rank1=Desc_count_rank1, Rank2=Desc_count_rank2, Rank3=Desc_count_rank3)

setwd("C:/Users/Weilun_Chiu/Desktop/QWE/CID_VI")
write_xlsx(Final_Data, "CI_VI_Final.xlsx")
