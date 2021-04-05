runway<-test

Fault<-select(runway, Fault.Found.1, Fault.Found.2, Fault.Found.3, Fault.Found.4, Fault.Found.5)

Result1<-as.data.frame(table(runway$MROC))
names(Result1)[1]<-"MROC"

temp<-select(runway, MROC, AssemblyNo)
temp1<-temp %>% group_by(AssemblyNo) %>% summarise(TotalCount=n())
temp1<-temp1 %>% arrange(desc(TotalCount))

rank<-as.character(temp1$AssemblyNo)
a<-rank[1]; b<-rank[2]; c<-rank[3]

temp2<-filter(temp, AssemblyNo %in% c(a,b,c))
temp3<-temp2 %>% group_by(MROC, AssemblyNo) %>% summarise(SeparateCount=n())
temp3$AssemblyNo<-factor(temp3$AssemblyNo, labels = c(a,b,c))
temp3<-temp3 %>% arrange(AssemblyNo, MROC)

library(tidyr)
temp4<-spread(temp3, MROC, SeparateCount)
len<-dim(temp4)[2]
Sum<-c(sum(temp4[1, 2:len], na.rm=TRUE), sum(temp4[2, 2:len], na.rm=TRUE), sum(temp4[3, 2:len], na.rm=TRUE))
Result2<-cbind(temp4, Sum)

FaultTemp<-gather(Fault, Pos, FaultFound, 1:5)
FaultTemp1<-as.data.frame(table(FaultTemp$FaultFound))
FaultTemp2<-FaultTemp1 %>% arrange(desc(Freq))

Result3<-FaultTemp2[2:6, 1:2]
colnames(Result3)[1]<-"Failures"

setwd("C:/Users/Weilun_Chiu/Desktop/QWE/MROC_Tab")
MROC_Tab_Data<-list(Result1, Result2, Result3)
write_xlsx(MROC_Tab_Data, "MROC_Tab.xlsx")

