library(readxl)
library(writexl)
library(dplyr)
library(stringr)

setwd("C:/Users/Weilun_Chiu/Desktop/QWE/Qspeak")
filename<-list.files()
len<-length(filename)
filename<-filename[2:len]

Qspeak<-do.call(rbind, lapply(filename, read.csv, header=TRUE, sep="|", na.strings = "NA"))

### If can't combine check
### a<-lapply(filename, read.csv, header=TRUE, sep="|", na.strings = "NA")
### lapply(a, dim) and sapply(a, colnames)

Index_KQI<-ifelse(Qspeak$Test.Stage.Failed=="IQC", TRUE, FALSE)
KQI<-Qspeak[Index_KQI, ]
Qspeak<-Qspeak[!Index_KQI,]

write_xlsx(KQI, "KQI_Qspeak.xlsx")

HDD<-Qspeak %>% filter(str_detect(TrackingCodeIn, "3P") | SupplierID %in% c("HGST", "SEAGATE", "SANDISK"))
HDD_Qspeak<-list(HDD_Qspeak=HDD)
write_xlsx(HDD_Qspeak, "HDD_Qspeak.xlsx")

HDD_logic<-Qspeak$SerialNo %in% HDD$SerialNo
Qspeak<-Qspeak[!HDD_logic,]

HVP<-c("AMD", "HYNIX", "INTEL", "KINGSTON", "MICRON", "SAMSUNG", "ELPIDA", "NANYA", "CN SAMSUNG", "INFINEON")
HVP<-c(HVP, "CR", "HINIX", "NICRON", "SAMUSNG", "SLBV3", "MY", "KIOXIA", "HYMIX", "KINGTON")
HVP<-c(HVP, "KINSTON", "INTEl", "IMTEL", "ELIPDA", "CAVIUM", " AMD", "INTEN")

logic<-Qspeak$SupplierID %in% HVP
QRunway<-Qspeak[!logic, ]
QHVP<-Qspeak[logic, ]

test<-filter(QRunway, Scrapped == "N")
scrap<-filter(QRunway, Scrapped == "Y")

Qspeak_Runway_Data<-list(Qspeak_Data=test, Qspeak_Scrap=scrap)
HVP_Qspeak<-list(HVP_Qspeak=QHVP)

write_xlsx(Qspeak_Runway_Data, "Qspeak_Report.xlsx")
write_xlsx(HVP_Qspeak, "HVP_Qspeak_Report.xlsx")

table(test$SupplierID)
length_hvp_sn<-sapply(QHVP$SerialNo, str_length)
sum(length_hvp_sn<13)
which(length_hvp_sn<13)

###Perform extraction of HVP Scrap data
###Perform extraction of HVP Scrap data
###Perform extraction of HVP Scrap data
###Perform extraction of HVP Scrap data

QHVP<-read_xlsx("HVP_Qspeak_Report.xlsx")
QHVP1<-filter(QHVP, Warranty == "N")
QHVP2<-filter(QHVP1, Root.Cause %in% c("Functl", "DAMAGD", "FUNCTL", "MCO"))
QHVP3<-mutate(QHVP2, Location=ifelse(SupplierID %in% c("AMD", "INTEL", "INTEl", "IMTEL", "CAVIUM", "AMD"), "CPU", "Memory"))
QHVP4<-mutate(QHVP3, Problem = paste(Fault.Found.1, Fault.Code.3, sep=""))

QHVP_SCRAP<-select(QHVP4, Date.Received, Date.Shipped, Spare.PartNo, AssemblyNo, SerialNo, Problem, Location, Warranty, Fault.TagNo, MROC, Comp.Date.Code.1)
QHVP_SCRAP$Date.Received<-as.Date(QHVP_SCRAP$Date.Received,format="%d-%B-%y")

min(QHVP_SCRAP$Date.Received)
max(QHVP_SCRAP$Date.Received)


#################################################################################
#################################################################################
##Need to donwload VI report based on the min and max Date.Received, and then save it as VI.xlsx in the same folder
VI<-read_xlsx("VI.xlsx")
SN<-QHVP_SCRAP$SerialNo
Len<-length(QHVP_SCRAP$SerialNo)
pos<-rep(0, length=Len)

for(i in 1:Len){
        pos[i]<-VI$`RMA#`[which(SN[i]==VI$S.N.)]
}

QHVP_SCRAP<-cbind(QHVP_SCRAP, RMA=pos)
QHVP_SCRAP<-select(QHVP_SCRAP, -(Date.Received))
QHVP_SCRAP<-rename(QHVP_SCRAP, Date.Reported=Date.Shipped, Date.Code=Comp.Date.Code.1)
QHVP_SCRAP<-QHVP_SCRAP %>% select(Date.Reported, RMA, AssemblyNo, SerialNo, Problem, Location, Warranty, Fault.TagNo, MROC, Date.Code)

HVP_Scrap<-list(HVP_SCRAP=QHVP_SCRAP)
write_xlsx(HVP_Scrap, "HVP_SCRAP_FROM_Qspeak.xlsx")
