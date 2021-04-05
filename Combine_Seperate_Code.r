library(readxl)
library(writexl)
library(dplyr)

setwd("C:/Users/Weilun_Chiu/Desktop/QWE/Qspeak")
filename<-list.files()
Qspeak<-do.call(rbind, lapply(filename, read.csv, header=TRUE, sep="|", na.strings = "NA"))
write_xlsx(Qspeak, path = "Temp.xlsx")

#This chunk of code separate data for HDD/SSD
HDD<-c("SEAGATE", "Sandisk", "HGST", "KIOXIA", "Seagate")
logic_HDD<-Qspeak$SupplierID %in% HDD
QHDD_PART1<-Qspeak[logic_HDD,]
QHDD_PART2<-Qspeak %>% filter(SupplierID %in% c("SAMSUNG", "SAMUSNG", "SAMSUNg"))
QHDD_PART2$len<-sapply(QHDD_PART2$EngBadgeNo, str_length)
QHDD_PART2<-QHDD_PART2 %>% filter(len<15)
QHDD<-rbind(QHDD_PART1, QHDD_PART2[,1:82])
HDD_Qspeak<-list(HDD_Qspeak=QHDD)
write_xlsx(HDD_Qspeak, "HDD_Qspeak_Report.xlsx")

#This chunk of code is for CPU/MEMORY DATA
HVP<-c("AMD", "HYNIX", "INTEL", "KINGSTON", "MICRON", "ELPIDA", "NANYA", "CN SAMSUNG", "INFINEON")
HVP<-c(HVP, "CR", "HINIX", "NICRON", "SLBV3", "MY", "HYNIx")
HVP<-c(HVP, "KINSTON", "INTEl", "IMTEL", "ELIPDA", "CAVIUM", " AMD", "Hynix", "Micron", "Intel")
logic<-Qspeak$SupplierID %in% HVP
QHVP_PART1<-Qspeak[logic, ]
QHVP_PART2<-Qspeak %>% filter(SupplierID %in% c("SAMSUNG", "SAMUSNG", "SAMSUNg"))
QHVP_PART2$len<-sapply(QHVP_PART2$EngBadgeNo, str_length)
QHVP_PART2<-QHVP_PART2 %>% filter(len>15)
QHVP<-rbind(QHVP_PART1, QHVP_PART2[, 1:82])

QRunway<-Qspeak[!logic,]
`%notin%` <- Negate(`%in%`)
QRunway<-QRunway %>% filter(SupplierID %notin% c("SEAGATE", "Sandisk", "HGST", "SAMSUNG", "SAMUSNG", "SAMSUNg", "Seagate", "KIOXIA"))
                          
test<-filter(QRunway, Scrapped == "N")
scrap<-filter(QRunway, Scrapped == "Y")

Qspeak_Runway_Data<-list(Qspeak_Data=test, Qspeak_Scrap=scrap)
HVP_Qspeak<-list(HVP_Qspeak=QHVP)

write_xlsx(Qspeak_Runway_Data, "Qspeak_Report.xlsx")
write_xlsx(HVP_Qspeak, "HVP_Qspeak_Report.xlsx")


###Perform extraction of HVP Scrap data
###Perform extraction of HVP Scrap data
###Perform extraction of HVP Scrap data
###Perform extraction of HVP Scrap data
QHVP1<-filter(QHVP, Warranty == "N")
QHVP2<-filter(QHVP1, Root.Cause %in% c("Functl", "DAMAGD", "FUNCTL", "MCO"))
QHVP3<-mutate(QHVP2, Location=ifelse(SupplierID %in% c("AMD", "INTEL", "INTEl", "IMTEL", "CAVIUM", "AMD"), "CPU", "Memory"))
QHVP4<-mutate(QHVP3, Problem = paste(Fault.Found.1, Fault.Code.3, sep=""))

QHVP_SCRAP<-select(QHVP4, Date.Received, Date.Shipped, Spare.PartNo, AssemblyNo, SerialNo, Problem, Location, Warranty, Fault.TagNo, MROC, Comp.Date.Code.1, SupplierID)
QHVP_SCRAP$Date.Received<-as.Date(QHVP_SCRAP$Date.Received,format="%d-%B-%y")

min(QHVP_SCRAP$Date.Received)
max(QHVP_SCRAP$Date.Received)


#################################################################################
#################################################################################
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
QHVP_SCRAP<-QHVP_SCRAP %>% select(Date.Reported, RMA, AssemblyNo, SerialNo, Problem, Location, Warranty, Fault.TagNo, MROC, Date.Code, SupplierID)

HVP_Scrap<-list(HVP_SCRAP=QHVP_SCRAP)
write_xlsx(HVP_Scrap, "HVP_SCRAP_FROM_Qspeak.xlsx")
