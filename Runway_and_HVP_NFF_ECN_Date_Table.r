setwd("C:/Users/Weilun_Chiu/Desktop/QWE/Qspeak")
test1<-read_xlsx("Qspeak_Report.xlsx", sheet = 1, na = "") ### Be aware of "NA" problem
temp<-test1 %>% select(AssemblyNo, Fault.Code.1)

sum_record<-data.frame(table(temp$AssemblyNo))
sum_record<-sum_record %>% arrange(desc(Freq))
sum_qty<-sum(sum_record$Freq)

nff<-temp %>% filter(is.na(Fault.Code.1))
nff_record<-data.frame(table(nff$AssemblyNo)) %>% arrange(desc(Freq))
names(nff_record)<-c("Assy", "Qty.")
nffpart<-as.character(nff_record$Assy)

len1<-length(nffpart)
total_qty<-rep(0, len1)
for(i in 1:len1){
  ind<-which(sum_record$Var1==nffpart[i])
  total_qty[i]<-sum_record$Freq[ind]
}

nff_record<-cbind(nff_record, total_qty)
nff_record$percent<-round(nff_record$Qty./nff_record$total_qty, 3)
nff_qty<-sum(nff_record$Qty.)

nff_rate<-as.data.frame(round(t(data.frame(NFF=nff_qty, Sum=sum_qty, Percent=nff_qty/sum_qty)),4))

nff_top5<-nff_record[1:5, 1:2]
others<-c("Others",nff_qty-sum(nff_top5$Qty.))
Total_NFF<-c("Total NFF", nff_qty)
nff_top5$Assy<-as.character(nff_top5$Assy)
nff_top5_1<-rbind(nff_top5, others, Total_NFF)

nff_record<-nff_record %>% arrange(desc(percent)) %>% select(Assy, total_qty, Qty., percent)
top5_nff_percent<-nff_record[1:5, ]

nff_record1<-nff_record %>% arrange(desc(Qty.))
top5_nff_qty_percent<-nff_record1[1:5, ]

top20_nff_percent<-nff_record[1:20, ]


NFF_Final<-list(NFF_Rate=nff_rate, NFF_TOP5=nff_top5_1, TOP5_NFF_Percent=top5_nff_percent, 
                TOP5_NFF_Qty_Percent=top5_nff_qty_percent, TOP20_NFF_Percent=top20_nff_percent, Raw.NFF=nff_record)
write_xlsx(NFF_Final, "NFF_Data.xlsx")


ecn<-temp %>% filter(Fault.Code.1 =="101")
ecn_record<-data.frame(table(ecn$AssemblyNo)) %>% arrange(desc(Freq))
names(ecn_record)<-c("Assy", "Qty.")
ecnpart<-as.character(ecn_record$Assy)
len2<-length(ecnpart)
total_qty1<-rep(0, len1)
for(i in 1:len2){
  ind<-which(sum_record$Var1==ecnpart[i])
  total_qty1[i]<-sum_record$Freq[ind]
}

ecn_record<-cbind(ecn_record, total_qty1)
ecn_record$percent<-round(ecn_record$Qty./ecn_record$total_qty, 3)
ecn_record<-ecn_record %>% select(Assy, total_qty1, Qty., percent)

ecn_qty<-sum(ecn_record$Qty.)
ecn_rate<-as.data.frame(round(t(data.frame(ECN=ecn_qty, Sum=sum_qty, Percent=ecn_qty/sum_qty)),4))

ecn_top5<-ecn_record[1:5, c(1,3)]
others<-c("Others",ecn_qty-sum(ecn_top5$Qty.))
Total_ECN<-c("Total ECN", ecn_qty)
ecn_top5$Assy<-as.character(ecn_top5$Assy)
ecn_top5_1<-rbind(ecn_top5, others, Total_ECN)

ecn_record<-ecn_record %>% arrange(desc(percent))
top5_ecn_percent<-ecn_record[1:5, ]

ecn_record1<-ecn_record %>% arrange(desc(Qty.))
top5_ecn_qty_percent<-ecn_record1[1:5, ]

top20_ecn_percent<-ecn_record[1:20, ]


ECN_Final<-list(ECN_Rate=ecn_rate, ECN_TOP5=ecn_top5_1, TOP5_ECN_Percent=top5_ecn_percent, 
                TOP5_ECN_Qty_Percent=top5_ecn_qty_percent, TOP20_ECN_Percent=top20_ecn_percent, Raw.ECN=ecn_record)
write_xlsx(ECN_Final, "ECN_Data.xlsx")





#############################HVP Section
temp<-QHVP %>% select(SupplierID, AssemblyNo, NFF)
temp1<-mutate(temp, Type=ifelse(SupplierID %in% c("AMD", "INTEL", "INTEl", "IMTEL", "CAVIUM", "AMD"), "CPU", "Memory"))


mem<-temp1 %>% filter(Type=="Memory")
mem1<-mem[, c(2,3)]
cpu<-temp1 %>% filter(Type=="CPU")
cpu1<-cpu[, c(2,3)]

#########################################################Memory
mem_total<-data.frame(table(mem1$AssemblyNo))
MEM_TOTAL_QTY<-sum(mem_total$Freq)
mem_nff<-mem1 %>% filter(NFF=="Y")
mem_nff1<-data.frame(table(mem_nff$AssemblyNo))
mem_nff_part<-as.character(mem_nff1$Var1)

len1<-length(mem_nff_part)
total_qty<-rep(0, len1)
for(i in 1:len1){
  ind<-which(mem_total$Var1==mem_nff_part[i])
  total_qty[i]<-mem_total$Freq[ind]
}

mem_nff_2<-cbind(mem_nff1, total_qty)
mem_nff_3<-mem_nff_2 %>% select(Var1, total_qty, Freq)
names(mem_nff_3)<-c("Assy", "Total_Qty", "NFF")
mem_nff_3$Percent<-round(mem_nff_3$NFF/mem_nff_3$Total_Qty, 4)

MEM_NFF_QTY<-sum(mem_nff_3$NFF)

MEM_NFF_PERCENT<-round(MEM_NFF_QTY/MEM_TOTAL_QTY, 4)
mem_nff_rate<-as.data.frame(t(data.frame(NFF=MEM_NFF_QTY, Sum=MEM_TOTAL_QTY, Percent=MEM_NFF_PERCENT)))

mem_nff_top5<-mem_nff_3 %>% arrange(desc(NFF)) %>% select(Assy, NFF)
mem_nff_top5<-mem_nff_top5[1:5,]
Others<-c("Others", MEM_NFF_QTY-sum(mem_nff_top5$NFF))
Total<-c("Total", MEM_NFF_QTY)
mem_nff_top5$Assy<-as.character(mem_nff_top5$Assy)
mem_nff_top5<-rbind(mem_nff_top5, Others, Total)

mem_nff_3<-mem_nff_3 %>% arrange(desc(NFF)) %>% arrange(desc(Percent))
mem_top5_nff_percent<-mem_nff_3[1:5,]
mem_nff_4<-mem_nff_3 %>% arrange(desc(NFF))
mem_top5_nff_qty<-mem_nff_4[1:5, ]


MEM_NFF_FINAL<-list(MEM_NFF_RATE=mem_nff_rate, MEM_NFF_TOP5=mem_nff_top5, MEM_TOP5_NFF_PERCENT=mem_top5_nff_percent,
                    MEM_TOP5_NFF_QTY=mem_top5_nff_qty, Raw.MEM.NFF=mem_nff_3)
write_xlsx(MEM_NFF_FINAL, "MEM_NFF_Data.xlsx")


#########################################################CPU
cpu_total<-data.frame(table(cpu1$AssemblyNo))
CPU_TOTAL_QTY<-sum(cpu_total$Freq)
cpu_nff<-cpu1 %>% filter(NFF=="Y")
cpu_nff1<-data.frame(table(cpu_nff$AssemblyNo))
cpu_nff_part<-as.character(cpu_nff1$Var1)

len1<-length(cpu_nff_part)
total_qty1<-rep(0, len1)
for(i in 1:len1){
  ind<-which(cpu_total$Var1==cpu_nff_part[i])
  total_qty1[i]<-cpu_total$Freq[ind]
}

cpu_nff_2<-cbind(cpu_nff1, total_qty1)
cpu_nff_3<-cpu_nff_2 %>% select(Var1, total_qty1, Freq)
names(cpu_nff_3)<-c("Assy", "Total_Qty", "NFF")
cpu_nff_3$Percent<-round(cpu_nff_3$NFF/cpu_nff_3$Total_Qty, 4)

CPU_NFF_QTY<-sum(cpu_nff_3$NFF)
CPU_NFF_PERCENT<-round(CPU_NFF_QTY/CPU_TOTAL_QTY, 4)
cpu_nff_rate<-as.data.frame(t(data.frame(NFF=CPU_NFF_QTY, Sum=CPU_TOTAL_QTY, Percent=CPU_NFF_PERCENT)))

cpu_nff_top5<-cpu_nff_3 %>% arrange(desc(NFF)) %>% select(Assy, NFF)
cpu_nff_top5<-cpu_nff_top5[1:5,]
Others<-c("Others", CPU_NFF_QTY-sum(cpu_nff_top5$NFF))
Total<-c("Total", CPU_NFF_QTY)
cpu_nff_top5$Assy<-as.character(cpu_nff_top5$Assy)
cpu_nff_top5<-rbind(cpu_nff_top5, Others, Total)

cpu_nff_3<-cpu_nff_3 %>% arrange(desc(NFF)) %>% arrange(desc(Percent))
cpu_top5_nff_percent<-cpu_nff_3[1:5,]
cpu_nff_4<-cpu_nff_3 %>% arrange(desc(NFF))
cpu_top5_nff_qty<-cpu_nff_4[1:5, ]

CPU_NFF_FINAL<-list(CPU_NFF_RATE=cpu_nff_rate, CPU_NFF_TOP5=cpu_nff_top5, CPU_TOP5_NFF_PERCENT=cpu_top5_nff_percent,
                    CPU_TOP5_NFF_QTY=cpu_top5_nff_qty, Raw.CPU.NFF=cpu_nff_3)
write_xlsx(CPU_NFF_FINAL, "CPU_NFF_Data.xlsx")

