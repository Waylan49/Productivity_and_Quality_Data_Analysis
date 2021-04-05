temp<-QHVP %>% select(SupplierID, AssemblyNo, NFF)
temp1<-mutate(temp, Type=ifelse(SupplierID %in% c("AMD", "INTEL", "INTEl", "IMTEL", "INTEN", "CAVIUM", "AMD"), "CPU", "Memory"))


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
