temp<-test %>% select(AssemblyNo, Fault.Code.1)

sum_record<-data.frame(table(temp$AssemblyNo))
sum_record<-sum_record %>% arrange(desc(Freq))
sum_qty<-sum(sum_record$Freq)

nff<-temp %>% filter(Fault.Code.1 =="")
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

