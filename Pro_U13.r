filename<-list.files()
Pro<-do.call(rbind, lapply(filename, read_xls))
Pro1<-Pro %>% filter(Station=="BGA Station") %>% select(P.N., "Error Reference")

colnames(Pro1)<-c("PN", "Loc")

#622217
Hen<-Pro1 %>% filter(str_detect(PN, "622217"))
Hen<-Hen %>% filter(str_detect(Loc, "13"))

#013548
Arg<-Pro1 %>% filterstr_detect(PN, "013548")
Arg<-Arg %>% filter(str_detect(Loc, "13"))

#635678
Sai<-Pro1 %>% filter(str_detect(PN, "635678"))
Sai<-Sai %>% filter(str_detect(Loc, "13"))

#635678
Vin<-Pro1 %>% filter(str_detect(PN, "664924"))
Vin<-Vin %>% filter(str_detect(Loc, "13"))

as.data.frame(table(Hen$Loc))
as.data.frame(table(Arg$Loc))
as.data.frame(table(Sai$Loc))
as.data.frame(table(Vin$Loc))
