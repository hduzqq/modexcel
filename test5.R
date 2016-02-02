library(xlsx)

library(dplyr)

rm(list=ls())
path <- "F:/study/R/r4/chuangde1"
path_after <- paste(path,"_after",sep="")
dir.create(path_after)
fname <- dir(path)
fname
for(num in 1:length(fname)){
#  num <- 1
  cat(num,"-",fname[num],"-",sep="")
  setwd(path)
  fn <- paste(path,"/",fname[num],sep="")
  y <- read.xlsx2(fn,1,startRow = 5,colIndex=c(1:2),
                  header =F)
  x <- y[y$X2 != "",]
  x.str <- strsplit(as.character(x$X1),"")
  x.str.new <- data.frame(Xi=matrix(NA,length(x.str),1))
  for(i in 1:length(x.str)){
   # i <- 1
    x.istr <- unlist(x.str[i])
    
    x.str.new[i,1] <- ifelse(length(x.istr) == 7,
                             paste(substr(as.character(x$X1)[i],1,4),x.istr[6],x.istr[7],sep=""),
                             ifelse(length(x.istr) == 9 ,
                                    paste(substr(as.character(x$X1)[i],1,6),x.istr[8],x.istr[9],sep=""),
                                    as.character(x$X1)[i]))
    
  }
  
  data0 <- cbind(x.str.new,x[,2])
  data1 <- rbind(NA,NA,NA,NA,NA,data0)
  dt <-  read.xlsx(fn,1,startRow = 5,
                    endRow = dim(data0)[1]+4,
                   colIndex = c(7,8),
                   as.data.frame = T,
                   encoding = "UTF-8",
                   keepFormulas = F,
                   header =F)
  
  dt1 <- matrix(NA,dim(dt)[1]+5,2)
  for(i in  1:dim(dt)[1]){
    #i <- 1
    if (dt$X7[i] == "½è"){
      dt1[i+5,1]  <- dt$X8[i]
    }else{ if(dt$X7[i] == "´û"){ dt1[i+5,2] <- dt$X8[i]} }
  }
  
  
  data2 <- cbind(data1,NA,NA,NA,NA,dt1)
  setwd(path_after)
  savename <- paste(strsplit(fname[num],".xls")[[1]][1],".xls",sep="")
  write.xlsx(data2,savename,col.names=F, row.names=F,showNA=F)
}


