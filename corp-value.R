library(quantmod)
library(leaps)
library(Hmisc) #describe
library(psych) #describe
library(GPArotation)
library(pastecs) #stat.desc
library(corrgram) # for corralation analysis
library(gvlma)
library(car)
library(relaimpo)

library(xlsx)
library(RSQLite)
library(RMySQL)

library(ggplot2) #add for ggplot
library(reshape2)
library(dplyr)

##deal with CSV file
##1. delete the "," in number
##2. change the date to 
par(mfrow=c(1,1))
mycolor=rainbow(20)

ncorp=14
nvar=6
#resinall=data.frame(matrix(NA,ncorp*n,nvar))
#resinall$var=NA
#resinall$corp=NA
resinall=data.frame()


for (i in 1:14){
#for (i in 4){
  wholesheet=read.xlsx("Valuation-Typical International Companies.xlsx",i,stringsAsFactors=F)
  #str(wholesheet)
  #which(names(wholesheet)=="NA..10")
  used=wholesheet[c(3,4,5,6,7,9),c(1,14,19,24,29,30)]
  #print(names(used)[1])
  udf=as.data.frame(t(used[,2:ncol(used)]))
  udf=as.data.frame(apply(udf,2,as.numeric))
  names(udf)=c("Cap","Values","EBITDA","Revenue","NI","Nusers")
  print(names(used)[1])
  print(udf)
  
  if (F){    
    print(cor(udf$Cap,udf,use="na.or.complete"))
    print(cor(udf,use="na.or.complete"))
    print(cor(udf))
    print(cor(udf$Cap,udf))
  }
  restbl=(matrix(NA,ncol(udf),ncol(udf)))
  for (j in 1:ncol(udf)){
    for (k in j:ncol(udf)){
      restbl[k,j]=cor(udf[,j],udf[,k],use="na.or.complete")
    }
  }
  
  resdf=as.data.frame(restbl)
  names(resdf)=names(udf)
  row.names(resdf)=names(udf)
  print(resdf)
  #write.xlsx(file=paste(i,names(used)[1],".xlsx"),resdf,showNA=F)
  write.xlsx(file=paste(file="Enterprise Correlation Coefficient Aanlysis Result.xlsx"),
             x=resdf,showNA=F,sheetName=names(used)[1],append=T)
  
  
  
  
  for (ivar in 1:ncol(udf)){
    res.corp=c()
    for (j in 1:ncol(udf)){
      res.corp=c(res.corp,cor(udf[,ivar],udf[,j],use="na.or.complete"))        
    }
    #print(res.corp)
    
    resinall=rbind(resinall,data.frame(t(res.corp),var=names(udf)[ivar],corp=names(used)[1]))
    #resinall[i,]=data.frame(res.corp,var=names(udf)[ivar],corp=names(used)[1])
    #row.names(resinall)[i]=names(used)[1]
  }
  
    
}

names(resinall)=c("Cap","Values","EBITDA","Revenue","NI","Nusers","var","corp")
resinall=select(resinall,corp,var,Cap,Values, EBITDA,Revenue,NI,Nusers)

resinall=na.omit(resinall)
print(resinall)

for (i in 1:ncol(udf)){
  #write.xlsx(file=paste(i," ",names(udf)[i],"-correlation coefficient.xlsx",sep=""),filter(resinall,var==names(udf)[i]),showNA=F)
  write.xlsx(file=paste(file="Enterprise Correlation Coefficient Aanlysis Result.xlsx"),
             x=filter(resinall,var==names(udf)[i]),showNA=F,sheetName=names(udf)[i],append=T)
}
