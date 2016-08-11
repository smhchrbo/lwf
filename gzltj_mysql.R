#
#部门工作量、相对价值点统计
#
setwd("~/R/rwd")
dbcon <- read.csv("db.csv",header = T,stringsAsFactors = F)

setwd("~/R/rwd/lwf")
if(require(RMySQL)&require(sqldf)&require(xlsx)&require(reshape)){
  conn <- dbConnect(RMySQL::MySQL(), 
                    host = dbcon$host,
                    user = dbcon$user, 
                    password = dbcon$password,
                    db=dbcon$db)

#定义查询的'开始日期'
d<-'2016-7-1'

stmt.z1<-paste("select * from v_bfgzl_zb where s like '%",d,"%'",sep="") #总部工作量
stmt.f1<-paste("select * from v_bfgzl_fb where s like '%",d,"%'",sep="") #分部工作量
stmt.mz1<-paste("select * from v_mzgzl_z where s like '%",d,"%'",sep="") #总部门诊
#stmt
dbGetQuery(conn,"SET NAMES UTF8")
qz<-dbGetQuery(conn,stmt.z1) #总部工作量
qf<-dbGetQuery(conn,stmt.f1) #分部工作量
qmz<-dbGetQuery(conn,stmt.mz1) #总部门诊
str(dbDisconnect(conn))
#避免sqldf与RMySQL冲突，此处先卸载MySQL包
detach("package:RMySQL", unload=TRUE)
}
#str(q1)

##
## 计算总部工作量权重
##
if(1){
q1<-qz
#1314病区特需病床权重处理
q1[which(q1$lbh==1&q1$w==20),]$w<-93.6
q1$w1w=q1$w1*q1$w
q1$w2w=q1$w2*q1$w
q1$w3w=q1$w3*q1$w
q1$w4w=q1$w4*q1$w
q1$w5w=q1$w5*q1$w
q1$w6w=q1$w6*q1$w
q1$w7w=q1$w7*q1$w
q1$w8w=q1$w8*q1$w
q1$w9w=q1$w9*q1$w
q1$w10w=q1$w10*q1$w
q1$w11w=q1$w11*q1$w
q1$w12w=q1$w12*q1$w
q1$w13w=q1$w13*q1$w
q1$w14w=q1$w14*q1$w
q1$w15w=q1$w15*q1$w
q1$w16w=q1$w16*q1$w
q1$w17w=q1$w17*q1$w


#write.csv(q1,paste('gzl',d,'.csv',sep=''))
system.time(write.xlsx2(q1,paste('gzl',d,'.xlsx',sep=''),sheetName='gzl',row.names=FALSE))
q2<-sqldf::sqldf('select lbh,[类别名],sum(w1w),sum(w2w),sum(w3w),sum(w4w),sum(w5w),sum(w6w),sum(w7w),sum(w8w),sum(w9w),sum(w10w),sum(w11w),sum(w12w),sum(w13w),sum(w14w),sum(w15w),sum(w16w),sum(w17w) from q1 group by lbh,[类别名];')
#write.csv(q2,paste('q2',d,'.csv',sep=''))
system.time(write.xlsx2(q2,paste('gzl',d,'.xlsx',sep=''),sheetName='total',append=TRUE,row.names=FALSE))
}


##
##计算脱氧核糖核酸工作量
##
##总部dna
qz.dna<-sqldf("select * from qz where xmmc like '%脱氧核糖核酸%'")
qz.dna<-qz.dna[,-c(1:10)]

#分部dna
qf.dna<-sqldf("select * from qf where xmmc like '%脱氧核糖核酸%'")
qf.dna<-qf.dna[,-c(1:10)]

#总部门诊dna
qmz.dna<-sqldf("select s,部门名称,数量 from qmz where xmmc like '%脱氧核糖核酸%'")
names(qmz.dna)<-c("s","variable","value")
##病房dna数据合并
dna<-cbind(qz.dna,qf.dna,row.names="s")
##病房数据处理及添加门诊行
md.dna<-melt(dna,id="s")
md.dna$variable<-as.vector(md.dna$variable)
md.dna<-rbind(md.dna,qmz.dna)
##删除0值
md.dna<-md.dna[md.dna$value!=0,]
#追加到xlsx表
system.time(write.xlsx2(md.dna,paste('gzl',d,'.xlsx',sep=''),sheetName='dna',append=TRUE,row.names=FALSE))
