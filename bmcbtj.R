#
#部门人员工资成本统计
#
setwd("~/R/rwd")
dbcon <- read.csv("db.csv",header = T,stringsAsFactors = F)

setwd("~/R/rwd/lwf")
if(require(RMySQL)&require(xlsx)){
  conn <- dbConnect(RMySQL::MySQL(), 
                   host = dbcon$host,
                   user = dbcon$user, 
                   password = dbcon$password,
                   db=dbcon$db)

#定义查询的月度，每次查询之前更改
d<-"201607"

str1 <- "SELECT `tb_qt`.`GH` AS `GH`, `tb_qt`.`XM` AS `XM`, `tb_qt`.`QH` AS `QH`, `tb_qt`.`B1` AS `B1`, `tb_qt`.`B2` AS `B2`, `tb_qt`.`B3` AS `B3`, `tb_gz`.`应发合计` AS `应发合计`, `tb_gz`.`部门名称` AS `部门名称` FROM tb_qt LEFT JOIN `tb_gz` ON `tb_qt`.`GH` = `tb_gz`.`职员代码` AND `tb_gz`.`FFY` ='"
str2 <- "' WHERE `tb_qt`.`FFY` ='"
str3 <- "' AND tb_qt.LB LIKE '%劳务费%' AND `tb_qt`.`ZF` LIKE '总%' AND CAST(`tb_qt`.`QH` AS INT) < 96 AND CAST(`tb_qt`.`QH` AS INT) <> 18 AND CAST(`tb_qt`.`QH` AS INT) <> 20 AND tb_qt.GH NOT LIKE '9%' AND tb_qt.GH NOT LIKE '6%' AND tb_qt.GH NOT LIKE '7%' ORDER BY  `tb_qt`.`GH`,CAST(`tb_qt`.`QH` AS INT)"
stmt1 <- paste(str1,d,str2,d,str3,sep="");
#stmt1

dbGetQuery(conn,"SET NAMES UTF8")
lwf<-dbGetQuery(conn,stmt1)
str(dbDisconnect(conn))

}
#str(lwf)


#write2excel
system.time(write.xlsx2(lwf,paste('bmgz',d,'.xlsx',sep=''),sheetName='bmgz',row.names=FALSE))
