library(readxl)
library(tidyverse)
library(xlsx)
library(rJava)
library(xlsxjars)
library(readxl)
library(miscTools)
library(writexl)
library(base)
library(dplyr)
folder<-setwd("E:/UnderJonior/Rexercises/excel练习/2021湖南省统计年鉴")
#构建相同文件路径下的列表
#通过已经构建的列表导入数据

#将文件名导入列表
file <- list.files(pattern=".xls") 
length(file)
#将文件和文件位置结合得到文件全命名
filePath <- sapply(file, function(x){ 
  paste(folder,x,sep='\\')})
#批量导入文件并声称列表
data <- lapply(filePath, function(x){
  read_xls(x,sheet=1)})
#定义筛选函数，筛选出行为指定条件的元祖
name=function(x){
  x%>% as_tibble() %>% 
    filter(stringr::str_detect(市县名称, '芙蓉区|天心区|岳麓区|开福区|雨花区|望城县|长沙县|宁乡县|浏阳市|荷塘区|芦淞区|石峰区|天元区|株洲县|  攸  县|茶陵县|炎陵县|醴陵市|雨湖区|岳塘区|湘潭县|湘乡市|韶山市|珠晖区|雁峰区|石鼓区|蒸湘区|南岳区|衡阳县|衡南县|衡山县|衡东县|祁东县|耒阳市|常宁市|北湖区|苏仙区|桂阳县|宜章县|永兴县|嘉禾县|临武县|汝城县|桂东县|安仁县|资兴市|零陵区|冷水滩区|祁阳县|东安县|双牌县|江永县|宁远县|蓝山县|新田县|江华县|娄星区|冷水江市|双峰县|涟源市|邵东县|新邵县|邵阳县|新宁县|湘阴县|汨罗市|道  县|
                               '))
}
#定义函数，保留前面的标题栏
head=function(x){
  x[c(1:11),]
}
#获取list中的文件数目
v<-1:length(data)
#将第一列市县名称属性命名为市县名称
for (i in v)
{
  colnames(data[i][[1]])[1]<-'市县名称'
}

#循环对原始数据进行筛选操作
data1<- lapply(data, name)  
#循环对原始数据进行截取标题栏的操作
data2<- lapply(data, head) 
datanew<-list()
result=datanew[1][[1]]
for(i in v){
data2[i][[1]][is.na(data2[i][[1]])] <- "0"
}
signal<-data2

for(i in v)
{
  signal_i<-signal[i][[1]]
  l1=ncol(signal[i][[1]])
  lr=nrow(signal[i][[1]])
  #选出前面标题行数的最大值
  for(l2 in (1:11))
  {
    signal_i_1=signal_i[l2,1]
    if(signal_i_1=="芙蓉区"){
      break;
    }
    
  }
  #由行开始
  for (n in (1:l1))
  {
    #将每一列的字符串连接
    for(j in (l2:1))
    {
      #将分散在各个行的标题连接在一起
      coherent=paste(signal1=signal[i][[1]][1,n],signal1=signal[i][[1]][j,n],sep='')
      #同个正则表达式除去多余数字
      coherent=gsub("[[^0-9]",'',coherent)
      #将signal中的数据框第一行改为修改后的标题
      signal1=(signal[i][[1]][1,n]=coherent)
    }
  }
  signal[i][[1]]= signal[i][[1]][1,]
  datas1=do.call(rbind,data1[i])
  datas2=do.call(rbind,signal[i])
  #将标题和数据连接在一起
  data3=rbind(datas2,datas1)%>%distinct
  #将合并好的数据框加入到列表中，并且通过Paste0()进行重命名
  datanew[[paste0('180',i)]]<-data3
}
#设置输出路径
folder<-setwd("E:/UnderJonior/Rexercises/excel练习/2021湖南省统计年鉴result")
for(i in v){
  write.xlsx(datanew[i][[1]],file=paste0("180", "", i,"",".xls"))
}

