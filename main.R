################################################################
# 操作步骤
# 1、解压所有文件
# 2、提取所有.xls
# 3、设置当前文件路径，修改以下代码
setwd("C:/Users/admin/Desktop/wenjian")
# 4、所需包加载，未安装过先install.packages()
library(openxlsx)
library(readxl)
# 5、运行所有代码
# 6、结果在当前目录下Results文件夹下
################################################################

# 清除环境
rm(list = ls())

# 文件名向量
file_names <- list.files(path = getwd(), pattern = "\\.xls$")  

# 创建新目录
dir.create('Results') 

# 循环文件
for (FILE in file_names) {
  # 导入原始数据
  source <- read_excel(FILE,col_names = FALSE)
  
  # 删除Sample Name及以上的行
  row <- which(source[,1]=='Sample Name')
  source <- source[-c(1:row),]
  rm(row)
  
  # 定义一个随机SampleName
  set.seed(123)
  # 生成字符向量
  characters <- c(letters, LETTERS)  # 包含所有小写和大写字母
  # 生成无规则不重复的长度为4的字符串向量
  RandSN <- character(20)
  for (i in 1:20) {
    RandSN[i] <- paste0(sample(characters, 4, replace = FALSE), collapse = "")
  }
  rm(i,characters)
  
  
  # 填充SampleName
  count <- 0
  for (i in 1:nrow(source)) {
    if (count %% 6==0){
      if (is.na(source$...1[i])){
        # 从向量中不放回抽取一个元素
        sampled_index <- sample(length(RandSN), 1)
        source$...1[i] <- RandSN[sampled_index]
        RandSN <- RandSN[-sampled_index]
        count <- count+1
      }else{
        count <- count+1
      }
    }else{
      if (is.na(source$...1[i])){
        source$...1[i] <- source$...1[i-1]
        count <- count+1
      }else{
        if (source$...1[i] != source$...1[i-1]){
          count <- 1
        }else{
          count <- count+1
        }
      }
    }
  }
  rm(count,i,RandSN)
  
  # 按第一列，即SampleName重排
  source <- source[order(source$...1),]
  
  
  # 定义out数据集
  SampleName <- c()
  CT1_pre <- c()
  CT2_pre <- c()
  CT3_pre <- c()
  CTmean_pre <- c()
  CTsd_pre <- c()
  CT1_pos <- c()
  CT2_pos <- c()
  CT3_pos <- c()
  CTmean_pos <- c()
  CTsd_pos <- c()
  Plate <- c()
  for (i in unique(source$...1)){
    sn <- source[source$...1 == i,]
    SampleName <- c(SampleName,i)
    
    # 前引物
    yinwu <- sn[tolower(sn$...2) == tolower("ND1")|tolower(sn$...2) == tolower("TEL"),]
    if (length(yinwu$...3)==0){
      CT1_pre <- c(CT1_pre,NA)
      CT2_pre <- c(CT2_pre,NA)
      CT3_pre <- c(CT3_pre,NA)
    }else if (length(yinwu$...3)==1){
      yinwu$...3 <- na.omit(yinwu$...3)
      CT1_pre <- c(CT1_pre,yinwu$...3[1])
      CT2_pre <- c(CT2_pre,NA)
      CT3_pre <- c(CT3_pre,NA)
    }else if (length(yinwu$...3)==2){
      yinwu$...3 <- na.omit(yinwu$...3)
      CT1_pre <- c(CT1_pre,yinwu$...3[1])
      CT2_pre <- c(CT2_pre,yinwu$...3[2])
      CT3_pre <- c(CT3_pre,NA)
    }else if (length(yinwu$...3)==3){
      CT1_pre <- c(CT1_pre,yinwu$...3[1])
      CT2_pre <- c(CT2_pre,yinwu$...3[2])
      CT3_pre <- c(CT3_pre,yinwu$...3[3])
    }else{
      CT1_pre <- c(CT1_pre,yinwu$...3[1])
      CT2_pre <- c(CT2_pre,yinwu$...3[2])
      CT3_pre <- c(CT3_pre,yinwu$...3[3])
    }
    my_vector <- yinwu$...4
    if (length(my_vector[grep("^-?[0-9]*\\.?[0-9]+$", my_vector)])==0){
      CTmean_pre <- c(CTmean_pre,NA)
    }else {
      CTmean_pre <- c(CTmean_pre,sample(my_vector[grep("^-?[0-9]*\\.?[0-9]+$", my_vector)],1))
    }
    my_vector <- yinwu$...5
    if (length(my_vector[grep("^-?[0-9]*\\.?[0-9]+$", my_vector)])==0){
      CTsd_pre <- c(CTsd_pre,NA)
    }else {
      CTsd_pre <- c(CTsd_pre,sample(my_vector[grep("^-?[0-9]*\\.?[0-9]+$", my_vector)],1))
    }
    
    
    
    # 后引物
    yinwu <- sn[tolower(sn$...2) == tolower("ACTB")|tolower(sn$...2) == tolower("ACTI"),]
    if (length(yinwu$...3)==0){
      CT1_pos <- c(CT1_pos,NA)
      CT2_pos <- c(CT2_pos,NA)
      CT3_pos <- c(CT3_pos,NA)
    }else if (length(yinwu$...3)==1){
      yinwu$...3 <- na.omit(yinwu$...3)
      CT1_pos <- c(CT1_pos,yinwu$...3[1])
      CT2_pos <- c(CT2_pos,NA)
      CT3_pos <- c(CT3_pos,NA)
    }else if (length(yinwu$...3)==2){
      yinwu$...3 <- na.omit(yinwu$...3)
      CT1_pos <- c(CT1_pos,yinwu$...3[1])
      CT2_pos <- c(CT2_pos,yinwu$...3[2])
      CT3_pos <- c(CT3_pos,NA)
    }else if (length(yinwu$...3)==3){
      CT1_pos <- c(CT1_pos,yinwu$...3[1])
      CT2_pos <- c(CT2_pos,yinwu$...3[2])
      CT3_pos <- c(CT3_pos,yinwu$...3[3])
    }else{
      CT1_pos <- c(CT1_pos,yinwu$...3[1])
      CT2_pos <- c(CT2_pos,yinwu$...3[2])
      CT3_pos <- c(CT3_pos,yinwu$...3[3])
    }
    my_vector <- yinwu$...4
    if (length(my_vector[grep("^-?[0-9]*\\.?[0-9]+$", my_vector)])==0){
      CTmean_pos <- c(CTmean_pos,NA)
    }else {
      CTmean_pos <- c(CTmean_pos,sample(my_vector[grep("^-?[0-9]*\\.?[0-9]+$", my_vector)],1))
    }
    my_vector <- yinwu$...5
    if (length(my_vector[grep("^-?[0-9]*\\.?[0-9]+$", my_vector)])==0){
      CTsd_pos <- c(CTsd_pos,NA)
    }else {
      CTsd_pos <- c(CTsd_pos,sample(my_vector[grep("^-?[0-9]*\\.?[0-9]+$", my_vector)],1))
    }
    
    #Plate <- c(Plate,sub("\\.[^.]+$", "", FILE))
    Plate <- c(Plate,"2023-08-15 C3307VJO")
    
  }
  rm(i,my_vector,sn,yinwu)
  
  # 输出数据框
  out <- data.frame(SampleName,CT1_pre,CT2_pre,CT3_pre,CTmean_pre,CTsd_pre,
                    CT1_pos,CT2_pos,CT3_pos,CTmean_pos,CTsd_pos,Plate)
  
  # 输出
  if (tolower('nd1') %in% tolower(unique(source$...2))){
    names(out) <- c("SampleName","CT1_pre_ND1","CT2_pre_ND1","CT3_pre_ND1","CTmean_pre_ND1","CTsd_preV_ND1", 
                    "CT1_pos_ACTB","CT2_pos_ACTB","CT3_pos_ACTB","CTmean_pos_ACTB","CTsd_pos_ACTB","Plate")
    # 创建一个工作簿
    wb <- createWorkbook()
    # 在工作簿中添加一个工作表
    addWorksheet(wb, "Mitochondria")
    # 将数据框中的数据写入工作表
    writeData(wb, sheet = "Mitochondria", x = out)
    # 保存工作簿为Excel文件
    saveWorkbook(wb, paste(sep = "","./Results/",sub("\\.[^.]+$", "", FILE),'.xlsx'), overwrite = TRUE)
    
  }else{
    names(out) <- c("SampleName","CT1_pre_TEL","CT2_pre_TEL","CT3_pre_TEL","CTmean_pre_TEL","CTsd_preV_TEL",
                    "CT1_pos_ACTI","CT2_pos_ACTI","CT3_pos_ACTI","CTmean_pos_ACTI","CTsd_pos_ACTI","Plate")
    # 创建一个工作簿
    wb <- createWorkbook()
    # 在工作簿中添加一个工作表
    addWorksheet(wb, "Telomere")
    # 将数据框中的数据写入工作表
    writeData(wb, sheet = "Telomere", x = out)
    # 保存工作簿为Excel文件
    saveWorkbook(wb, paste(sep = "","./Results/",sub("\\.[^.]+$", "", FILE),'.xlsx'), overwrite = TRUE)
  }
}
# Coded by Quanhong Liu 2023.8.27
