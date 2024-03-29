V_Tax: Make it easier
====================

用于V创税务局的简易程序
====================

### 前置的python模块：
* openpyxl
* configparser
* os
* glob

### 提供的函数：
* read_write_sheet(readworksheet1,readworksheet2) 两个参数分别为利润表的路径和纳税申报表的路径，功能为读取利润表中营业收入、营业成本、销售费用、管理费用、财务费用、投资收益、营业外收入、营业外支出、企业增值税、企业所得税、工资及五险一金，录入**财务数据汇总表**中
* exchange(dir)参数为目录，功能为修改该目录下的所有xls文件为xlsx文件

### 主要功能：

* 依据配置文件（config.ini）生成某期各企业文件夹，用以存放企业上交的表格
* 将各企业从柠檬云导出的.xls格式表格转换成openpyxl模块支持的.xlsx格式
* 重命名各企业上交的表格
* 从**利润表、纳税申报表**中提取数据录入**财务数据汇总表**中

### 各项文件介绍:
#### 配置文件 config.ini
```ini
[期数设置]
time = 第一期

[初始化]
initialized = no
第一期各企业数据 = False
第二期各企业数据 = False
第三期各企业数据 = False
第四期各企业数据 = False
第五期各企业数据 = False
```

配置文件中的[期数设置]节，中time的值，应用汉字分别填列：第一期、第二期、第三期、第四期、第五期 否则会出问题（程序中用if判断）
程序也会根据这个值录入汇总表的不同工作簿中

[初始化]节中的initialized 值为yes或者no

yes则会跳过 读取“初始化企业编号” 生成字典dict_players环节

下面的第X期各企业数据若为True，则代表对应期数的各企业文件夹已生成，程序会提示 检测到'+time+'企业汇总文件夹已生成，跳过此步骤。'

#### 企业编号文件 初始化企业编号.xlsx
该文件需要在最开始填写，用于在整个过程中在表格中搜索识别公司名称对应的行数等
填写规范为：
A列 机构类型+两位数编号 如 物流01 供应02 贸易04 制造05等
B列 公司全称（**需要能和企业上交的报表中编制单位对得上**）

#### Demo文件夹，存放了2022秋季第三批 第五期的真实数据和相应的配置文件

#### 其他可能会出现的种种bug/错误 之后再细说

![error1](http://lychee.alpacayyy.top:3801/uploads/big/aaa8a324efbad6ce5ee24c1c9b578ba3.png)

初始化表中的企业名后面有个空格

![error2](http://lychee.alpacayyy.top:3801/uploads/big/228816e4bb0f1074c562ae44ba200413.png)

企业上交的报表中没有 “纳税”关键字 程序没有正确为excel改名








