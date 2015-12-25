# FASTTIP_CATIA_METHOD
操作自动钻铆编程所用的TVA文件及PPR文档
Operate TVA and PPR Document for Auto-rivet programming
##实现自动钻铆生产中所需的常用功能：
* 从工程数模中进行TVA文件提取（Creation）；
* 根据数模对TVA紧固件类型进行检查及修复（Check Fasteners）
* TVA文件树结构的检查并最终错误点位（Check、Revise TVA）；
* 紧固件数量、加工类型统计(Check)
* TVA文件点位根据加工类型进行批量移动（Revise ProcessType）；
* 两个TVA文件点位比较（Compare）
* 将点位追加至TVA（TVA间点位复制）（Append）
* 修复常见错误（重复点，点位不在蒙皮外表面）（Fix）
* 将点位信息同步进数据库（Update the database）
* 显示，隐藏特定的加工类型，用于制作图纸（Show/Hide）
