# Application Name
  邮件发布系统

# Description
  此系统主要功能是通过Email方式发送如工资单等各类单据，并可定制相关报表。

# Install & Deployment Guide
  如果是重新部署，则需要制定安装包；

如果是升级，请注意：
  
对3.1.5之前的版本，建议重新部署。
  
对3.1.5以后的版本进行更新，需要处理：
  1. 备份原有的数据表的内容；
  2. 拷贝下列文件列表：
     - salary.exe
     - Init.dat
     - 更新说明.txt
  3. 对系统进行初始化；
  4. 恢复原有数据
