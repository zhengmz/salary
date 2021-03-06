## 一、使用前的须知：

1. 不要擅且改动程序安装目录下任何文件。
2. 本软件可作为发布邮件之用，也可对数据进行保存或处理，如报表
3. 本软件是单机版，不会泄露任何数据
4. 环境要求不高，除了VB运行库，还需要有ADO2.0以上版本(包含在安装包中)。

## 二、本软件使用方便，操作简单，步骤如下：

1. 维护员工信息，可通过手工增加，也可从Excel文件导入；
2. 维护服务配置，建议通过向导来对各种单据的格式在系统中进行设置；
3. 将单据的Excel文件导入到系统中，然后进行发布。

## 三、用户界面菜单：

```
    文件(F)
    ----关闭(C)    		关闭当前活动的子窗口
    ----关闭所有    		关闭所有的子窗口
    ----退出(X)     		退出系统
    操作(O)
    ----员工信息(E)		维护员工的基本信息：员工编码、姓名和邮箱地址
    ----服务配置(S)		维护服务配置主界面
    ----服务配置向导(W)		通过向导将工资单格式进行设置
    ----数据导入(L)		导入工资单数据，也可直接发布
    ----邮件发布(Ctrl+S)	查询工资单数据，可进行邮件发布，也可删除
    报表(R)
    ----报表配置(C)		配置报表格式
    ----报表生成(P)		可查询、生成、保持报表数据，并导出Excel文件
    工具(T)
    ----选项(O)			维护系统的一些选项，包括数据字典
    ----系统工具(S)
    --------自动修复		此工具可对影响系统运行的一些数据进行自动修复
    --------备份与恢复		此工具支持对系统数据的备份和恢复
    --------初始化		此工具是为系统第一次运行进行设置一些基本系统数据
    帮助(H)
    ----使用手册(M)
    ----常见问题(F)
    ----更新说明(U)
    ----关于(A)
```

## 四、导入文件格式：

1. 员工基本信息
   - Excel格式
   - 文件中要导入的工作表名称为sheet1
   - 第一行不导入，当作标题栏
   - 从第二行开始导入，第一列为员工编号，第二列为员工姓名，第三列为邮箱地址
   - 员工编号必须唯一
2. 工资发送表
   - Excel格式\CSV格式
   - 第一行不导入，当作为标题栏，且要与配置中的显示名称对应
   - 有效数据从第二行开始

## 五、员工基本信息的维护：

1. 从菜单“操作”->“员工信息”进入，完成对员工基本信息的维护
2. 界面分为上中下三部分，上面为数据导入部分，中间是数据显示和编辑区，下面是对记录进行修改，数据导入部分是对整张表的维护，可以导入或清除数据
3. 可以直接在中间部分对数据进行直接的修改，但必须保证编号的唯一性
4. 提供“导出”功能，是将软件中的员工基本信息导到一个Excel文件，方便修改
5. 发生变化时，请及时更新其相应的信息
6. 员工编号不要求连续，不限制其组成内容的格式
