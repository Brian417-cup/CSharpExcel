# C# Excel With SQL For Windows
C#连接Excel并执行SQL操作



# 输入

  ## 支持格式

- CSharpForExcel_Plus

  [下载地址](./bin/Release/CSharpForExcel_Plus.exe)

  ```shell
  Excel 2003 格式的文件 *.xls
  ```

  

- CSharpForExcel

  [下载地址](./bin/Release/CSharpForExcel.exe)

  **需要安装** [ACE 32位插件]()
  
  ```shell
  Excel 2003 格式的文件 *.xls
  Excel 2007 格式的文件 *.xlsx
  ```

  ## 表格限制
  
    支持多张表的操作，但是这多张表格应该在一个.xls文件中
  ## 单元格限制 
    没有合并单元格和非合并单元格并存的情况，第一行有列名 (其将作为最终SQL语句查询的列名)


# 支持操作

  ## 查询记录
    select SQL语句
  ## 插入记录
    intert SQL语句
  ## 更新记录
    update SQL语句
  ## 删除整张表
     drop SQL语句



# 输出(针对查询结果的文件保存)

  ## 支持格式
    Excel 2003 格式的文件(Windows系统的电脑上要安装Office2010 或更高版本 WPS可能不行) 或  .csv格式文件(推荐！，可以在不装有Office或WPS的电脑上进行导出)



# 样例测试
  /bin/Release/CSharpForExcel.exe  直接打开运行即可  
  另：配备一个"/bin/Release/测试数据.xls" 来进行测试规范
