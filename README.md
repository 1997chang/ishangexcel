# ishangexcel
用于方便导出导入以及下载Excel文件

## 1.使用文档：	
①一个注解（@ExcelField）
②一个构建类（ExcelFactory）

###1.1 @ExcelField注解使用：
* name属性：必须提供，设置Excel表格的Head头。可以设置多个值，自动合并相同内容的单元格。可看下面使用案例。
* visible属性：表示该列是否显示。默认为true。
* format属性：格式化属性，例如：yyyy-MM-dd HH:mm:ss（时间），###0.00（小数），显示2位小数位数。完全采用Excel的格式化样式。
* columnWidth属性：设置当前列的宽度，默认20的字符。

###1.2 ExcelFactory的使用
```

* writeExcel(List data, String fileName)

    表示将List数据内容写入到第一个Sheet中，并且Sheet名称为sheet1，只用于写一个Sheet文件内容
    
    @param data 表示写入到Sheet数据内容
    
    @param fileName 表示写入的文件名称，可以是XLSX文件格式，这样写入的数据量较大，或者xls文件格式

* writeExcel(Map<String, List> data, String fileName)

    用于将Map数据写入到Excel文件中。一个List显示一个Sheet内容
    
    @param data：Key：Sheet表格名称，Value：Sheet对应的写入数据内容。用于写入多个Sheet
    
    @param fileName 表示将数据写入到那个文件中

* writeExcel(List data, OutputStream outputStream) 
    
    表示将List数据内容写入到第一个Sheet中，并且Sheet名称为sheet1，只用于写一个Sheet文件内容，默认是使用Xlsx文件写入
    
    @param data 表示写入到Sheet数据内容
    
    @param outputStream 表示将数据写入到文件的输出流

* void writeExcel(List data, OutputStream outputStream, boolean xssf)

    表示将List数据内容写入到第一个Sheet中，并且Sheet名称为sheet1，只用于写一个Sheet文件内容
    
    @param data 表示写入到Sheet数据内容
    
    @param outputStream 表示将数据写入到文件的输出流
    
    @param xssf 是否是XSSF文件格式，true表示XLSX文件格式，false表示XLS文件格式

* downLoadExcel(List data, HttpServletResponse response)

    完成Excel的下载任务，默认使用XLSX文件格式
    
    @param data 表示写入到Excel中文件的内容
    
    @param response response对应，完成Excel的下载任务

    注意：这里的下载，不用对response进行任何设置，内部已经设置下载的相关代码，例如：设置response.setContentType("multipart/form-data");等这样我们只要专注于下载的数据内容，直接传入response就可以了。
```