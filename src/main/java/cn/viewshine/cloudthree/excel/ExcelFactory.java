package cn.viewshine.cloudthree.excel;

import cn.viewshine.cloudthree.excel.exception.WriteExcelException;
import cn.viewshine.cloudthree.excel.write.WriteExcel;

import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.net.URLDecoder;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.Collections;
import java.util.List;
import java.util.Map;

/**
 * 表示Excel的工厂类，用于创建新的Excel（XLS或者XLSX），用于读取XLS或者XLSX文件的内容
 * @author changwei[changwei@viewshine.cn]
 * @version 1.0
 */
public final class ExcelFactory {

    /**
     * 用于完成下载时候，定义文件的名称
     */
    private static final DateTimeFormatter YYYY_MM_DD_HH_MM_SS_PATTERN =
            DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss");

    /**
     * 用于将Map数据写入到Excel文件中。
     * 注意：默认是使用OOXML写入到Excel文件中，文件的后缀名为：.xlsx
     * @param data：Key：Sheet表格名称，Value：Sheet对应的写入数据内容。用于写入多个Sheet
     * @param fileName 表示将数据写入到那个文件中
     */
    public static void writeExcel(Map<String, List<?>> data, String fileName) {
        boolean xssf = getFileType(fileName);
        WriteExcel.writeExcelByFileName(data, fileName, xssf);
    }

    /**
     * 对Excel数据进行导出，这里进行的是多Sheet导出
     * @param data 表示导出的数据内容，每一个Sheet对应一个MAP
     * @param headName 表示各个Sheet中表格的表头
     * @param fileName 表示文件名称
     */
    public static void writeExcel(Map<String, List<List<String>>> data, Map<String, List<List<String>>> headName,
                                  String fileName) {
        boolean xssf = getFileType(fileName);
        WriteExcel.writeExcelByFileName(data, headName, fileName, xssf);
    }

    /**
     * 表示将List数据内容写入到第一个Sheet中，并且Sheet名称为sheet1，只用于写一个Sheet文件内容
     * 注意：默认是使用OOXML写入到Excel文件中，文件的后缀名为：.xlsx
     * @param data 表示写入到Sheet数据内容
     * @param fileName 表示写入的文件名称
     */
    public static void writeExcel(List<?> data, String fileName) {
        writeExcel(Collections.singletonMap("sheet1",data), fileName);
    }

    /**
     * 用于将Map数据写入到Excel文件中。
     * @param data Key：Sheet表格名称，Value：Sheet对应的写入数据内容。用于写入多个Sheet
     * @param outputStream 表示写入数据的输出流
     */
    public static void writeExcel(Map<String, List<?>> data, OutputStream outputStream, boolean xssf) {
        WriteExcel.wirteExcelByStream(data, outputStream, xssf);
    }

    /**
     * 用于将Map数据写入到Excel文件中。默认使用XLSX文件格式写入
     * @param data Key：Sheet表格名称，Value：Sheet对应的写入数据内容。用于写入多个Sheet
     * @param outputStream 表示写入数据的输出流
     */
    public static void writeExcel(Map<String, List<?>> data, OutputStream outputStream) {
        writeExcel(data, outputStream, true);
    }

    /**
     * 表示将List数据内容写入到第一个Sheet中，并且Sheet名称为sheet1，只用于写一个Sheet文件内容，默认是使用Xlsx文件写入
     * @param data 表示写入到Sheet数据内容
     * @param outputStream 表示将数据写入到文件的输出流
     */
    public static void writeExcel(List<?> data, OutputStream outputStream) {
        writeExcel(Collections.singletonMap("sheet1",data), outputStream, true);
    }

    /**
     * 表示将List数据内容写入到第一个Sheet中，并且Sheet名称为sheet1，只用于写一个Sheet文件内容
     * @param data 表示写入到Sheet数据内容
     * @param outputStream 表示将数据写入到文件的输出流
     * @param xssf 是否是XSSF文件格式，true
     */
    public static void writeExcel(List<?> data, OutputStream outputStream, boolean xssf) {
        writeExcel(Collections.singletonMap("sheet1",data), outputStream, xssf);
    }

    /**
     * 完成Excel的下载任务，默认使用XLSX文件格式
     * @param data 表示写入到Excel中文件的内容
     * @param response response对应，完成Excel的下载任务
     */
    public static void downLoadExcel(List<?> data, HttpServletResponse response) {
        downloadExcel(Collections.singletonMap( "sheet", data), true, response);
    }

    /**
     * 完成Excel的下载任务
     * @param data 表示写入到Excel中文件的内容
     * @param xssf 表示是否是XLSX文件格式，TRUE表示是XLSX文件格式，FLASE表示不是XLSX文件格式
     * @param response response对应，完成Excel的下载任务
     */
    public static void downloadExcel(Map<String, List<?>> data, boolean xssf, HttpServletResponse response) {
        //设置下载文件的名称，名称为当前时间
        StringBuilder fileName = new StringBuilder(LocalDateTime.now().format(YYYY_MM_DD_HH_MM_SS_PATTERN));
        //设置下载文件的格式
        try {
            response.setContentType("multipart/form-data");
            response.setCharacterEncoding("utf-8");
            if (xssf) {
                fileName.append(".xlsx");
            } else {
                fileName.append(".xls");
            }
            response.setHeader("Content-disposition", "attachment;filename=" + URLDecoder.decode(fileName.toString(), "utf-8"));

            WriteExcel.wirteExcelByStream(data, response.getOutputStream(), xssf);
        } catch (UnsupportedEncodingException e) {
            throw new WriteExcelException("文件名称解析错误，fileName: " + fileName.toString(), e);
        } catch (IOException e) {
            throw new WriteExcelException("获取response的输出流错误", e);
        }
    }

    /**
     * 用于获取文件的类型，如果是xlsx表示为XSSF模式，如果是xls表示为非XSSF模式
     * @param fileName 表示文件名称
     * @return 是否是XSSF模式
     */
    static boolean getFileType(String fileName){
        if (fileName.endsWith(".xlsx")) {
            return true;
        } else if (fileName.endsWith(".xls")) {
            return false;
        } else {
            throw new WriteExcelException("写入的文件名称不合法，后缀是xlsx或者xls");
        }
    }

}
