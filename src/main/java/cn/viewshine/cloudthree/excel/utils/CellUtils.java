package cn.viewshine.cloudthree.excel.utils;

import cn.viewshine.cloudthree.excel.annotation.ExcelField;
import cn.viewshine.cloudthree.excel.metadata.ColumnProperty;
import net.sf.cglib.beans.BeanMap;
import org.apache.poi.ss.usermodel.*;

import java.time.*;
import java.time.format.DateTimeFormatter;
import java.time.temporal.TemporalAccessor;
import java.util.Calendar;
import java.util.Date;
import java.util.List;
import java.util.stream.IntStream;

/**
 * @author: 常伟
 * @create: 2019/8/15 21:15
 * @email: kmustchang@qq.com
 * @version: 1.0
 * @Description:
 */
public class CellUtils {

    /**
     * 向当前Sheet中写入主要内容数据
     * @param cell 当前Sheet表格
     * @param data 写入的数据内容
     * @param columnField
     */
    public static void writeContentDataAndStyle(Cell cell, BeanMap data, ColumnProperty columnField) {
        //设置单元格的样式
        cell.setCellStyle(columnField.getCellStyle());

        //设置单元格的值
        Object value = data.get(columnField.getColumnField().getName());
        if (value == null) {
            cell.setCellValue("");
            return;
        }

        switch (cell.getCellType()) {
            case NUMERIC:
                if (Date.class.equals(columnField.getColumnField().getType())) {
                    cell.setCellValue((Date) value);
                } else if (Calendar.class.equals(columnField.getColumnField().getType())) {
                    cell.setCellValue((Calendar) value);
                }
                //如果是Java8之后的时间，我们直接使用格式化字符串显示内容
                else if (TemporalAccessor.class.isAssignableFrom(columnField.getColumnField().getType())) {
                    String format = columnField.getColumnField().getDeclaredAnnotation(ExcelField.class).format();
                    if (format.isEmpty()) {
                        cell.setCellValue(value.toString());
                    } else {
                        cell.setCellValue(DateTimeFormatter.ofPattern(format).format((TemporalAccessor) value));
                    }
                } else {
                    cell.setCellValue(((Number) value).doubleValue());
                }
                break;
            case BOOLEAN:
                cell.setCellValue((Boolean) value);
                break;
            default:
                cell.setCellValue(value.toString());
        }
    }

    /**
     * 想Excel表格的Head头中添加相关的Head头数据
     * @param row 表示当前行
     * @param headData 表示行数据
     */
    public static void addOneRowHeadDataToCurrentSheet(Row row, List<String> headData, CellStyle cellStyle){
        if (headData != null && headData.size() > 0){
            IntStream.range(0, headData.size()).forEach(i -> {
                Cell cell = row.createCell(i, CellType.STRING);
                cell.setCellValue(headData.get(i));
                cell.setCellStyle(cellStyle);
            });
        }
    }

}
