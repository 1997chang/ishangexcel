package cn.viewshine.cloudthree.excel.utils;

import cn.viewshine.cloudthree.excel.metadata.ColumnProperty;
import net.sf.cglib.beans.BeanMap;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;

import java.lang.reflect.Field;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.ZoneId;
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
        switch (cell.getCellType()) {
            case NUMERIC:
                //表示的数字
                if (Date.class.equals(columnField.getColumnField().getType())) {
                    //表示Data类型
                    cell.setCellValue((Date)data.get(columnField.getColumnField().getName()));
                } else if (Calendar.class.equals(columnField.getColumnField().getType())) {
                    cell.setCellValue((Calendar)data.get(columnField.getColumnField().getName()));
                } else if (LocalDateTime.class.equals(columnField.getColumnField().getType())) {
                    cell.setCellValue(Date.from(((LocalDateTime)data.get(columnField.getColumnField().getName())).
                            atZone(ZoneId.systemDefault()).toInstant()));
                } else if (LocalDate.class.equals(columnField.getColumnField().getType())) {
                    cell.setCellValue(Date.from(((LocalDate)data.get(columnField.getColumnField().getName())).
                            atStartOfDay(ZoneId.systemDefault()).toInstant()));
                } else {
                    cell.setCellValue(((Number)data.get(columnField.getColumnField().getName())).doubleValue());
                }
                break;
            case STRING:
                cell.setCellValue(data.getOrDefault(columnField.getColumnField().getName(),"").toString());
                break;
            case BOOLEAN:
                cell.setCellValue((boolean)data.getOrDefault(columnField.getColumnField().getName(),true));
                break;
            default:
                cell.setCellValue(data.getOrDefault(columnField.getColumnField().getName(), "").toString());
        }
        cell.setCellStyle(columnField.getCellStyle());
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
