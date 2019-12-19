package cn.viewshine.cloudthree.excel.context;

import cn.viewshine.cloudthree.excel.exception.WriteExcelException;
import cn.viewshine.cloudthree.excel.metadata.ColumnProperty;
import cn.viewshine.cloudthree.excel.utils.CellUtils;
import cn.viewshine.cloudthree.excel.utils.FieldUtils;
import cn.viewshine.cloudthree.excel.utils.StyleUtils;
import net.sf.cglib.beans.BeanMap;
import org.apache.commons.collections4.CollectionUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;
import java.util.stream.IntStream;

import static cn.viewshine.cloudthree.excel.utils.CellRangeUtils.mergeCell;
import static cn.viewshine.cloudthree.excel.utils.CellUtils.addOneRowHeadDataToCurrentSheet;

/**
 * 这个表示写入Excel的上下文。
 * @author changwei[changwei@viewshine.cn]
 * @version 1.0
 */
public class WriteContext {

    /**
     * 如果创建的是SXSSF的文件格式，默认的滑动窗口大小为200
     */
    private static final int DEFAULT_WINDOWS_COUNT = 200;

    /**
     * 表示写入的WorkBook
     */
    private Workbook workbook;

    /**
     * 表示所有的Excel的所有Sheet默认HEAD格式
     */
    private CellStyle defaultHeadCellStyle;

    /**
     * 表示所有的Excel的所有Sheet默认HEAD格式
     */
    private CellStyle defaultContentCellStyle;

    /**
     * 表示用户自定义的样式
     */
    private CellStyle currentSheetHeadCellStyle;

    /**
     * 表示用户自定义的样式
     */
    private CellStyle currentSheetContentCellStyle;

    /**
     * 表示是否需要HEAD头内容
     */
    private boolean needHead = true;

    /**
     * 表示是否是XSSF文件格式（.xlsx）
     */
    private boolean xssf;

    /**
     * 表示一个Sheet表格对应一个Class类，String表示Sheet名称，value表示对应的CLass类别
     */
    private Map<String, Class> sheetClass;

    /**
     * 这个表示创建一个空的SXSSF或者HSSF
     * 并设置默认的Head头样式，以及内容样式
     * @param xssf 当为true的时候，创建一个SXSSF的workBook，当为false时候，创建一个HSSF
     */
    public WriteContext(boolean xssf, Map<String, Class> sheetClass) {
        this(xssf);
        this.sheetClass = sheetClass;
    }

    public WriteContext(boolean xssf) {
        this.xssf = xssf;
        //1.创建workBook，用于写入文件内容
        if (xssf) {
            workbook = new SXSSFWorkbook(DEFAULT_WINDOWS_COUNT);
        } else {
            try {
                workbook = WorkbookFactory.create(xssf);
            } catch (IOException e) {
                e.printStackTrace();
                throw new WriteExcelException("使用HSSF创建Excel文件失败。。。");
            }
        }
        //设置样式
        defaultHeadCellStyle = StyleUtils.buildHeadCellStyle(workbook);
        defaultContentCellStyle = StyleUtils.buildContentCellStyle(workbook);

        //TODO 设置用户自定义的样式
    }

    /**
     * 将数据内容写入到fileName文件中，每一个Map值对应一个Sheet表格
     * @param data 表示写入到Excel数据的内容
     * @param fileName 表示Excel文件路径地址以及名称
     */
    public void write(Map<String, List> data, String fileName) {
        writeAllDataToExcel(data);
        //这里使用try-with-resource
        saveByFile(fileName);
    }

    /**
     * 用于将数据内容以及表格头数据内容写入到指定文件中
     * @param data 表格的数据内容
     * @param headName 表格列头的数据
     * @param fileName 写入的文件
     */
    public void write(Map<String, List<List<String>>> data, Map<String, List<List<String>>> headName, String fileName) {
        writeAllDataToExcel(data, headName);
        saveByFile(fileName);
    }

    private void writeAllDataToExcel(Map<String, List<List<String>>> data, Map<String, List<List<String>>> headName) {
        data.forEach((sheetName, contentData) -> {
            //1.为每一个List数据，创建一个Sheet
            Sheet currentSheet = workbook.createSheet(sheetName);
            IntStream.range(0, headName.size()).forEach(i -> currentSheet.setColumnWidth(i, 20 * 256));
            writeHeadToSheet(currentSheet, headName.get(sheetName));
            writeContentToSheet(currentSheet, contentData);
        });
    }

    /**
     * 将数据内容写入到输出流中，每一个Map值对应一个Sheet表格
     * @param data 写入的数据
     * @param outputStream 输出流
     */
    public void write(Map<String, List> data, OutputStream outputStream) {
        writeAllDataToExcel(data);
        saveByStream(outputStream);
    }

    /**
     * 将全部数据写入到Excel表格。注意一个Map.Entry对应一个表格Sheet
     * @param data 写入的数据
     */
    private void writeAllDataToExcel(Map<String, List> data) {
        data.forEach((sheetName, sheetData) -> {
            //1.为每一个List数据，创建一个Sheet
            Sheet currentSheet = workbook.createSheet(sheetName);
            List<ColumnProperty> classFieldList = FieldUtils.getAllColumnPropertyOfSingleClass(sheetClass.get(sheetName),
                            getCurrentActiveContentCellStyle(), workbook);

            //2.设置列宽
            IntStream.range(0,classFieldList.size()).forEach(i -> currentSheet.setColumnWidth(i,
                    classFieldList.get(i).getColumnWidth() * 256));

            //3.如果需要写入Excel的HEAD的话
            if (needHead) {
                writeHeadToSingleSheet(currentSheet, classFieldList);
            }

            //4.将数据内容写入到Sheet中
            wirteContentToSingleSheet(currentSheet, sheetData, classFieldList);
        });
    }

    /**
     * 在一个Sheet中写入数据内容，也就是将一个List写入到一个Sheet页中。
     * @param sheet 表示将数据内容写入到那个Sheet中
     * @param dataList 表示一个Sheet中的所有数据内容
     * @param columnField 表示sheet中每一列的样式内容
     */
    private void wirteContentToSingleSheet(Sheet sheet, List dataList, List<ColumnProperty> columnField) {
        if (CollectionUtils.isEmpty(dataList)) {
            return;
        }
        //确定要写入的行数
        int lastRowNum = computeLastRow(sheet);
        //遍历每一行的数据内容
        for (Object data : dataList) {
            Row row = sheet.createRow(lastRowNum++);
            //遍历所有的列
            IntStream.range(0, columnField.size()).forEach(i -> {
                Cell cell = row.createCell(i, columnField.get(i).getCellType());
                CellUtils.writeContentDataAndStyle(cell, BeanMap.create(data), columnField.get(i));
            });
        }
    }

    /**
     * 用于向Sheet中写入数据内容
     * @param sheet 写入的Sheet表格对象
     * @param sheetData 写入的数据内容
     */
    private void writeContentToSheet(Sheet sheet, List<List<String>> sheetData) {
        int lastRowNum = computeLastRow(sheet);
        //遍历每一行的数据内容
        for (List<String> data: sheetData) {
            Row row = sheet.createRow(lastRowNum++);
            //遍历所有的列
            IntStream.range(0, data.size()).forEach(i -> {
                Cell cell = row.createCell(i, CellType.STRING);
                cell.setCellStyle(getCurrentActiveContentCellStyle());
                cell.setCellValue(data.get(i));
            });
        }
    }

    /**
     * 用于完成在一个Sheet中写入文件头内容
     * @param sheet 表示当前写入的Sheet表格
     * @param fieldList 表示一个Sheet中所有列的属性
     */
    private void writeHeadToSingleSheet(Sheet sheet, List<ColumnProperty> fieldList) {
        //获取所有字段在ExcelField注解中的value值,从而合并单元格
        writeHeadToSheet(sheet, fieldList.stream().map(ColumnProperty::getHeadString).collect(Collectors.toList()));
    }

    /**
     * 用于向指定的Sheet表格中
     * @param sheet
     * @param headName
     */
    private void writeHeadToSheet(Sheet sheet, List<List<String>> headName) {
        //列标题中最大行数，以及开始行
        int rowMaxCount = headName.parallelStream().mapToInt(List::size).max().orElse(0);
        int startRow = sheet.getLastRowNum();

        //2.合并Head中对应的单元格
        if (rowMaxCount > 1) {
            mergeCell(startRow, rowMaxCount, headName, sheet);
        }

        //3.填充HEAD数据头内容
        IntStream.range(0, rowMaxCount).forEach(i -> {
            Row row = sheet.createRow(startRow + i);
            addOneRowHeadDataToCurrentSheet(row, headName.stream().map(list->list.get(i)).collect(Collectors.toList()), getCurrentActiveHeadCellStyle());
        });
    }

    /**
     * 得到当前激活的头样式
     * @return 当前激活的头样式
     */
    private CellStyle getCurrentActiveHeadCellStyle() {
        return currentSheetHeadCellStyle != null ? currentSheetHeadCellStyle : defaultHeadCellStyle;
    }

    /**
     * 获得当前激活的内容样式
     * @return 激活的内容样式
     */
    private CellStyle getCurrentActiveContentCellStyle() {
        return currentSheetContentCellStyle != null ? currentSheetContentCellStyle : defaultContentCellStyle;
    }

    /**
     * 将当前workBook写入到文件中
     * @param fileName 文件名
     */
    private void saveByFile(String fileName) throws WriteExcelException {
        // You must close the OutputStream yourself. HSSF does not close it for you.
        try (FileOutputStream out = new FileOutputStream(fileName)) {
            workbook.write(out);
        } catch (IOException e) {
            e.printStackTrace();
            throw new WriteExcelException("写入文件出现错误，方法：saveByFile", e);
        } finally {
            try {
                if (xssf) {
                    //Note that SXSSF allocates temporary files that you must always clean up explicitly, by calling the dispose method.
                    //注意：如果是SXSSF的话，必须显示调用workbook的dispose方法
                    ((SXSSFWorkbook)workbook).dispose();
                }
                if (workbook != null) {
                    workbook.close();
                }
            } catch (IOException e) {
                e.printStackTrace();
                throw new WriteExcelException("关闭操作出现问题", e);
            }
        }
    }

    /**
     * 将整个Excel保存到输出流中，
     * 注意：输出流必须我们手动关闭
     * @param outputStream 输出流
     */
    private void saveByStream(OutputStream outputStream) {
        try {
            workbook.write(outputStream);
        } catch (IOException e) {
            e.printStackTrace();
            throw new WriteExcelException("写入文件流出现错误，方法：saveByStream", e);
        } finally {
            try {
                //NOTE: You must close the OutputStream yourself. HSSF does not close it for you.
                if (outputStream != null) {
                    outputStream.close();
                }
                if (xssf) {
                    ((SXSSFWorkbook)workbook).dispose();
                }
                if (workbook != null) {
                    workbook.close();
                }
            } catch (IOException e) {
                e.printStackTrace();
                throw new WriteExcelException("关闭操作出现问题", e);
            }
        }
    }

    /**
     * 用于计算当前Sheet中最后一个有效行
     * @param sheet 当前sheet表格
     * @return 行数
     */
    private static int computeLastRow(Sheet sheet) {
        //确定要写入的行数
        int lastRowNum = sheet.getLastRowNum();
        //如果当前行不等于0，或者第零行数据不为空的话，让行数+1.因为Sheet返回的是最后有数据的一行下标
        boolean rowNeedPlusOne = lastRowNum !=0 || sheet.getRow(0) != null;
        if (rowNeedPlusOne) {
            lastRowNum++;
        }
        return lastRowNum;
    }

}
