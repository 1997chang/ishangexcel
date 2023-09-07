package cn.viewshine.cloudthree.excel;

import cn.viewshine.cloudthree.excel.context.WriteContext;
import cn.viewshine.cloudthree.excel.exception.WriteExcelException;
import org.apache.commons.collections4.CollectionUtils;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.File;
import java.util.List;

/**
 * 创建一个自主控制的
 *
 * @author moxiao
 * @Email changwei@viewshine.cn
 * @date 2021/1/23
 */
public final class StreamExcelFactory {

    /**
     * 单个Sheet最大长度为50W，超过之后自动创建另一个Sheet
     */
    private static final int MAX_COUNT_PER_SHEET = 500_000;

    /**
     * 表示当前行数
     */
    private int currentRow = 0;

    /**
     * 表示文件名称
     */
    private final String fileName;

    /**
     * 表示Sheet的前缀
     */
    private String sheetNamePrefix = "sheet";

    /**
     * 当前Sheet的索引
     */
    private int currentSheetIndex = 1;

    /**
     * 当前上下文
     */
    private final WriteContext writeContext;

    /**
     * 构建Excel流写入
     * @param fileName 文件名称
     */
    public StreamExcelFactory(String fileName, List<List<String>> headName) {
        File file = new File(fileName);
        if (file.exists()) {
            throw new WriteExcelException("传递的文件存在，请检查... 传递的文件名为：" + fileName);
        }
        this.fileName = fileName;
        boolean xssf = ExcelFactory.getFileType(fileName);
        //创建写Excel的上下文信息，包括WorkBook等
        writeContext = new WriteContext(xssf, fileName, false);
        writeContext.writeHeadToSheet(currentSheetName(), headName);
    }

    public StreamExcelFactory(String fileName, boolean useTemplate) {
        File file = new File(fileName);
        if (file.exists()) {
            throw new WriteExcelException("传递的文件存在，请检查... 传递的文件名为：" + fileName);
        }
        this.fileName = fileName;
        boolean xssf = ExcelFactory.getFileType(fileName);
        //创建写Excel的上下文信息，包括WorkBook等
        writeContext = new WriteContext(xssf, fileName, useTemplate);
    }
    
    
    public void writeHead(List<List<String>> headName) {
        writeContext.writeHeadToSheet(currentSheetName(), headName);
    }

    /**
     * 写入单行数据内容
     * @param data 写入的数据内容
     */
    public void writeData(List<String> data) {
        writeData(data, 0, true);
    }

    /**
     * 写入单行数据内容
     * @param data 写入的数据内容
     * @param useTemplate 是否使用模板中的样式
     */
    public void writeData(List<String> data, boolean useTemplate) {
        writeData(data, 0, useTemplate);
    }

    /**
     * 写入单行数据内容
     * @param data 写入的数据内容
     * @param startColumn 从哪列还是写入
     */
    public void writeData(List<String> data, int startColumn) {
        writeData(data, startColumn, true);
    }
    
    /**
     * 写入单行数据内容
     * @param data 写入的数据内容
     * @param startColumn 从哪里开始开始
     * @param useTemplate 是否使用模板中的样式                   
     */
    public void writeData(List<String> data, int startColumn, boolean useTemplate) {
        writeData(data, null, startColumn, useTemplate);
    }

    /**
     * 写入单行数据内容
     * @param data 写入的数据内容
     * @param cellStyleList 指定数据的样式
     */
    public void writeData(List<String> data, List<CellStyle> cellStyleList) {
        writeData(data, cellStyleList, 0);
    }

    public void writeData(List<String> data, List<CellStyle> cellStyleList, int startColumn) {
        writeData(data, cellStyleList, startColumn, false);
    }

    private void writeData(List<String> data, List<CellStyle> cellStyleList, int startColumn, boolean useTemplate) {
        if (CollectionUtils.isNotEmpty(cellStyleList) && data.size() != cellStyleList.size()) {
            throw new WriteExcelException("指定单元格的样式与写入数据的数量不一致");
        }
        writeContext.writeContentToSheet(currentSheetName(), data, cellStyleList, startColumn, currentSheetIndex == 1, useTemplate);
        currentRow++;
        if (currentRow >= MAX_COUNT_PER_SHEET) {
            currentSheetIndex++;
            currentRow = 0;
        }
    }

    /**
     * 获取模板的样式
     * @param row 指定行
     * @param column 指定列
     * @return 模板的指定行和指定列的样式
     */
    public CellStyle fetchTemplateCellStyle(int row, int column) {
        return writeContext.fetchCellStyle(row, column);
    }

    /**
     * 最终操作
     */
    public void finish() {
        writeContext.saveByFile(fileName);
    }

    private String currentSheetName() {
        return sheetNamePrefix + currentSheetIndex;
    }

    public void setSheetNamePrefix(String sheetNamePrefix) {
        this.sheetNamePrefix = sheetNamePrefix;
    }

    public String getFileName() {
        return fileName;
    }

    public int getCurrentRow() {
        return currentRow;
    }

    public Workbook getWorkBook() {
        return writeContext.getWorkbook();
    }
}
