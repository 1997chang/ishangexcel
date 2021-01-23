package cn.viewshine.cloudthree.excel;

import cn.viewshine.cloudthree.excel.context.WriteContext;
import cn.viewshine.cloudthree.excel.exception.WriteExcelException;

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
        writeContext = new WriteContext(xssf, fileName);
        writeContext.writeHeadToSheet(currentSheetName(), headName);
    }

    /**
     * 写入单行数据内容
     * @param data 数据内容
     */
    public void writeData(List<String> data) {
        writeContext.writeContentToSheet(currentSheetName(), data);
        currentRow++;
        if (currentRow >= MAX_COUNT_PER_SHEET) {
            currentSheetIndex++;
            currentRow = 0;
        }
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
}
