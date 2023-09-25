package cn.viewshine.cloudthree.excel.context;

import cn.viewshine.cloudthree.excel.exception.WriteExcelException;
import cn.viewshine.cloudthree.excel.metadata.ColumnProperty;
import cn.viewshine.cloudthree.excel.utils.CellUtils;
import cn.viewshine.cloudthree.excel.utils.FieldUtils;
import cn.viewshine.cloudthree.excel.utils.StringUtils;
import cn.viewshine.cloudthree.excel.utils.StyleUtils;
import net.sf.cglib.beans.BeanMap;
import org.apache.commons.collections4.CollectionUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.*;
import java.util.stream.Collectors;
import java.util.stream.IntStream;

import static cn.viewshine.cloudthree.excel.utils.CellRangeUtils.mergeCell;
import static cn.viewshine.cloudthree.excel.utils.CellUtils.addOneRowDataToCurrentSheet;

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

    private static final int XSSF_MAX_COUNT_PER_SHEET = 500_000;

    private static final int NOT_XSSF_MAX_COUNT_PRE_SHEET = 60_000;

    /**
     * 默认列宽
     */
    private static final int COLUMN_WIDTH = 20;

    /**
     * 表头表尾的最大长度
     */
    private static final int HEAD_TAIL_MAX_LENGTH = 8;

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
     * 是否使用模板
     */
    private XSSFWorkbook useTemplateWorkBook;
    
    /**
     * 这个表示创建一个空的SXSSF或者HSSF
     * 并设置默认的Head头样式，以及内容样式
     * @param xssf 当为true的时候，创建一个SXSSF的workBook，当为false时候，创建一个HSSF
     */
    public WriteContext(boolean xssf, Map<String, Class> sheetClass) {
        this(xssf, (String) null, false);
        this.sheetClass = sheetClass;
    }

    public WriteContext(boolean xssf, String filePath, boolean useTemplate) {
        this.xssf = xssf;
        boolean fileExits = Objects.nonNull(filePath) && Files.exists(Paths.get(filePath));
        //1.创建workBook，用于写入文件内容
        if (xssf) {
            if (fileExits || useTemplate) {
                try {
                    InputStream useTemplateInputStream;
                    if (fileExits) {
                        useTemplateInputStream = WriteContext.class.getClassLoader().getResourceAsStream(filePath);
                    } else {
                        useTemplateInputStream = WriteContext.class.getClassLoader().getResourceAsStream("exceltemplate/CustomizeReport.xlsx");
                    }
                    if (useTemplateInputStream != null) {
                        useTemplateWorkBook = new XSSFWorkbook(useTemplateInputStream);
                    }
                } catch (IOException e) {
                    e.printStackTrace();
                    throw new WriteExcelException("使用XSSF打开已存在文件错误.");
                }
            }
            workbook = new SXSSFWorkbook(null, DEFAULT_WINDOWS_COUNT, true);
        } else {
            try {
                File exitsFile = null;
                if (fileExits) {
                    exitsFile = new File(filePath);
                }
                workbook = WorkbookFactory.create(exitsFile);
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
    public void write(Map<String, List<?>> data, String fileName) {
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
    public void write(Map<String, List<List<String>>> data, 
                      Map<String, List<List<String>>> headName,
                      String title,
                      List<String> head,
                      List<String> tail,
                      String fileName) {
        writeAllDataToExcel(data, headName, title, head, tail);
        saveByFile(fileName);
    }

    private void writeAllDataToExcel(Map<String, List<List<String>>> data, 
                                     Map<String, List<List<String>>> headName,
                                     String title,
                                     List<String> head,
                                     List<String> tail) {
        data.forEach((sheetName, value) -> {
            Sheet currentSheet = workbook.getSheet(sheetName);
            boolean createSheet = Objects.isNull(currentSheet) ||
                    exceedSheetMaxCount(computeLastRow(currentSheet), value.size());
            if (createSheet) {
                Sheet sheet = createSheet(sheetName, true);
                if (!StringUtils.isBlank(title)) {
                    writeContentToSheet(sheetName, Collections.singletonList(title), null, 0, true, true);
                }
                writeTableHeadOrTail(sheetName, head, false);
                if (CollectionUtils.isNotEmpty(headName.get(sheetName))) {
                    writeHeadToSheet(sheet, headName.get(sheetName));
                }
                currentSheet = sheet;
            }
            writeContentToSheet(currentSheet, value);
            writeTableHeadOrTail(sheetName, tail, true);
        });
    }
    
    private void writeTableHeadOrTail(String sheetName, List<String> data, boolean addSpaceLine) {
        if (CollectionUtils.isNotEmpty(data)) {
            CellStyle firstCellStyle = fetchCellStyle(1, 1);
            CellStyle secondCellStyle = fetchCellStyle(1, 2);
            if (addSpaceLine) {
                writeContentToSheet(sheetName, Collections.singletonList(""), 
                        Collections.singletonList(secondCellStyle), 0, true, false);
            }
            if (data.size() > HEAD_TAIL_MAX_LENGTH) {
                List<CellStyle> cellStyleList = Arrays.asList(firstCellStyle, secondCellStyle,
                        firstCellStyle, secondCellStyle, firstCellStyle, secondCellStyle, firstCellStyle, secondCellStyle);
                List<List<String>> partition = split(data);
                for (int i = 0; i < partition.size() - 1; i++) {
                    writeContentToSheet(sheetName, partition.get(i), cellStyleList, 1, true, false);
                }
                data = partition.get(partition.size() - 1);
            }
            List<CellStyle> cellStyleList = new ArrayList<>(data.size());
            while (cellStyleList.size() < data.size()) {
                cellStyleList.add(firstCellStyle);
                cellStyleList.add(secondCellStyle);
            }
            writeContentToSheet(sheetName, data, cellStyleList, 1, true, false);
        }
    }

    /**
     * 将数据内容写入到输出流中，每一个Map值对应一个Sheet表格
     * @param data 写入的数据
     * @param outputStream 输出流
     */
    public void write(Map<String, List<?>> data, OutputStream outputStream) {
        writeAllDataToExcel(data);
        saveByStream(outputStream);
    }

    /**
     * 将全部数据写入到Excel表格。注意一个Map.Entry对应一个表格Sheet
     * @param data 写入的数据
     */
    private void writeAllDataToExcel(Map<String, List<?>> data) {
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
            writeContentToSingleSheet(currentSheet, sheetData, classFieldList);
        });
    }

    /**
     * 在一个Sheet中写入数据内容，也就是将一个List写入到一个Sheet页中。
     * @param sheet 表示将数据内容写入到那个Sheet中
     * @param dataList 表示一个Sheet中的所有数据内容
     * @param columnField 表示sheet中每一列的样式内容
     */
    private void writeContentToSingleSheet(Sheet sheet, List<?> dataList, List<ColumnProperty> columnField) {
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
    public void writeContentToSheet(Sheet sheet, List<List<String>> sheetData) {
        int lastRowNum = computeLastRow(sheet);
        //遍历每一行的数据内容
        for (List<String> data: sheetData) {
            Row row = sheet.createRow(lastRowNum++);
            addOneRowDataToCurrentSheet(row, data, null, defaultContentCellStyle, 0, null);
        }
    }
    

    /**
     * 写入一行数据到指定的sheet中，他并不会保存文件
     * @param sheetName sheet名称
     * @param sheetData 一行数据内容
     * @param startColumn 从那列开始写
     */
    public void writeContentToSheet(String sheetName, 
                                    List<String> sheetData, 
                                    List<CellStyle> cellStyleList, 
                                    int startColumn, 
                                    boolean firstSheet,
                                    boolean useTemplate) {
        writeRowToSheet(sheetName, sheetData, cellStyleList, startColumn, getCurrentActiveContentCellStyle(), 
                firstSheet, useTemplate);
    }

    /**
     * 写入一行数据到指定的sheet中，他并不会保存文件，并使用指定的样式
     * @param sheetName sheet名称
     * @param sheetData 一行数据内容
     * @param cellStyleList 每一列的样式
     * @param startColumn 从那列开始写
     * @param defaultContentCellStyle 如果列样式为空，使用默认的样式
     */
    public void writeRowToSheet(String sheetName, 
                                List<String> sheetData, 
                                List<CellStyle> cellStyleList, 
                                int startColumn, 
                                CellStyle defaultContentCellStyle, 
                                boolean firstSheet,
                                boolean useTemplate) {
        Sheet sheet = createSheet(sheetName, firstSheet);
        int lastRowNum = computeLastRow(sheet);
        Row row = sheet.createRow(lastRowNum);
        if (useTemplateWorkBook != null && firstSheet) {
            XSSFSheet templateSheet = useTemplateWorkBook.getSheetAt(0);
            XSSFRow templateSheetRow = templateSheet.getRow(lastRowNum);
            if (templateSheetRow != null) {
                row.setHeight(templateSheetRow.getHeight());
            }
        }
        //遍历所有的列
        addOneRowDataToCurrentSheet(row, sheetData, cellStyleList, defaultContentCellStyle, startColumn,
                useTemplate ? useTemplateWorkBook : null);
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
    public void writeHeadToSheet(Sheet sheet, List<List<String>> headName) {
        if (Objects.isNull(headName)) {
            return;
        }
        //列标题中最大行数，以及开始行
        int rowMaxCount = headName.parallelStream().mapToInt(List::size).max().orElse(0);
        int startRow = computeLastRow(sheet);

        //2.合并Head中对应的单元格
        if (rowMaxCount > 1) {
            mergeCell(startRow, rowMaxCount, headName, sheet);
        }

        //3.填充HEAD数据头内容
        IntStream.range(0, rowMaxCount).forEach(i -> {
            Row row = sheet.createRow(startRow + i);
            addOneRowDataToCurrentSheet(row, headName.stream().map(List -> List.get(i)).collect(Collectors.toList()), 
                    null, getCurrentActiveHeadCellStyle(), 0, null);
        });
    }

    /**
     * 写入Excel头到指定sheetName中
     * @param sheetName sheet文件名
     * @param headName 写入的headName头内容
     */
    public void writeHeadToSheet(String sheetName, List<List<String>> headName) {
        Sheet sheet = workbook.getSheet(sheetName);
        if (Objects.isNull(sheet)) {
            sheet = workbook.createSheet(sheetName);
            sheet.setDefaultColumnWidth(COLUMN_WIDTH);
        }
        if (CollectionUtils.isNotEmpty(headName)) {
            writeHeadToSheet(sheet, headName);
        }
    }

    public CellStyle fetchCellStyle(int row, int column) {
        return CellUtils.fetchCellStyle(row, column, workbook, useTemplateWorkBook);
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
    public void saveByFile(String fileName) throws WriteExcelException {
        // You must close the OutputStream yourself. HSSF does not close it for you.
        try {
            Files.createDirectories(Paths.get(fileName).getParent());
        } catch (IOException e) {
            e.printStackTrace();
            throw new WriteExcelException("文件目录不存在");
        }

        File file = new File(fileName);
        boolean fileExits = file.exists();
        if (fileExits) {
            file = new File(getBackFileName(fileName));
        }
        try (FileOutputStream out = new FileOutputStream(file)) {
            workbook.write(out);
        } catch (IOException e) {
            e.printStackTrace();
            throw new WriteExcelException("写入文件出现错误，方法：saveByFile", e);
        } finally {
            try {
                if (workbook instanceof SXSSFWorkbook) {
                    //Note that SXSSF allocates temporary files that you must always clean up explicitly, by calling the dispose method.
                    //注意：如果是SXSSF的话，必须显示调用workbook的dispose方法
                    ((SXSSFWorkbook)workbook).dispose();
                }
                if (workbook != null) {
                    workbook.close();
                }
                if (useTemplateWorkBook != null) {
                    useTemplateWorkBook.close();
                }
            } catch (IOException e) {
                e.printStackTrace();
                throw new WriteExcelException("关闭操作出现问题", e);
            }
        }
        if (fileExits) {
            new File(fileName).delete();
            file.renameTo(new File(fileName));
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
                if (useTemplateWorkBook != null) {
                    useTemplateWorkBook.close();
                }
            } catch (IOException e) {
                throw new WriteExcelException("关闭操作出现问题", e);
            }
        }
    }

    /**
     * 用于计算当前Sheet中最后一个有效行
     * @param sheet 当前sheet表格
     * @return 行数
     */
    private static int computeLastRow(final Sheet sheet) {
        //确定要写入的行数
        return sheet.getPhysicalNumberOfRows();
    }

    /**
     * 确定当前Sheet是否超过最大单页Sheet大小
     * @param currentRow 当前Sheet的大小
     * @param dataSize 写入数据的个数
     * @return true 超过，需要分页，false，可以直接写入
     */
    private boolean exceedSheetMaxCount(int currentRow, int dataSize) {
        return xssf ? currentRow + dataSize > XSSF_MAX_COUNT_PER_SHEET :
                currentRow + dataSize > NOT_XSSF_MAX_COUNT_PRE_SHEET;
    }

    /**
     * 用于确定最终的sheet的名称
     * @param baseSheetName 基本的sheet名称
     * @return sheet名称
     */
    private String getSheetName(String baseSheetName) {
        String sheetName = baseSheetName;
        while(Objects.nonNull(workbook.getSheet(sheetName))) {
            sheetName += "1";
        }
        return sheetName;
    }

    /**
     * 如果文件存在得到备份文件名
     * @param fileName 文件名
     * @return 备份文件名
     */
    private static String getBackFileName(String fileName) {
        int suffix = fileName.lastIndexOf(".");
        return fileName.substring(0, suffix) +"_back." + fileName.substring(suffix + 1);
    }

    private Sheet createSheet(String sheetName) {
        return createSheet(sheetName, false);
    }

    private Sheet createSheet(String sheetName, boolean addMergedRegion) {
        Sheet sheet = workbook.getSheet(sheetName);
        if (Objects.isNull(sheet)) {
            sheet = workbook.createSheet(sheetName);
            if (addMergedRegion && useTemplateWorkBook != null) {
                useTemplateWorkBook.getSheetAt(0).getMergedRegions().forEach(sheet::addMergedRegion);
            }
            sheet.setDefaultColumnWidth(COLUMN_WIDTH);
        }
        return sheet;
    }

    public Workbook getWorkbook() {
        return workbook;
    }

    private static <T> List<List<T>> split(List<T> data) {
        List<List<T>> result = Collections.emptyList();
        if (CollectionUtils.isNotEmpty(data)) {
            int length = (data.size() - 1) / HEAD_TAIL_MAX_LENGTH + 1;
            result = new ArrayList<>(length);
            int i = 0;
            int start = 0;
            while (i < length) {
                int end = Math.min(start + HEAD_TAIL_MAX_LENGTH, data.size());
                result.add(data.subList(start, end));
                start += HEAD_TAIL_MAX_LENGTH;
                i++;
            }
        }
        return result;
    }
}
