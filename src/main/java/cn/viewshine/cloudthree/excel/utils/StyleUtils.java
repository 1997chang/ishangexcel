package cn.viewshine.cloudthree.excel.utils;

import cn.viewshine.cloudthree.excel.metadata.ExcelCellStyle;
import org.apache.poi.ss.usermodel.*;

import java.util.Objects;

/**
 * 单元格的样式工具类
 * @author changwei[changwei@viewshine.cn]
 */
public final class StyleUtils {

    private StyleUtils() { }
    /**
     * 表示公共单元格样式
     * @param workbook
     * @return
     */
    private static CellStyle buildCommonCellStyle(Workbook workbook) {
        CellStyle result = workbook.createCellStyle();
        //设置字体
        Font font = workbook.createFont();
        font.setFontName("宋体");
        font.setFontHeightInPoints((short)12);
        result.setFont(font);

        //对齐方式
        result.setAlignment(HorizontalAlignment.CENTER);
        result.setVerticalAlignment(VerticalAlignment.CENTER);

        //设置单元格的边框
        result.setBorderBottom(BorderStyle.THIN);
        result.setBorderRight(BorderStyle.THIN);

        result.setLocked(true);
        return result;
    }


    public static CellStyle buildCellStyle(Workbook workbook,
                                           ExcelCellStyle cellStyle) {
        CellStyle result = buildCommonCellStyle(workbook);
        Font font = workbook.createFont();
        if (cellStyle.getFontStyle().getFontName() != null && !Objects.equals("", cellStyle.getFontStyle().getFontName()) ) {
            font.setFontName(cellStyle.getFontStyle().getFontName());
        } else {
            font.setFontName("宋体");
        }
        if (cellStyle.getFontStyle().getFontSize() != 0) {
            font.setFontHeightInPoints(cellStyle.getFontStyle().getFontSize());
        } else {
            font.setFontHeightInPoints((short)12);
        }
        if (cellStyle.getFontStyle().getColor() != null) {
            font.setColor(cellStyle.getFontStyle().getColor());
        }
        font.setBold(cellStyle.getFontStyle().getBold());
        result.setFont(font);
        if (!result.getAlignment().equals(cellStyle.getHorizontalAlignment())) {
            result.setAlignment(cellStyle.getHorizontalAlignment());
        }
        if (!result.getVerticalAlignment().equals(cellStyle.getVerticalAlignment())) {
            result.setVerticalAlignment(cellStyle.getVerticalAlignment());
        }
        return result;
    }


    /**
     * 设置Excel表格头的样式
     * 注意：如果修改字体必须创建一个新的字体，不然的话，就会回到初始字体
     * @param workbook
     * @return
     */
    public static CellStyle buildHeadCellStyle(Workbook workbook){
        CellStyle result = buildCommonCellStyle(workbook);
        //设置字体
        Font font = workbook.createFont();
        font.setFontName("宋体");
        font.setFontHeightInPoints((short)14);
        font.setBold(true);
        result.setFont(font);

        //填充效果
        result.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        result.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.index);
        return result;
    }

    /**
     * 用于创建Excel内容的单元格样式
     * @param workbook
     * @return
     */
    public static CellStyle buildContentCellStyle(Workbook workbook) {
        return buildCommonCellStyle(workbook);
    }

    /**
     * 根据dataformat构建一个全新的单元格样式，如果dataFormat为空的话，就返回默认的样式
     * @param workbook 当前工作薄
     * @param cloneStyle 要克隆的样式，默认说成默认的样式，在这个样式的基础上就行修改
     * @param dataFormat 格式化的样式
     * @return
     */
    public static CellStyle buildNewCellStyleByDataFormat(Workbook workbook, CellStyle cloneStyle, String dataFormat) {
        if (dataFormat.isEmpty()) {
            return cloneStyle;
        }
        //创建一个新的样式，然后对样式进行克隆
        CellStyle result = workbook.createCellStyle();
        result.cloneStyleFrom(cloneStyle);
        result.setDataFormat(workbook.getCreationHelper().createDataFormat().getFormat(dataFormat));
        return result;
    }

    /**
     * 从现有的单元格样式中生成一个新的，注意：不同workBook的CellStyle不同共用，必须先在当前workbook中创建一个CellStyle，然后进行复制
     * @param workbook 当前workbook
     * @param cellStyle 要克隆的样式
     * @return 返回的样式
     */
    public static CellStyle cloneCellStyle(Workbook workbook, CellStyle cellStyle) {
        CellStyle result = workbook.createCellStyle();
        result.cloneStyleFrom(cellStyle);
        return result;
    }
}
