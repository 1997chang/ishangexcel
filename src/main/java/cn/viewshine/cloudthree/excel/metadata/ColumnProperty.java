package cn.viewshine.cloudthree.excel.metadata;

import cn.viewshine.cloudthree.excel.annotation.ExcelField;
import cn.viewshine.cloudthree.excel.utils.FieldUtils;
import cn.viewshine.cloudthree.excel.utils.StyleUtils;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Workbook;

import java.lang.reflect.Field;
import java.util.Arrays;
import java.util.List;

/**
 * 这个表示一列的属性
 * @author changwei[changwei@viewshine.cn]
 */
public class ColumnProperty {

    /**
     * 表示单元格对应的Field
     */
    private Field columnField;

    /**
     * 这个表示一列对应的样式
     */
    private CellStyle cellStyle;

    /**
     * 一列中的Head头标题
     */
    private List<String> headString;

    /**
     * 表示一列对应的单元格类型（NUMBER，STRING，BOOLEAN，等等）
     */
    private CellType cellType;

    private int columnWidth;

    public ColumnProperty() {
    }

    public ColumnProperty(Field field, CellStyle cellStyle, Workbook workbook) {
        ExcelField excelField = field.getDeclaredAnnotation(ExcelField.class);
        this.columnField = field;
        this.cellType = FieldUtils.getCellTypeByField(field);
        this.columnWidth = excelField.columnWidth();

        //获取列的头标题信息
        this.headString = Arrays.asList(excelField.name());

        //如果format不为空的话，则创建一个新的样式，否则使用默认样式
        this.cellStyle = StyleUtils.buildNewCellStyleByDataFormat(workbook, cellStyle, excelField.format());
    }

    public Field getColumnField() {
        return columnField;
    }

    public CellStyle getCellStyle() {
        return cellStyle;
    }

    public List<String> getHeadString() {
        return headString;
    }

    public CellType getCellType() {
        return cellType;
    }

    public int getColumnWidth() {
        return columnWidth;
    }
}
