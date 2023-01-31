package cn.viewshine.cloudthree.excel.metadata;

import lombok.Getter;
import lombok.Setter;
import org.apache.poi.ss.usermodel.*;

/**
 * 用于表示用户自定义的单元格的样式
 * @see org.apache.poi.ss.usermodel.CellStyle
 * @author changwei[changwei@viewshine.cn]
 */
@Setter
@Getter
public class ExcelCellStyle {

    /**
     * 单元格的字体样式
     */
    private FontStyle fontStyle;

    /**
     * 设置填充的模式样式
     */
    private FillPatternType fillPatternType;

    /**
     * 设置背景颜色
     */
    private Short backgroundColor;

    /**
     * 设置前景颜色
     */
    private Short foregroundColor;

    /**
     * 设置水平对其方式
     */
    private HorizontalAlignment horizontalAlignment;

    /**
     * 设置垂直对其方式
     */
    private VerticalAlignment verticalAlignment;

    /**
     *
     */
    private Boolean quotePrefix;

    /**
     * 设置单元内容是否在一行中显示
     */
    private Boolean wrapped;

    /**
     * 是否隐藏
     */
    private Boolean hidden;

    /**
     * 设置单元格的样式是否进行锁住
     */
    private Boolean lock;

    /**
     * 设置字体旋转的度数。
     * 注意：HSSF的值在-90~90之间，XSSF在0~180之间
     */
    private Short rotation;

    /**
     * 设置前置空格的数量
     */
    private Short indent;

    /**
     * 设置左边框样式
     */
    private BorderStyle borderLeft;

    /**
     * 设置右边框的样式
     */
    private BorderStyle borderRight;

    /**
     * 设置底边框的样式
     */
    private BorderStyle borderBottom;

    /**
     * 设置上边框的样式
     */
    private BorderStyle borderTop;

    /**
     * 设置左边框的颜色
     */
    private Short leftBorderColor;

    /**
     * 设置右边框的颜色
     */
    private Short rightBorderColor;

    /**
     * 设置底边框的颜色
     */
    private Short bottomBorderColor;

    /**
     * 设置上边框的颜色
     */
    private Short topBorderColor;

    /**
     * 是否自动进行收缩
     */
    private Boolean shrinkToFit;

}
