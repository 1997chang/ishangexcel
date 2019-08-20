package cn.viewshine.cloudthree.excel.metadata;

import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.VerticalAlignment;

/**
 * @Description: 用于表示用户自定义的单元格的样式
 * @Author: ChangWei
 * @Email: changwei@viewshine.cn
 * @Date: 2019/8/15
 */
public class CellStyle {

    /**
     * 表示字体名称
     */
    private String fontName;

    /**
     * 表示字体大小
     */
    private short fontSize;

    /**
     * 表示是否是粗体
     */
    private boolean bold;

    /**
     * 设置字体颜色
     */
    private IndexedColors fontColor;

    /**
     * 设置背景颜色
     */
    private IndexedColors backGroundColor;

    /**
     * 设置水平对其方式
     */
    private HorizontalAlignment horizontalAlignment;

    /**
     * 设置垂直对其方式
     */
    private VerticalAlignment verticalAlignment;


    public CellStyle setFontName(String fontName) {
        this.fontName = fontName;
        return this;
    }

    public CellStyle setFontSize(short fontSize) {
        this.fontSize = fontSize;
        return this;
    }

    public CellStyle setBold(boolean bold) {
        this.bold = bold;
        return this;
    }

    public CellStyle setFontColor(IndexedColors fontColor) {
        this.fontColor = fontColor;
        return this;
    }

    public CellStyle setBackGroundColor(IndexedColors backGroundColor) {
        this.backGroundColor = backGroundColor;
        return this;
    }

    public CellStyle setHorizontalAlignment(HorizontalAlignment horizontalAlignment) {
        this.horizontalAlignment = horizontalAlignment;
        return this;
    }

    public CellStyle setVerticalAlignment(VerticalAlignment verticalAlignment) {
        this.verticalAlignment = verticalAlignment;
        return this;
    }
}
