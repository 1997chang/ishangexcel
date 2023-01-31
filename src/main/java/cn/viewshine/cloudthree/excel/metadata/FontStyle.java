package cn.viewshine.cloudthree.excel.metadata;

import cn.viewshine.cloudthree.excel.utils.StringUtils;
import lombok.Getter;
import lombok.Setter;

/**
 * 在POI中，字体样式每次进行设置，都必须放置一个新的字体，并且无法获取当前字体的样式。在POI中字体设置：
 * @see org.apache.poi.ss.usermodel.Font 字体样式
 * @author ChangWei[changwei@viewshine.cn]
 */
@Getter
@Setter
public class FontStyle {

    /**
     * 表示字体名称
     */
    private String fontName;

    /**
     * 表示字体大小
     */
    private Short fontHeightInPoints;

    /**
     * 表示是否是粗体
     */
    private Boolean bold;

    /**
     * 表示是否是斜体
     */
    private Boolean italic;

    /**
     * 用于设置字体的下划线
     */
    private Byte underline;

    /**
     * 设置字体的颜色
     * @see org.apache.poi.ss.usermodel.IndexedColors
     */
    private Short color;

    /**
     * 字体大小
     */
    private short fontSize;

    public static FontStyle defaultHeadFontStyle() {
        FontStyle fontStyle = new FontStyle();
        fontStyle.setBold(true);
        fontStyle.setFontName("宋体");
        fontStyle.setFontHeightInPoints((short) 15);
        return fontStyle;
    }

    public static FontStyle defaultContentFontStyle() {
        FontStyle fontStyle = new FontStyle();
        fontStyle.setBold(false);
        fontStyle.setFontName("宋体");
        fontStyle.setFontHeightInPoints((short) 15);
        return fontStyle;
    }

    /**
     * 用于将target没有设置的样式，设置为source指定的样式
     * @param source 原字体样式
     * @param target 目标字体样式
     */
    public static void merge(FontStyle source, FontStyle target) {
        if (target == null || source == null) {
            return;
        }
        if (StringUtils.isBlank(target.getFontName()) && !StringUtils.isBlank(source.getFontName())) {
            target.setFontName(source.getFontName());
        }
        if (target.getBold() == null && source.getBold() != null) {
            target.setBold(source.getBold());
        }
        if (target.getItalic() == null && source.getItalic() != null) {
            target.setItalic(source.getItalic());
        }
        if (target.getFontHeightInPoints() == null && source.getFontHeightInPoints() != null) {
            target.setFontHeightInPoints(source.getFontHeightInPoints());
        }
        if (target.getColor() == null && source.getColor() != null) {
            target.setColor(source.getColor());
        }
        if (target.getUnderline() == null && source.getUnderline() != null) {
            target.setUnderline(source.getUnderline());
        }
    }

}
