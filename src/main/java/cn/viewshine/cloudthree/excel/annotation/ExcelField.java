package cn.viewshine.cloudthree.excel.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * 用于导入/导出Excel表格中Header头部分的信息
 * @author changwei[changwei@viewshine.cn]
 * @version 1.0
 */
@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.FIELD)
public @interface ExcelField {

    /**
     * 如何使用这个注解必须进行设置该值，
     * 用于表示Excel中头部分的信息
     * @return 单元格的标题
     */
    String[] name();

    /**
     * 用于表示是否在Excel中显示该列
     * @return 单元格是否可见（针对一列）
     */
    boolean visible() default true;

    /**
     * 用于表示一列中单元格输出的格式。
     * 例如：yyyy-MM-dd HH:mm:ss
     * 又或者：##0.000等等
     * @return 单元格的格式
     */
    String format() default "";

    /**
     * 表示列的宽度
     * @return 单元格的宽度
     */
    int columnWidth() default 20;
}
