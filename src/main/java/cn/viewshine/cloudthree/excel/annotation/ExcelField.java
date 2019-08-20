package cn.viewshine.cloudthree.excel.annotation;


import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * @Auhtor: Changwei
 * @Email: changwei@viewshine.cn
 * @Date: 2019/8/3
 * @Description: 用于导入/导出Excel表格中Header头部分的信息
 */
@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.FIELD)
public @interface ExcelField {

    /**
     * 如何使用这个注解必须进行设置该值，
     * 用于表示Excel中头部分的信息
     * @return
     */
    String[] name();

    /**
     * 用于表示是否在Excel中显示该列
     * @return
     */
    boolean visible() default true;

    /**
     * 用于表示一列中单元格输出的格式。
     * 例如：yyyy-MM-dd HH:mm:ss
     * 又或者：##0.000等等
     * @return
     */
    String format() default "";

    /**
     * 表示列的宽度
     * @return
     */
    int columnWidth() default 20;
}
