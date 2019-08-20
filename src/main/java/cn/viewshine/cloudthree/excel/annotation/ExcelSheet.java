package cn.viewshine.cloudthree.excel.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * @Auhtor: Changwei
 * @Email: changwei@viewshine.cn
 * @Date: 2019/8/11
 * @Description: 这个表示写入Sheet的名称
 */
@Target(ElementType.TYPE)
@Retention(RetentionPolicy.RUNTIME)
public @interface ExcelSheet {

    String name() default "";
}
