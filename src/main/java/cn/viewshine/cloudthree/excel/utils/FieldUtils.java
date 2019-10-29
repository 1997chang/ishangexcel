package cn.viewshine.cloudthree.excel.utils;

import cn.viewshine.cloudthree.excel.annotation.ExcelField;
import cn.viewshine.cloudthree.excel.metadata.ColumnProperty;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Workbook;

import java.lang.reflect.Field;
import java.lang.reflect.Type;
import java.time.LocalDateTime;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.stream.Collectors;

/**
 * @author changwei[changwei@viewshine.cn]
 * @version 1.0
 */
public class FieldUtils {

    private FieldUtils() { }

    public static CellType getCellTypeByField(Field field) {
        Type genericType = field.getGenericType();
        if (genericType instanceof Class) {
            Class fieldClass = (Class) genericType;
            if (fieldClass.isEnum() || fieldClass.equals(String.class)) {
                return CellType.STRING;
            } else if (fieldClass.equals(Boolean.TYPE) || fieldClass.equals(Boolean.class)) {
                return CellType.BOOLEAN;
            } else {
                return CellType.NUMERIC;
            }
        } else {
            return CellType.STRING;
        }
    }

    /**
     * 获取所有的带有@ExcelField字段的列表，并且为每一个字段设置默认的样式，
     * @param zclass CLass类
     * @param cellStyle 表示默认的样式
     * @return
     */
    public static List<ColumnProperty> getAllColumnPropertyOfSingleClass(Class zclass, CellStyle cellStyle, Workbook workbook){
        List<Field> result = new ArrayList<>();
        while (zclass != null) {
            result.addAll(0, Arrays.asList(zclass.getDeclaredFields()));
            zclass = zclass.getSuperclass();
        }
        //将带有@ExcelField注解，并且visible为true的显示出来，其他都不显示，然后设置默认样式
        return result.stream().filter(FieldUtils::visible).map(field -> new ColumnProperty(field, cellStyle, workbook)).collect(Collectors.toList());
    }

    /**
     * 用于判断这个Field是否显示在Excel中
     * @param field
     * @return 如果显示返回 true，否则返回false
     */
    private static boolean visible(Field field) {
        return field.isAnnotationPresent(ExcelField.class) && field.getDeclaredAnnotation(ExcelField.class).visible();
    }
}
