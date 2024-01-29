package cn.viewshine.cloudthree.excel.converter.primitive;

import cn.viewshine.cloudthree.excel.converter.Converter;


/**
 * @author ChangWei[changwei@viewshine.cn]
 */
public class ShortConverter implements Converter<Short> {
    @Override
    public String converter(Short number) {
        if (number == null) {
            return "";
        }
        return number.toString();
    }
}
