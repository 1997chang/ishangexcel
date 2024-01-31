package cn.viewshine.cloudthree.excel.converter.primitive;

import cn.viewshine.cloudthree.excel.converter.Converter;


/**
 * @author ChangWei[changwei@viewshine.cn]
 */
public class IntegerConverter implements Converter<Integer> {
    @Override
    public String converter(Integer number) {
        if (number == null) {
            return "";
        }
        return number.toString();
    }
}
