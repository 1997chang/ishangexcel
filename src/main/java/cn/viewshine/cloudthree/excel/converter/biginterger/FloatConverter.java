package cn.viewshine.cloudthree.excel.converter.biginterger;

import cn.viewshine.cloudthree.excel.converter.Converter;

/**
 * @author ChangWei[changwei@viewshine.cn]
 */
public class FloatConverter implements Converter<Float> {
    @Override
    public String converter(Float value) {
        if (value == null || value.isNaN()) {
            return "";
        }
        return value.toString();
    }
}
