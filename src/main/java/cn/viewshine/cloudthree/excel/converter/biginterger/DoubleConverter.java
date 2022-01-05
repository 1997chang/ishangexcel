package cn.viewshine.cloudthree.excel.converter.biginterger;

import cn.viewshine.cloudthree.excel.converter.Converter;

/**
 * @author moxiao
 * @Email changwei@viewshine.cn
 * @date 2022/1/5
 */
public class DoubleConverter implements Converter<Double> {

    @Override
    public String converter(Double value) {
        if (value == null || value.isNaN()) {
            return "";
        }
        return value.toString();
    }
}
