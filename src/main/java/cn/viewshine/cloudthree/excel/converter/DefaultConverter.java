package cn.viewshine.cloudthree.excel.converter;

import java.util.Objects;

/**
 * @author ChangWei[changwei@viewshine.cn]
 */
public class DefaultConverter implements Converter<Object> {

    @Override
    public String converter(Object o) {
        return Objects.toString(o, "");
    }
}
