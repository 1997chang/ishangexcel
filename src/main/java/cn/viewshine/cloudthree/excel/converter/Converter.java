package cn.viewshine.cloudthree.excel.converter;

/**
 * @author ChangWei[changwei@viewshine.cn]
 */
public interface Converter<T> {

    String converter(T t);

}
