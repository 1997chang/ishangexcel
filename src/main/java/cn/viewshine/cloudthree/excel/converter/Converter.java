package cn.viewshine.cloudthree.excel.converter;

/**
 * @author ChangWei[changwei@viewshine.cn]
 */
public interface Converter<T> {

    /**
     * 将制定类型转化为字符串
     * @param t 转化的数据
     * @return 转化后的字符串
     */
    String converter(T t);

}
