package cn.viewshine.cloudthree.excel.exception;

import cn.viewshine.cloudthree.excel.converter.Converter;

/**
 * 数据转化错误
 * @author ChangWei[changwei@viewshine.cn]
 */
public class DataConvertException extends RuntimeException {

    /**
     * 需要转化的数据
     */
    private Object data;

    private Converter<?> converter;

    public DataConvertException(Object data, Converter<?> converter, String message) {
        super(message);
        this.data = data;
        this.converter = converter;
    }

    public DataConvertException(Object data, Converter<?> converter, String message, Throwable cause) {
        super(message, cause);
        this.data = data;
        this.converter = converter;
    }

    public Object getData() {
        return data;
    }

    public void setData(Object data) {
        this.data = data;
    }

    public Converter<?> getConverter() {
        return converter;
    }

    public void setConverter(Converter<?> converter) {
        this.converter = converter;
    }
}
