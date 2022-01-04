package cn.viewshine.cloudthree.excel.converter.biginterger;

import cn.viewshine.cloudthree.excel.converter.Converter;

import java.math.BigDecimal;
import java.math.RoundingMode;

/**
 * @author moxiao
 * @Email changwei@viewshine.cn
 * @date 2021/12/29
 */
public class BigDecimalConverter implements Converter<BigDecimal> {

    /**
     * 表示小数点截取的位数
     */
    private int scale = 2;

    /**
     * 截取的方式，默认截断的方式
     */
    private RoundingMode mode = RoundingMode.DOWN;

    @Override
    public String converter(BigDecimal bigDecimal) {
        if (bigDecimal == null) {
            return "";
        }
        return bigDecimal.setScale(scale, mode).toPlainString();
    }

    public void setScale(int scale) {
        this.scale = scale;
    }

    public void setMode(RoundingMode mode) {
        this.mode = mode;
    }
}
