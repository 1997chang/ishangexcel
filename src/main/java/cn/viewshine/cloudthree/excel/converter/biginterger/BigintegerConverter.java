package cn.viewshine.cloudthree.excel.converter.biginterger;

import cn.viewshine.cloudthree.excel.converter.Converter;

import java.math.BigInteger;

/**
 * @author ChangWei[changwei@viewshine.cn]
 */
public class BigintegerConverter implements Converter<BigInteger> {

    @Override
    public String converter(BigInteger bigInteger) {
        if (bigInteger == null) {
            return "";
        }
        return bigInteger.toString();
    }
}
