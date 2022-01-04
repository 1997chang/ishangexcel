package cn.viewshine.cloudthree.excel.converter.biginterger;

import org.junit.Assert;
import org.junit.Test;

import java.math.BigDecimal;

import static org.junit.Assert.*;

/**
 * @author moxiao
 * @Email changwei@viewshine.cn
 * @date 2022/1/4
 */
public class BigDecimalConverterTest {

    @Test
    public void convertDefaultTest() {
        BigDecimalConverter bigDecimalConverter = new BigDecimalConverter();
        String convertString = bigDecimalConverter.converter(new BigDecimal("5.322"));
        Assert.assertEquals("5.32", convertString);
    }

    @Test
    public void convertCustomTest() {
        BigDecimalConverter bigDecimalConverter = new BigDecimalConverter();
        bigDecimalConverter.setScale(3);
        String convertString = bigDecimalConverter.converter(new BigDecimal("5.322"));
        Assert.assertEquals("5.322", convertString);
    }

    @Test
    public void nullConvertTest() {
        BigDecimalConverter bigDecimalConverter = new BigDecimalConverter();
        bigDecimalConverter.setScale(3);
        String convertString = bigDecimalConverter.converter(null);
        Assert.assertEquals("", convertString);
    }

}