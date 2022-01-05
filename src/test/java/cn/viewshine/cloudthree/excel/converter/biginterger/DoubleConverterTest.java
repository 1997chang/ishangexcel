package cn.viewshine.cloudthree.excel.converter.biginterger;

import org.junit.Assert;
import org.junit.Test;

/**
 * @author moxiao
 * @Email changwei@viewshine.cn
 * @date 2022/1/5
 */
public class DoubleConverterTest {

    @Test
    public void nullConvertTest() {
        DoubleConverter doubleConverter = new DoubleConverter();
        String convertString = doubleConverter.converter(null);
        Assert.assertEquals("", convertString);
    }

    @Test
    public void normalConvertTest() {
        DoubleConverter doubleConverter = new DoubleConverter();
        String convertString = doubleConverter.converter(5.5);
        Assert.assertEquals("5.5", convertString);
    }
}