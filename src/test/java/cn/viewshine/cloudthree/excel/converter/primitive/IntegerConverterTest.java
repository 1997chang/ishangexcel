package cn.viewshine.cloudthree.excel.converter.primitive;

import org.junit.Assert;
import org.junit.Test;

public class IntegerConverterTest {

    @Test
    public void convertShortTest() {
        IntegerConverter integerConverter = new IntegerConverter();
        String convertString = integerConverter.converter(new Integer("123"));
        Assert.assertEquals("123", convertString);
    }

    @Test
    public void convertCustomTest() {
        IntegerConverter integerConverter = new IntegerConverter();
        int number = 7;
        String convertString = integerConverter.converter(number);
        Assert.assertEquals("7", convertString);
    }

    @Test
    public void nullConvertShortTest() {
        IntegerConverter integerConverter = new IntegerConverter();
        String convertString = integerConverter.converter(null);
        Assert.assertEquals("", convertString);
    }
}