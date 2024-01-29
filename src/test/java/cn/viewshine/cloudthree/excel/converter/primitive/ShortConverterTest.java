package cn.viewshine.cloudthree.excel.converter.primitive;

import cn.viewshine.cloudthree.excel.converter.biginterger.BigDecimalConverter;
import org.junit.Assert;
import org.junit.Test;

import java.math.BigDecimal;

public class ShortConverterTest{

    @Test
    public void convertShortTest() {
        ShortConverter shortConverter = new ShortConverter();
        String convertString = shortConverter.converter(new Short("234"));
        Assert.assertEquals("234", convertString);
    }

    @Test
    public void convertCustomTest() {
        ShortConverter shortConverter = new ShortConverter();
        short number = 4;
        String convertString = shortConverter.converter(number);
        Assert.assertEquals("4", convertString);
    }

    @Test
    public void nullConvertShortTest() {
        ShortConverter shortConverter = new ShortConverter();
        String convertString = shortConverter.converter(null);
        Assert.assertEquals("", convertString);
    }

}