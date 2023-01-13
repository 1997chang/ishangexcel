package cn.viewshine.cloudthree.excel.converter.biginterger;

import org.junit.Assert;
import org.junit.Test;

public class FloatConverterTest{

    @Test
    public void nullConvertTest() {
        FloatConverter floatConverter = new FloatConverter();
        String converter = floatConverter.converter(null);
        Assert.assertEquals("", converter);
    }

    @Test
    public void normalConvertTest() {
        FloatConverter floatConverter= new FloatConverter();
        String converter = floatConverter.converter(0.11f);
        Assert.assertEquals("0.11", converter);
    }

}