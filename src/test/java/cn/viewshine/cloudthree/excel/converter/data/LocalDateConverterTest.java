package cn.viewshine.cloudthree.excel.converter.data;


import org.junit.Assert;
import org.junit.Test;

public class LocalDateConverterTest {
    
    @Test
    public void convertDefaultTest() {
        LocalDateConverter localDateConverter = new LocalDateConverter();
        String convertString = localDateConverter.converter(null);
        Assert.assertEquals("", convertString);
    }
}