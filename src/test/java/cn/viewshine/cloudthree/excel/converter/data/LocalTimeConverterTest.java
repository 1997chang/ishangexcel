package cn.viewshine.cloudthree.excel.converter.data;


import org.junit.Assert;
import org.junit.Test;

public class LocalTimeConverterTest {
    @Test
    public void convertDefaultTest() {
        LocalTimeConverter localTimeConverter = new LocalTimeConverter();
        String convertString = localTimeConverter.converter(null);
        Assert.assertEquals("", convertString);
    }
}