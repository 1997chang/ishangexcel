package cn.viewshine.cloudthree.excel.converter;


import cn.viewshine.cloudthree.excel.converter.enums.Sex;
import org.junit.Assert;
import org.junit.Test;

public class EnumNameConverterTest {
    @Test
    public void nullConvertTest() {
        EnumNameConverter enumNameConverter = new EnumNameConverter();
        String converter = enumNameConverter.converter(null);
        Assert.assertEquals("", converter);
    }

    @Test
    public void normalConvertTest() {
        EnumNameConverter enumNameConverter = new EnumNameConverter();
        String converter = enumNameConverter.converter(Sex.MAN);
        Assert.assertEquals("MAN", converter);
    }

    @Test
    public void womanConvertTest() {
        EnumNameConverter enumNameConverter = new EnumNameConverter();
        String converter = enumNameConverter.converter(Sex.WOMAN);
        Assert.assertEquals("WOMAN", converter);
    }
}