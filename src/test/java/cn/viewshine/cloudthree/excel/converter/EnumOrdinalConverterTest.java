package cn.viewshine.cloudthree.excel.converter;

import cn.viewshine.cloudthree.excel.converter.enums.Sex;
import org.junit.Assert;
import org.junit.Test;

public class EnumOrdinalConverterTest{

    @Test
    public void nullConvertTest() {
        EnumOrdinalConverter enumOrdinalConverter = new EnumOrdinalConverter();
        String converter = enumOrdinalConverter.converter(null);
        Assert.assertEquals("", converter);
    }

    @Test
    public void normalConvertTest() {
        EnumOrdinalConverter enumOrdinalConverter= new EnumOrdinalConverter();
        String converter = enumOrdinalConverter.converter(Sex.MAN);
        Assert.assertEquals("0", converter);
    }

    @Test
    public void womanConvertTest() {
        EnumOrdinalConverter enumOrdinalConverter = new EnumOrdinalConverter();
        String converter = enumOrdinalConverter.converter(Sex.WOMAN);
        Assert.assertEquals("1", converter);
    }

}