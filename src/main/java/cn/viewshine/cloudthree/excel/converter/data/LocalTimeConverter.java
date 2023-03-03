package cn.viewshine.cloudthree.excel.converter.data;

import cn.viewshine.cloudthree.excel.converter.Converter;

import java.time.LocalTime;
import java.time.format.DateTimeFormatter;

public class LocalTimeConverter implements Converter<LocalTime> {

    private static final DateTimeFormatter HH_MM_SS = DateTimeFormatter.ofPattern("HH:mm:ss");
    
    @Override
    public String converter(LocalTime localDate) {
        if (localDate == null) {
            return "";
        }
        return localDate.format(HH_MM_SS);
    }
}
