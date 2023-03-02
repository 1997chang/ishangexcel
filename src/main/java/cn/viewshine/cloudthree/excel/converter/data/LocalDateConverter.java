package cn.viewshine.cloudthree.excel.converter.data;

import cn.viewshine.cloudthree.excel.converter.Converter;

import java.time.LocalDate;
import java.time.format.DateTimeFormatter;

public class LocalDateConverter implements Converter<LocalDate> {

    private static final DateTimeFormatter YYYY_MM_DD = DateTimeFormatter.ofPattern("yyyy-MM-dd");
    
    @Override
    public String converter(LocalDate localDate) {
        if (localDate == null) {
            return "";
        }
        return localDate.format(YYYY_MM_DD);
    }
}
