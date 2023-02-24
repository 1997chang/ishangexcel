package cn.viewshine.cloudthree.excel.converter;

public class EnumNameConverter implements Converter<Enum<?>> {

    /**
     * 将Enumerate转化为对应的名称
     * @param e 转化的数据
     * @return
     */
    @Override
    public String converter(Enum<?> e) {
        if (e == null) {
            return "";
        }
        return e.name();
    }
}
