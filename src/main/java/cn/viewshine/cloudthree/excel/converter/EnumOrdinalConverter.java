package cn.viewshine.cloudthree.excel.converter;

public class EnumOrdinalConverter implements Converter<Enum<?>> {

    /**
     * 将Enumerate转化为对应的数字
     * @param anEnum 转化的数据
     * @return
     */
    @Override
    public String converter(Enum<?> anEnum) {
        if (anEnum == null) {
            return "";
        }
        return anEnum.ordinal() + "";
    }
}
