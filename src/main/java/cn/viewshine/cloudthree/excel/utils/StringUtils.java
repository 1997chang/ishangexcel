package cn.viewshine.cloudthree.excel.utils;

/**
 * @author ChangWei[changwei@viewshine.cn]
 */
public final class StringUtils {

    private StringUtils(){}

    public static boolean isBlank(String st) {
        return st == null || st.trim().isEmpty();
    }

}
