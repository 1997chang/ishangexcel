package cn.viewshine.cloudthree.excel.utils;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;

import java.util.ArrayList;
import java.util.List;
import java.util.Objects;

/**
 * 表示单元格合并范围
 * @author changwei[changwei@viewshine.cn]
 * @version 1.0
 */
public class CellRangeUtils {

    private CellRangeUtils() {}

    /**
     * 合并单元格
     * @param startRow 当前Sheet开始行
     * @param rowMaxCount 列标题的最大行数
     * @param headList 列标题
     */
    public static void mergeCell(int startRow, int rowMaxCount, List<List<String>> headList, Sheet sheet) {
        List<CellRangeAddress> cellRangeList = getCellRangeList(startRow, rowMaxCount, headList);
        cellRangeList.forEach(sheet::addMergedRegion);
    }

    /**
     * 获取所有可以合并的单元格的区域
     * @param startRow 表示开始行
     * @param rowMaxCount 最大的行数
     * @param headList Head列列表
     * @return
     */
    public static List<CellRangeAddress> getCellRangeList(int startRow, int rowMaxCount, List<List<String>> headList) {
        List<CellRangeAddress> result = new ArrayList<>();
        //获取对应每一行的数据内容
        List<List<String>> headRowList = getRowList(headList, rowMaxCount);

        //遍历所有列
        for (int c = 0; c < headList.size(); c++) {
            for (int r = 0; r < rowMaxCount; r++ ){
                //得到最后一个相等的行以及列的下标
                int lastRow = getLastEquals(r, headList.get(c));
                int lastColumn = getLastEquals(c, headRowList.get(lastRow));
                //lastRow > 0 && lastColumn > 0 表示以前没有合并过，
                boolean canMerge = (lastColumn > c || lastRow > r) && lastRow >= 0 && lastColumn >= 0;
                if (canMerge) {
                    result.add(new CellRangeAddress(startRow + r, startRow + lastRow, c, lastColumn));
                }
                r = lastRow;
            }
        }
        return result;
    }

    /**
     *  用于返回list列表中从startIndex开始。等于startIndex下标值的最后一个下标
     * @param startIndex 从那个下标开始
     * @param list 列表
     * @return 最后一个等于startIndex下标值的下标
     */
    private static int getLastEquals(int startIndex, List<String> list) {
        //startIndex下标对应的值
        String value = list.get(startIndex);
        //如果startIndex大于0，并且与前一个相同的话，表示以前遍历过，直接返回-1。进行跳过
        if (startIndex > 0 && Objects.equals(value,list.get(startIndex-1))) {
            return -1;
        }

        int result = startIndex + 1;
        for (; result < list.size(); result++) {
            if (!Objects.equals(value, list.get(result))) {
                break;
            }
        }
        return result - 1;
    }

    /**
     * 用于将列数据转化为行数据
     * @param headColumnList Head的列数据
     * @param maxRowCount 最大行数
     * @return 返回行数据
     */
    private static List<List<String>> getRowList(List<List<String>> headColumnList, int maxRowCount) {
        List<List<String>> result = new ArrayList<>(maxRowCount);
        int columnCount = headColumnList.size();
        for (int r = 0; r < maxRowCount; r++) {
            List<String> rowData = new ArrayList<>(columnCount);
            for (List<String> stringList : headColumnList) {
                if (r >= stringList.size()) {
                    stringList.add(stringList.get(r - 1));
                }
                rowData.add(stringList.get(r));
            }
            result.add(rowData);
        }
        return result;
    }
}
