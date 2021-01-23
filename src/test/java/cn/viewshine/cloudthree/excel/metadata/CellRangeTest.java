package cn.viewshine.cloudthree.excel.metadata;

import cn.viewshine.cloudthree.excel.utils.CellRangeUtils;
import org.apache.poi.ss.util.CellRangeAddress;
import org.junit.Assert;
import org.junit.Ignore;
import org.junit.Test;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

/**
 * 单元格合并测试类
 * @author  changwei[changwei@viewshine.cn]
 */
@Ignore
public class CellRangeTest {

    /**
     * 用于测试合并单元格
     */
    @Test
    public void mergeCellTest() {
        List<List<String>> head = new ArrayList<List<String>>();
        head.add(Arrays.asList("实验一班成绩表,学号".split(",")));
        head.add(Arrays.asList("实验一班成绩表,姓名".split(",")));
        head.add(Arrays.asList("实验一班成绩表,语文".split(",")));
        head.add(Arrays.asList("实验一班成绩表,数学".split(",")));
        head.add(Arrays.asList("实验一班成绩表,英语".split(",")));
        List<CellRangeAddress> cellRangeList = CellRangeUtils.getCellRangeList(0, 2, head);
        Assert.assertEquals(cellRangeList.size(), 1);
    }

}
