package cn.viewshine.cloudthree.excel.fieldutils;

import cn.viewshine.cloudthree.excel.utils.FieldUtils;
import cn.viewshine.cloudthree.excel.vo.WriteModelVo;
import org.apache.poi.ss.usermodel.CellType;
import org.junit.Assert;
import org.junit.Test;
import java.lang.reflect.Field;

/**
 * 单元格样式测试类
 * @author changwei[changwei@viewshine.cn]
 */
public class CellTypeTest {

    /**
     * 用于验证FieldUtils的工具，根据Field计算出CellType对应的类型
     */
    @Test
    public void testField(){
        Field[] declaredFields = WriteModelVo.class.getDeclaredFields();
        Assert.assertEquals(CellType.STRING, FieldUtils.getCellTypeByField(declaredFields[0]));
        Assert.assertEquals(CellType.STRING, FieldUtils.getCellTypeByField(declaredFields[1]));
        Assert.assertEquals(CellType.NUMERIC, FieldUtils.getCellTypeByField(declaredFields[2]));
        Assert.assertEquals(CellType.NUMERIC, FieldUtils.getCellTypeByField(declaredFields[3]));
        Assert.assertEquals(CellType.NUMERIC, FieldUtils.getCellTypeByField(declaredFields[4]));
        Assert.assertEquals(CellType.NUMERIC, FieldUtils.getCellTypeByField(declaredFields[5]));
        Assert.assertEquals(CellType.BOOLEAN, FieldUtils.getCellTypeByField(declaredFields[6]));
        Assert.assertEquals(CellType.NUMERIC, FieldUtils.getCellTypeByField(declaredFields[7]));
    }
}
