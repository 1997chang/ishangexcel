package cn.viewshine.cloudthree.excel.fieldutils;

import cn.viewshine.cloudthree.excel.utils.FieldUtils;
import cn.viewshine.cloudthree.excel.vo.WriteModelVo;
import org.apache.poi.ss.usermodel.CellType;
import org.junit.Assert;
import org.junit.Before;
import org.junit.Test;

import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.lang.reflect.TypeVariable;
import java.util.Arrays;

/**
 * @author: 常伟
 * @create: 2019/8/13 23:32
 * @email: kmustchang@qq.com
 * @version: 1.0
 * @Description:
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
