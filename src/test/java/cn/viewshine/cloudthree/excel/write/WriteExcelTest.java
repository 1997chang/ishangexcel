package cn.viewshine.cloudthree.excel.write;

import cn.viewshine.cloudthree.excel.ExcelFactory;
import cn.viewshine.cloudthree.excel.vo.Sex;
import cn.viewshine.cloudthree.excel.vo.WriteModelVo;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.ss.util.CellUtil;
import org.apache.poi.xssf.usermodel.XSSFPivotTable;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Before;
import org.junit.Test;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.math.BigDecimal;
import java.time.LocalDateTime;
import java.util.*;
import java.util.concurrent.ThreadLocalRandom;
import java.util.concurrent.TimeUnit;

import static org.apache.poi.ss.SpreadsheetVersion.EXCEL2007;
import static org.apache.poi.ss.usermodel.IndexedColors.GREY_25_PERCENT;

/**
 * @Description:
 * @Author: ChangWei
 * @Email: changwei@viewshine.cn
 * @Date: 2019/8/15
 */
public class WriteExcelTest {

    private List<WriteModelVo> data = new ArrayList();

    @Before
    public void init() {
        for (int i = 0; i < 10000; i++) {
            LocalDateTime now = LocalDateTime.now();
            Sex sex;
            boolean tuanyuan;
            if ((i & 1) == 1) {
                sex = Sex.MAN;
                tuanyuan = true;
            } else {
              sex = Sex.WOMAN;
              tuanyuan = false;
            }

            BigDecimal money = BigDecimal.valueOf(ThreadLocalRandom.current().nextFloat());
            BigDecimal price = BigDecimal.valueOf(ThreadLocalRandom.current().nextDouble());
            data.add(new WriteModelVo(i+"","常"+i,sex,20+i,1999-i,money,price,tuanyuan,now.plusDays(i)));
        }
    }

    @Test
    public void wirteExcel() {
        long l = System.currentTimeMillis();
        Map<String,List> datamap = new LinkedHashMap<>();
        datamap.put("sheet1",data);
        datamap.put("sheet2",data);
        ExcelFactory.writeExcel(datamap,"D:/test1.xlsx");
        System.out.println(System.currentTimeMillis()-l);
    }
}
