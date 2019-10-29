package cn.viewshine.cloudthree.excel.write;

import cn.viewshine.cloudthree.excel.ExcelFactory;
import cn.viewshine.cloudthree.excel.vo.Sex;
import cn.viewshine.cloudthree.excel.vo.WriteModelVo;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;
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
            data.add(new WriteModelVo(i+"",null,sex,20+i,1999-i,money,price,null,LocalDateTime.now(),new Date()));
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

    @Test
    public void conditionTest() throws IOException {
        Workbook wb = new XSSFWorkbook();
        sameCell(wb.createSheet("Same Cell"));
        String file = "cf-poi.xlsx";
        FileOutputStream out = new FileOutputStream(file);
        wb.write(out);
        out.close();
        System.out.println("Generated: " + file);
        wb.close();
    }

    /**
     * Highlight cells based on their values
     */
    void sameCell(Sheet sheet) {
        for (int i=0; i<40; i++) {
            int rn = i+1;
            Row r = sheet.createRow(i);
            r.createCell(0).setCellValue("This is row " + rn + " (" + i + ")");
            String str = "";
            if (rn%2 == 0) {
                str = str + "even ";
            }
            if (rn%3 == 0) {
                str = str + "x3 ";
            }
            if (rn%5 == 0) {
                str = str + "x5 ";
            }
            if (rn%10 == 0) {
                str = str + "x10 ";
            }
            if (str.length() == 0) {
                str = "nothing special...";
            }
            r.createCell(1).setCellValue("It is " + str);
        }
        sheet.autoSizeColumn(0);
        sheet.autoSizeColumn(1);

        sheet.getRow(1).createCell(3).setCellValue("Even rows are blue");
        sheet.getRow(2).createCell(3).setCellValue("Multiples of 3 have a grey background");
        sheet.getRow(4).createCell(3).setCellValue("Multiples of 5 are bold");
        sheet.getRow(9).createCell(3).setCellValue("Multiples of 10 are red (beats even)");

        SheetConditionalFormatting sheetCF = sheet.getSheetConditionalFormatting();

        // Condition 1: Row divides by 10, red (will beat #1)
        ConditionalFormattingRule rule1 =
                sheetCF.createConditionalFormattingRule("MOD(ROW(),10)=0");
        FontFormatting font1 = rule1.createFontFormatting();
        font1.setFontColorIndex(IndexedColors.RED.index);

        // Condition 2: Row is even, blue
        ConditionalFormattingRule rule2 =
                sheetCF.createConditionalFormattingRule("MOD(ROW(),2)=0");
        FontFormatting font2 = rule2.createFontFormatting();
        font2.setFontColorIndex(IndexedColors.BLUE.index);

        // Condition 3: Row divides by 5, bold
        ConditionalFormattingRule rule3 =
                sheetCF.createConditionalFormattingRule("MOD(ROW(),5)=0");
        FontFormatting font3 = rule3.createFontFormatting();
        font3.setFontStyle(false, true);

        // Condition 4: Row divides by 3, grey background
        ConditionalFormattingRule rule4 =
                sheetCF.createConditionalFormattingRule("MOD(ROW(),3)=0");
        PatternFormatting fill4 = rule4.createPatternFormatting();
        fill4.setFillBackgroundColor(IndexedColors.GREY_25_PERCENT.index);
        fill4.setFillPattern(PatternFormatting.SOLID_FOREGROUND);

        // Apply
        CellRangeAddress[] regions = {
                CellRangeAddress.valueOf("A1:F41")
        };

        ConditionalFormattingRule [] cfRules =
                {
                        rule1, rule2
                };

        sheetCF.addConditionalFormatting(regions,cfRules);
//        sheetCF.addConditionalFormatting(regions, rule2);
//        sheetCF.addConditionalFormatting(regions, rule1);
//        sheetCF.addConditionalFormatting(regions, rule3);
//        sheetCF.addConditionalFormatting(regions, rule4);
    }

    @Test
    public void customCOlor() throws IOException {
        XSSFWorkbook wb = new XSSFWorkbook();
        XSSFSheet sheet = wb.createSheet();
        XSSFRow row = sheet.createRow(0);
        XSSFCell cell = row.createCell( 0);
        cell.setCellValue("custom XSSF colors");

        XSSFCellStyle style1 = wb.createCellStyle();
        style1.setFillForegroundColor(new XSSFColor(new java.awt.Color(255, 0, 0), new DefaultIndexedColorMap()));
        style1.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        String file = "cf-poi1.xlsx";
        FileOutputStream out = new FileOutputStream(file);
        wb.write(out);
        out.close();
        System.out.println("Generated: " + file);
        wb.close();
    }
}
