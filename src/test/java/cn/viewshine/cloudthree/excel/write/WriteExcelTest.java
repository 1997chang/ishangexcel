package cn.viewshine.cloudthree.excel.write;

import cn.viewshine.cloudthree.excel.ExcelFactory;
import cn.viewshine.cloudthree.excel.vo.Sex;
import cn.viewshine.cloudthree.excel.vo.Student;
import cn.viewshine.cloudthree.excel.vo.Teacher;
import cn.viewshine.cloudthree.excel.vo.WriteModelVo;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

import java.math.BigDecimal;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.*;
import java.util.concurrent.ThreadLocalRandom;
import java.util.stream.IntStream;

/**
 * @author changwei[changwei@viewshine.cn]
 */
public class WriteExcelTest {

    public void init(List<WriteModelVo> data) {
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

    /**
     * 单Sheet也写入
     */
    @Test
    public void wirteExcel() {
        List<WriteModelVo> data = new ArrayList();
        init(data);
        long l = System.currentTimeMillis();
        ExcelFactory.writeExcel(data,"D:/test1.xlsx");
        System.out.println(System.currentTimeMillis()-l);
    }

    @Test
    public void conditionTest(){
        Workbook workbook = new XSSFWorkbook(); // or new XSSFWorkbook();
        Sheet sheet = workbook.createSheet();

        SheetConditionalFormatting sheetCF = sheet.getSheetConditionalFormatting();

        ConditionalFormattingRule rule1 = sheetCF.createConditionalFormattingRule(ComparisonOperator.EQUAL, "0");
        FontFormatting fontFmt = rule1.createFontFormatting();
        fontFmt.setFontStyle(true, false);
        fontFmt.setFontColorIndex(IndexedColors.DARK_RED.index);

        BorderFormatting bordFmt = rule1.createBorderFormatting();
        bordFmt.setBorderBottom(BorderStyle.THIN);
        bordFmt.setBorderTop(BorderStyle.THICK);
        bordFmt.setBorderLeft(BorderStyle.DASHED);
        bordFmt.setBorderRight(BorderStyle.DOTTED);

        PatternFormatting patternFmt = rule1.createPatternFormatting();
        patternFmt.setFillBackgroundColor(IndexedColors.YELLOW.index);

        ConditionalFormattingRule rule2 = sheetCF.createConditionalFormattingRule(ComparisonOperator.BETWEEN, "-10", "10");
        ConditionalFormattingRule [] cfRules =
                {
                        rule1, rule2
                };

        CellRangeAddress[] regions = {
                CellRangeAddress.valueOf("A3:A5")
        };
        sheetCF.addConditionalFormatting(regions, cfRules);
    }

    /**
     * 多Sheet也写入
     */
    @Test
    public void writeMultiSheetTest(){
        List<Student> studentList = new ArrayList<>();
        loadStudent(studentList);

        List<Teacher> teacherList = new ArrayList<>();
        loadTeacher(teacherList);

        Map<String,List> sheetData = new HashMap<>();
        sheetData.put("学生", studentList);
        sheetData.put("教师", teacherList);
        ExcelFactory.writeExcel(sheetData, "D:/multiSheet.xlsx");
    }

    private void loadTeacher(List<Teacher> teacherList) {
        LocalDate current = LocalDate.now();
        IntStream.range(1,1000).forEach(i ->{
            Teacher teacher = new Teacher();
            teacher.setId(10000L+i);
            teacher.setHireDate(current.plusDays(i));
            teacher.setName("教师No."+i);
            teacherList.add(teacher);
        });
    }

    private void loadStudent(List<Student> studentList) {
        IntStream.range(1,1000).forEach(i ->{
            Student student = new Student();
            student.setId(10000L+i);
            student.setChineseScore(ThreadLocalRandom.current().nextInt());
            student.setMathematicScore(ThreadLocalRandom.current().nextInt());
            student.setName("学生No."+i);
            studentList.add(student);
        });
    }

}
