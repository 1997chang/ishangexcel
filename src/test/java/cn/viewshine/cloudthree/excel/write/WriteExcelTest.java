package cn.viewshine.cloudthree.excel.write;

import cn.viewshine.cloudthree.excel.ExcelFactory;
import cn.viewshine.cloudthree.excel.annotation.ExcelField;
import cn.viewshine.cloudthree.excel.vo.Sex;
import cn.viewshine.cloudthree.excel.vo.Student;
import cn.viewshine.cloudthree.excel.vo.Teacher;
import cn.viewshine.cloudthree.excel.vo.WriteModelVo;
import net.sf.cglib.beans.BeanMap;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;
import org.junit.Ignore;
import org.junit.Test;

import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.*;
import java.util.concurrent.ThreadLocalRandom;
import java.util.stream.IntStream;

/**
 * @author changwei[changwei@viewshine.cn]
 */
//@Ignore
public class WriteExcelTest {

    public void init(List<WriteModelVo> data) {
        for (int i = 0; i < 1048574; i++) {
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

    private void initMap(List<List<String>> contentData) {
        for (int i = 0; i < 1000; i++) {
            List<String> itemData = new ArrayList<>();
            itemData.add(String.valueOf(i));
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
            itemData.add("姓名" + i);
            itemData.add(sex.name());
            itemData.add(String.valueOf(20 + i));
            itemData.add(String.valueOf(tuanyuan));
            itemData.add(String.valueOf(money.doubleValue()));
            contentData.add(itemData);
        }
    }

    private void initHeadMap(List<List<String>> headMap) {
        List<String> id = Arrays.asList("威星", "ID");
        List<String> name = Arrays.asList("威星", "姓名");
        List<String> sex = Arrays.asList("威星", "性别");
        List<String> age = Arrays.asList("威星", "年龄");
        List<String> tuanyuan = Arrays.asList("威星", "是否团员");
        List<String> money = Arrays.asList("威星", "收入");
        headMap.addAll(Arrays.asList(id, name, sex, age, tuanyuan, money));
    }

    @Test
    public void test() {
        ExcelEntityInAnnotation entity = new ExcelEntityInAnnotation();
        int i = 0;
        entity.setAge((long) i);
        entity.setName("chang "+ i);
        entity.setBirthDate(new Date());
        entity.setPrice(new BigDecimal(i));
        BeanMap beanMap = BeanMap.create(entity);
        System.out.println(beanMap);

    }

    /**
     * 单Sheet也写入
     */
    @Test
    public void wirteExcel() {
        List<WriteModelVo> data = new ArrayList<>();
        init(data);
        long l = System.currentTimeMillis();
        ExcelFactory.writeExcel(data,"D:/test1.xlsx");
        System.out.println(System.currentTimeMillis()-l);
    }

    @Test
    public void createDirecotory() {
        try {
            Files.createDirectories(Paths.get("./ssss/ssssss/chang/aa.txt"));
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    @Test
    public void writeExcelByMap() {
        List<List<String>> contentData = new ArrayList<>(1000);
        initMap(contentData);
        List<List<String>> headName = new ArrayList<>(6);
        initHeadMap(headName);
        ExcelFactory.writeExcel(Collections.singletonMap("sheet1", contentData), Collections.singletonMap("sheet1",
                headName), "sssss/testmap.xlsx");
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

        Map<String,List<?>> sheetData = new HashMap<>();
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


    @Test
    public void conditionTest1() throws IOException {
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

    static class ExcelEntityInAnnotation {

        /**
         * 姓名
         */
        @ExcelField(name = "姓名")
        private String name;

        /**
         * 年龄
         */
        @ExcelField(name = "年龄")
        private Long age;

        /**
         * 出生年月
         */
        @ExcelField(name = "出生年月", format = "yyyy-MM-dd HH:mm:ss")
        private Date birthDate;

        /**
         * 价格
         */
        @ExcelField(name = "价格")
        private BigDecimal price;

        public String getName() {
            return name;
        }

        public void setName(String name) {
            this.name = name;
        }

        public Long getAge() {
            return age;
        }

        public void setAge(Long age) {
            this.age = age;
        }

        public Date getBirthDate() {
            return birthDate;
        }

        public void setBirthDate(Date birthDate) {
            this.birthDate = birthDate;
        }

        public BigDecimal getPrice() {
            return price;
        }

        public void setPrice(BigDecimal price) {
            this.price = price;
        }
    }
}
