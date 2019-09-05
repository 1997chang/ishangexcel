package cn.viewshine.cloudthree.excel.vo;

import cn.viewshine.cloudthree.excel.annotation.ExcelField;
import java.math.BigDecimal;
import java.time.LocalDateTime;
import java.util.Date;

/**
 * @author changwei[changwei@viewshine.cn]
 */
public class WriteModelVo extends FatherVo {

    @ExcelField(name = {"威星表格","姓名"})
    private String name;

    @ExcelField(name = {"威星表格","性别"})
    private Sex sex = Sex.MAN;

    @ExcelField(name = {"威星表格","年龄"})
    private int age;

    @ExcelField(name = {"威星表格","年份"})
    private Integer year;

    @ExcelField(name = {"威星表格","金钱"},columnWidth = 40)
    private BigDecimal money;

    @ExcelField(name = {"威星表格","单价"},format = "###0.00")
    private BigDecimal price;

    @ExcelField(name = {"威星表格","是否团员"})
    private Boolean tuanyuan;

    @ExcelField(name = {"威星表格","出生日期"})
    private LocalDateTime localDateTime;

    @ExcelField(name = {"威星表格","注册时间"},format = "yyyy-MM-dd")
    private Date date;

    public WriteModelVo(String id,String name, Sex sex, int age, Integer year, BigDecimal money, BigDecimal price,
                         Boolean tuanyuan, LocalDateTime localDateTime, Date date) {
        super(id);
        this.name = name;
        this.sex = sex;
        this.age = age;
        this.year = year;
        this.money = money;
        this.price = price;
        this.tuanyuan = tuanyuan;
        this.localDateTime = localDateTime;
        this.date = date;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public Sex getSex() {
        return sex;
    }

    public void setSex(Sex sex) {
        this.sex = sex;
    }

    public int getAge() {
        return age;
    }

    public void setAge(int age) {
        this.age = age;
    }

    public Integer getYear() {
        return year;
    }

    public void setYear(Integer year) {
        this.year = year;
    }

    public BigDecimal getMoney() {
        return money;
    }

    public void setMoney(BigDecimal money) {
        this.money = money;
    }

    public BigDecimal getPrice() {
        return price;
    }

    public void setPrice(BigDecimal price) {
        this.price = price;
    }

    public Boolean getTuanyuan() {
        return tuanyuan;
    }

    public void setTuanyuan(Boolean tuanyuan) {
        this.tuanyuan = tuanyuan;
    }

    public LocalDateTime getLocalDateTime() {
        return localDateTime;
    }

    public void setLocalDateTime(LocalDateTime localDateTime) {
        this.localDateTime = localDateTime;
    }

    public Date getDate() {
        return date;
    }

    public void setDate(Date date) {
        this.date = date;
    }
}
