package cn.viewshine.cloudthree.excel.vo;

import cn.viewshine.cloudthree.excel.annotation.ExcelField;
import org.junit.Test;

import javax.xml.bind.ValidationEvent;
import java.lang.reflect.*;
import java.math.BigDecimal;
import java.time.LocalDateTime;
import java.util.List;
import java.util.Map;
import java.util.Objects;

/**
 * @author: 常伟
 * @create: 2019/8/13 23:21
 * @email: kmustchang@qq.com
 * @version: 1.0
 * @Description:
 */
public class WriteModelVo extends FatherVo {

    @ExcelField(name = "姓名")
    private String name;

    @ExcelField(name = "性别")
    private Sex sex = Sex.MAN;

    @ExcelField(name = "年龄")
    private int age;

    @ExcelField(name = "年份")
    private Integer year;

    @ExcelField(name = "金钱",columnWidth = 40)
    private BigDecimal money;

    @ExcelField(name = "单价",format = "###0.00")
    private BigDecimal price;

    @ExcelField(name = "是否团员")
    private Boolean tuanyuan;

    @ExcelField(name = "日期",format = "yyyy-MM-dd")
    private LocalDateTime localDateTime;

    public WriteModelVo(String id,String name, Sex sex, int age, Integer year, BigDecimal money, BigDecimal price,
                         Boolean tuanyuan, LocalDateTime localDateTime) {
        super(id);
        this.name = name;
        this.sex = sex;
        this.age = age;
        this.year = year;
        this.money = money;
        this.price = price;
        this.tuanyuan = tuanyuan;
        this.localDateTime = localDateTime;
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
}
