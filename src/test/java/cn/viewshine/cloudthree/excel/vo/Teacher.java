package cn.viewshine.cloudthree.excel.vo;

import cn.viewshine.cloudthree.excel.annotation.ExcelField;

import java.time.LocalDate;

/**
 * @author: changWei[changwei@viewshine.cn]
 */
public class Teacher {

    @ExcelField(name = {"第二个Sheet","教师编号"})
    private Long id;

    @ExcelField(name = {"第二个Sheet","教师姓名"})
    private String name;

    @ExcelField(name = {"第二个Sheet","入职时间"},format = "yyyy-MM-dd")
    private LocalDate hireDate;

    public Long getId() {
        return id;
    }

    public void setId(Long id) {
        this.id = id;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public LocalDate getHireDate() {
        return hireDate;
    }

    public void setHireDate(LocalDate hireDate) {
        this.hireDate = hireDate;
    }
}
