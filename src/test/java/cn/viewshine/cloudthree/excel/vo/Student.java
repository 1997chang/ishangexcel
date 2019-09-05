package cn.viewshine.cloudthree.excel.vo;

import cn.viewshine.cloudthree.excel.annotation.ExcelField;

/**
 * @author: changWei[changwei@viewshine.cn]
 */
public class Student {

    @ExcelField(name = {"第一个Sheet","学生学号"})
    private Long id;

    @ExcelField(name = {"第一个Sheet","学生姓名"})
    private String name;

    @ExcelField(name = {"第二个Sheet","语文成绩"})
    private Integer chineseScore;

    @ExcelField(name = {"第二个Sheet","数学成绩"})
    private Integer  mathematicScore;

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

    public Integer getChineseScore() {
        return chineseScore;
    }

    public void setChineseScore(Integer chineseScore) {
        this.chineseScore = chineseScore;
    }

    public Integer getMathematicScore() {
        return mathematicScore;
    }

    public void setMathematicScore(Integer mathematicScore) {
        this.mathematicScore = mathematicScore;
    }
}
