package cn.viewshine.cloudthree.excel.vo;

import cn.viewshine.cloudthree.excel.annotation.ExcelField;

/**
 * @author changwei[changwei@viewshine.cn]
 */
public class FatherVo {
    @ExcelField(name = {"威星表格","证件号"})
    private String id;


    public FatherVo(String id) {
        this.id = id;
    }

    public String getId() {
        return id;
    }

    public void setId(String id) {
        this.id = id;
    }
}
