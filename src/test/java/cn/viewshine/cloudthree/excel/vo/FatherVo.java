package cn.viewshine.cloudthree.excel.vo;

import cn.viewshine.cloudthree.excel.annotation.ExcelField;

/**
 * @Description:
 * @Author: ChangWei
 * @Email: changwei@viewshine.cn
 * @Date: 2019/8/20
 */
public class FatherVo {
    @ExcelField(name = "证件号")
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
