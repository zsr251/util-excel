package com.javasoso.util.excel;

import java.util.Date;

/**
 * 测试model
 *
 * @author jasonzhu
 * @date 2018/4/14
 */
public class TestModel {
    @ExcelIn(0)
    @ExcelOut(value = 0,name = "姓名")
    private String realName;
    @ExcelIn(1)
    @ExcelOut(value = 1,name = "生日")
    private Date birD;
    @ExcelOut(value = 2,name = "生日年月日",dateFormat = "yyyy年MM月dd日")
    private Date bD;
    @ExcelIn(3)
    @ExcelOut(value = 3,name = "年龄")
    private Integer age;
    @ExcelIn(4)
    private String remark;

    public String getRealName() {
        return realName;
    }

    public void setRealName(String realName) {
        this.realName = realName;
    }

    public Date getBirD() {
        return birD;
    }

    public void setBirD(Date birD) {
        this.birD = birD;
    }

    public Date getbD() {
        return bD;
    }

    public void setbD(Date bD) {
        this.bD = bD;
    }

    public Integer getAge() {
        return age;
    }

    public void setAge(Integer age) {
        this.age = age;
    }

    public String getRemark() {
        return remark;
    }

    public void setRemark(String remark) {
        this.remark = remark;
    }

    @Override
    public String toString() {
        return "TestModel{" +
            "realName='" + realName + '\'' +
            ", birD=" + birD +
            ", bD=" + bD +
            ", age=" + age +
            ", remark='" + remark + '\'' +
            '}';
    }
}
