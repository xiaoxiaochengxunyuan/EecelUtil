package com.dengsheng.utils.ExcelUtil.test.model;

import com.dengsheng.utils.ExcelUtil.interfaces.ExcelClass;
import com.dengsheng.utils.ExcelUtil.interfaces.ExcelConfig;

@ExcelClass(value="UserFamily", sheetName="家庭")
public class UserFamily{
    @ExcelConfig(value="关系", index =0)
    private String relation;
    @ExcelConfig(value="姓名", index =1)
    private String name;
    @ExcelConfig(value="年龄", index =2)
    private Integer age;

    public UserFamily(){}

    public UserFamily(String relation, String name, Integer age) {
        this.relation = relation;
        this.name = name;
        this.age = age;
    }

    public String getRelation() {
        return relation;
    }

    public void setRelation(String relation) {
        this.relation = relation;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public Integer getAge() {
        return age;
    }

    public void setAge(Integer age) {
        this.age = age;
    }
}
