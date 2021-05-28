package com.dengsheng.utils.ExcelUtil.test.model;

import com.dengsheng.utils.ExcelUtil.interfaces.ExcelClass;
import com.dengsheng.utils.ExcelUtil.interfaces.ExcelConfig;
import com.dengsheng.utils.ExcelUtil.model.Relations;

@ExcelClass(value="User", sheetName="用户")
public class User extends Relations {
    @ExcelConfig(value="序号", index =0)
    private Integer id;
    @ExcelConfig(value="姓名", index =1)
    private String username;
    @ExcelConfig(value="头部", index =2)
    private String head;
    @ExcelConfig(value="性别", index =3)
    private Integer sex;
    @ExcelConfig(value="手机号", index =4)
    private String phone;

    private String adress;

    public Integer getId() {
        return id;
    }

    public void setId(Integer id) {
        this.id = id;
    }

    public String getUsername() {
        return username;
    }

    public void setUsername(String username) {
        this.username = username;
    }

    public String getHead() {
        return head;
    }

    public void setHead(String head) {
        this.head = head;
    }

    public Integer getSex() {
        return sex;
    }

    public void setSex(Integer sex) {
        this.sex = sex;
    }

    public String getPhone() {
        return phone;
    }

    public void setPhone(String phone) {
        this.phone = phone;
    }

    public String getAdress() {
        return adress;
    }

    public void setAdress(String adress) {
        this.adress = adress;
    }
}
