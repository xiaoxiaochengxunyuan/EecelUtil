package com.dengsheng.utils.ExcelUtil.test.model;


import com.dengsheng.utils.ExcelUtil.interfaces.ExcelClass;
import com.dengsheng.utils.ExcelUtil.interfaces.ExcelConfig;
import com.dengsheng.utils.ExcelUtil.test.enums.Currency;


@ExcelClass(value="UserWallet", sheetName="钱包")
public class UserWallet{
    @ExcelConfig(value="钱包类型", index =0, enumClass= Currency.class, enumFieldName="type")
    private String typeName;
    @ExcelConfig(value="数量", index =1)
    private Integer num;
    // 钱包类型
    private Integer type;

    public UserWallet(){}

    public UserWallet(Integer type, Integer num) {
        this.type = type;
        this.num = num;
    }

    public String getTypeName() {
        return typeName;
    }

    public void setTypeName(String typeName) {
        this.typeName = typeName;
    }

    public Integer getNum() {
        return num;
    }

    public void setNum(Integer num) {
        this.num = num;
    }

    public Integer getType() {
        return type;
    }

    public void setType(Integer type) {
        this.type = type;
    }
}
