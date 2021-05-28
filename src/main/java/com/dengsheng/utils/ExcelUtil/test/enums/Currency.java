package com.dengsheng.utils.ExcelUtil.test.enums;

import com.dengsheng.utils.ExcelUtil.interfaces.ExcelFieldEunm;
import lombok.Getter;

public enum Currency implements ExcelFieldEunm {
    RMB(1, "人命币"),
    DOLLAR(2, "美元");

    @Getter
    private Integer key;
    @Getter
    private String value;

    Currency(int key, String value) {
        this.key = key;
        this.value = value;
    }
}
