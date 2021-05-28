package com.dengsheng.utils.ExcelUtil.exceptions;

public class ExcelException extends RuntimeException{
    //异常信息
    private String message;

    private ExcelException(){}

    ExcelException(String msg){
        this.message = msg;
    }

    @Override
    public String getMessage() {
        return message;
    }

}


