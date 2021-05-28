package com.dengsheng.utils.ExcelUtil.model;

import com.dengsheng.utils.ExcelUtil.interfaces.ExcelClass;
import org.apache.commons.lang3.StringUtils;

import java.util.List;

public class Relation<T>{
    /**
     * 参数的数据类型
     */
    private Class clazz ;
    /**
     * 参数数据
     */
    private List<T> data;
    /**
     * 是否需要字段说明
     */
    private Boolean needHead;
    /**
     * 参数类型，用于区分类别
     */
    private String relationTypeName;

    public Relation(){};

    public Relation(Class<T> clazz, List<T> data, Boolean needHead, String relationTypeName) {
        this.clazz = clazz;
        this.data = data;
        this.needHead = needHead;
        if(StringUtils.isEmpty(relationTypeName)){
            if(clazz.isAnnotationPresent(ExcelClass.class)){
                this.relationTypeName = clazz.getAnnotation(ExcelClass.class).value();
            }
        } else {
            this.relationTypeName = relationTypeName;
        }

    }

    public Class getClazz() {
        return clazz;
    }

    public void setClazz(Class clazz) {
        this.clazz = clazz;
    }

    public List<T> getData() {
        return data;
    }

    public void setData(List<T> data) {
        this.data = data;
    }

    public Boolean getNeedHead() {
        return needHead;
    }

    public void setNeedHead(Boolean needHead) {
        this.needHead = needHead;
    }

    public String getRelationTypeName() {
        return relationTypeName;
    }

    public void setRelationTypeName(String relationTypeName) {
        this.relationTypeName = relationTypeName;
    }
}
