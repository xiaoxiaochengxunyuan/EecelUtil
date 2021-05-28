package com.dengsheng.utils.ExcelUtil.model;

import java.util.List;

public class Relations<T>{
    // 关系数据
    private List<Relation<T>> relations;

    public List<Relation<T>> getRelations() {
        return relations;
    }

    public void setRelations(List<Relation<T>> relations) {
        this.relations = relations;
    }
}