package com.dengsheng.utils.ExcelUtil.test;

import com.dengsheng.utils.ExcelUtil.ExcelUtil;
import com.dengsheng.utils.ExcelUtil.model.Relation;
import com.dengsheng.utils.ExcelUtil.test.enums.Currency;
import com.dengsheng.utils.ExcelUtil.test.model.User;
import com.dengsheng.utils.ExcelUtil.test.model.UserFamily;
import com.dengsheng.utils.ExcelUtil.test.model.UserWallet;
import com.dengsheng.utils.FileUtil.FileUtils;

import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

public class Test {

    public static void main(String[] args) throws IOException {
        List<User> userData = createUserData();
        // 导入
        System.out.println(FileUtils.getConTextPath()+"/user.xlsx");
        ExcelUtil.exportExcel(userData, FileUtils.getConTextPath()+"/user.xlsx", User.class, true);

        // 导出
        List<User> importData = ExcelUtil.importExcel(User.class,FileUtils.getConTextPath()+"/user.xlsx", true);
        System.out.println(importData.toString());
    }

    public static List<User> createUserData(){
        List<User> list = new ArrayList<>();
        for (int i = 0; i < 5000; i++) {
            User user1 = new User();
            user1.setId(i);
            user1.setUsername("张三");
            user1.setHead("http://qzapp.qlogo.cn/qzapp/101357640/3C94155CAB4E28517D8435BF404B52F1/100");
            user1.setSex(0);
            user1.setPhone("1880000" + i);
            user1.setAdress("1000001");


            UserFamily father = new UserFamily("父子", user1.getUsername() + "的父亲", i + 20);
            UserFamily mother = new UserFamily("母亲", user1.getUsername() + "的母亲", i + 20);
            List<UserFamily> families = new ArrayList<>();
            families.add(father);
            families.add(mother);
            Relation familyRelation = new Relation<UserFamily>(UserFamily.class, families, false, "UserFamily");

            UserWallet rmb = new UserWallet(Currency.RMB.getKey(), 100 + i);
            UserWallet dollar = new UserWallet(Currency.DOLLAR.getKey(), 100 + i);
            List<UserWallet> wallets = new ArrayList<>();
            wallets.add(rmb);
            wallets.add(dollar);

            Relation walletRelation = new Relation<>(UserWallet.class, wallets, false, "UserWallet");

            List<Relation> relations = new ArrayList<>();
            relations.add(familyRelation);
            relations.add(walletRelation);
            user1.setRelations(relations);

            list.add(user1);
        }
        return list;
    }
}
