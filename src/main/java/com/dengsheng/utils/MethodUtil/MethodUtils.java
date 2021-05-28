package com.dengsheng.utils.MethodUtil;

public class MethodUtils {

    private static final String SET_PREFIX = "set";
    private static final String GET_PREFIX = "get";

    private static String capitalize(String name) {
        if (name == null || name.length() == 0) {
            return name;
        }
        return name.substring(0, 1).toUpperCase() + name.substring(1);
    }

    public static String setMethodName(String propertyName) {
        return SET_PREFIX + capitalize(propertyName);
    }

    public static String getMethodName(String propertyName) {
        return GET_PREFIX + capitalize(propertyName);
    }
}
