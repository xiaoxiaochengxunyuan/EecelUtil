package com.dengsheng.utils.ExcelUtil;

import com.dengsheng.utils.DateUtil.DateUtil;
import com.dengsheng.utils.ExcelUtil.enums.EmptyEums;
import com.dengsheng.utils.ExcelUtil.exceptions.ExcelFileNotFind;
import com.dengsheng.utils.ExcelUtil.exceptions.ExcelNotFiledIndexException;
import com.dengsheng.utils.ExcelUtil.interfaces.ExcelClass;
import com.dengsheng.utils.ExcelUtil.interfaces.ExcelConfig;
import com.dengsheng.utils.ExcelUtil.model.Relation;
import com.dengsheng.utils.ExcelUtil.model.Relations;
import com.dengsheng.utils.MethodUtil.MethodUtils;
import lombok.Getter;
import org.apache.commons.lang3.StringUtils;
import org.apache.commons.lang3.time.DateFormatUtils;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.reflections.Reflections;

import java.io.*;
import java.lang.annotation.Annotation;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.math.BigDecimal;
import java.math.BigInteger;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class ExcelUtil<T> {

    @Getter
    private String fileName;

    @Getter
    private Boolean containHead;

    @Getter
    private Class<T> clazz;

    private ExcelUtil(){}

    public ExcelUtil(String fileName, Boolean containHead, Class clazz){
        this.fileName = fileName;
        this.containHead = containHead;
        this.clazz = clazz;
    }


    public static  <T> void exportExcel(List<T> list, String fileName, Class<T> clazz, Boolean containHead) {
        ExcelUtil excelUtil = new ExcelUtil(fileName, containHead, clazz);
        Workbook workbook = null;
        File exportFile = new File(excelUtil.getFileName());
        FileOutputStream fileOut = null;
        // 以文件的形式输出工作簿对象
        try {
//            File exportFile = new File(excelUtil.getFileName());
//            if (exportFile.exists()) {
//                exportFile.delete();
//            }
//            exportFile.createNewFile();
            if(excelUtil.getFileName().endsWith(".xlsx")){
                workbook = new XSSFWorkbook();
            } else {
                workbook = new HSSFWorkbook();
            }
            excelUtil.writeSheet(workbook, list, clazz, excelUtil.getContainHead());

            fileOut = new FileOutputStream(exportFile);
            workbook.write(fileOut);
            fileOut.flush();
            fileOut.close();
        } catch (Exception e) {
            e.printStackTrace();
        }finally {
            try {
                if (null != fileOut) {
                    fileOut.close();
                }
            } catch (IOException e) {
//                logger.warning("关闭输出流时发生错误，错误原因：" + e.getMessage());
            }
        }
    }


    public static <T extends Relations> List<T> importExcel(Class<T> clazz, String fileName, Boolean containHead) throws IOException {
        Workbook workbook = null;
        Sheet sheet = null;
        Row row = null;
        ExcelUtil excelUtil = new ExcelUtil(fileName, containHead, clazz);
//        String fileName = "C:/Users/Ds/Desktop/student.xls";
        // 读取excel
        List<T> list = new ArrayList<>();
        try {
            list = excelUtil.readExcel();
        } catch (IOException | InvocationTargetException | NoSuchMethodException e) {
            e.printStackTrace();
        }
        return list;
    }

    private <T extends Relations> List<T> readExcel() throws IOException, InvocationTargetException, NoSuchMethodException {
        //获取
        List<T> datas = new ArrayList<>();

        Workbook workbook = null;
        Sheet sheet = null;
        Row row = null;
        //读取Excel文件
        File excelFile = new File(this.fileName.trim());
        InputStream is = null;
        try {
            is = new FileInputStream(excelFile);
            //获取Excel工作薄
            if (excelFile.getName().endsWith("xlsx")) {
                workbook = new XSSFWorkbook(is);
            } else {
                workbook = new HSSFWorkbook(is);
            }
            if (workbook == null) {
                throw new ExcelFileNotFind(this.fileName+"未找到");
            }

            // 注解了ExcelClass的类，map的key为注解的valueh值
            Map<String, Class<?>> classMap = annotationExcelClassToMap();

            //获取Excel需要解析的sheet，根据注解的的sheetname获取，否则就取第一个
            if(this.clazz.isAnnotationPresent(ExcelClass.class)){
                // 获取 "类" 上的注解
                String sheetName = (this.clazz.getAnnotation(ExcelClass.class)).sheetName();
                String value = (this.clazz.getAnnotation(ExcelClass.class)).value();
                // classMap移除表格主体数据
                classMap.remove(value);
                // 根据对象的注解找到对应的sheet
                sheet = workbook.getSheet(sheetName);
            } else {
                sheet = workbook.getSheetAt(0);
            }
            // 主体数据解析
            List<Field> tableFields = getClzzOrdeByIndex(this.clazz);
            // 记录当前数据类型
            Class<?> analysisClass = null;
            // 记录上一个数据类型，用于判单当前数据类型和上一个数据类型是否一直
            Class<?> lastAnalysisClass = null;
            List<Field> analysisFields = tableFields;
            // 读取数据
            for(int rowNum = 1; rowNum <= sheet.getLastRowNum(); rowNum++) {
                row = sheet.getRow(rowNum);
                // 跳过空行
                if (isRowEmpty(row)){
                    continue;
                }
                List<Field> needAnalysis = new ArrayList<>();
                String firstCellValue = getCellStringValue(row.getCell(0));
                // 第一个单元格是否为空
                if(firstCellValue != null && !StringUtils.isEmpty(firstCellValue)){
                    // 是否是关系字段
                    if(classMap.keySet().contains(firstCellValue)){
                        analysisClass = classMap.get(firstCellValue);
                        analysisFields = getClzzOrdeByIndex(analysisClass);
                    } else {
                        // 设置为主题字段
                        analysisClass = this.clazz;
                        analysisFields = tableFields;
                    }
                }
                // 获取一行
//                row = sheet.getRow(rowNum);
                // 获取对象实列
                Object instance = new Object();
                // 是否为关系对象
                if(analysisClass.equals(this.clazz)){
                    instance = this.clazz.newInstance();
                } else {
                    instance = analysisClass.newInstance();
                }
                // 根据字段读取数据
                for(int i =0; i < analysisFields.size(); i++){
                    Field field = analysisFields.get(i);
                    Short index = field.getAnnotation(ExcelConfig.class).index();
                    if(index != null && index > -1){
                        String methodName = MethodUtils.setMethodName(field.getName());
                        Method method = analysisClass.getMethod(methodName, field.getType());
                        // 获取字段对应的单元格
                        Cell cell = null;
                        // 判断是 数据本体，还是关系数据
                        if(!analysisClass.equals(this.clazz)){
                            // 关系数据，第一列为 关系名称，所以跳过
                            cell = row.getCell(index+1);
                        } else {
                            cell = row.getCell(index);
                        }
                        if(cell != null){
                            // 转为为字符串
                            cell.setCellType(Cell.CELL_TYPE_STRING);
                            // 日期字段担负处理
                            if (DateUtil.isDateFied(field)) {
                                Date date = cell.getDateCellValue();
                                if(date != null){
                                    method.invoke(instance, cell.getDateCellValue());
                                }
                            } else {
                                String value = getCellStringValue(cell);
                                method.invoke(instance, convertType(field.getType(), value.trim()));
                            }
                        }
                    } else {
                        throw new ExcelNotFiledIndexException(field.getName()+"无index值，无法解析数据");
                    }
                }
                // 如果是关系数据
                if(!analysisClass.equals(this.clazz)){
                    List<Relations> dataRelations = datas.get(datas.size()-1).getRelations();
                    // 初始化 relations值
                    if(dataRelations == null){
                        datas.get(datas.size()-1).setRelations(new ArrayList<Relations>());
                        dataRelations = datas.get(datas.size()-1).getRelations();
                    }
                    // 获取最有一个本体数据的List的最后一个本体的关系数据
                    if(analysisClass.equals(lastAnalysisClass)){
                        // 获取关系中的最后一个类型，并把当前的relation添加进去
                        Relations dataRelation = dataRelations.get(dataRelations.size()-1);
                        Relation relation = (Relation)dataRelation.getRelations().get(dataRelation.getRelations().size()-1);
                        relation.getData().add(instance);
                    } else {
                        // 如果和上一个不同，说明是不同关系数据，创建
                        Object finalInstance = instance;
                        Relation relation = new Relation(analysisClass, new ArrayList<Object>(){{add(finalInstance);};}, null, null);
                        Relations relations = new Relations();
                        relations.setRelations(new ArrayList<Relation>(){{add(relation);}});
                        dataRelations.add(relations);
                    }
                } else {
                    // 本体数据
                    datas.add((T)instance);
                }
                // 保存上一个数据解析的类型
                lastAnalysisClass = analysisClass;
            }
            is.close();
        } catch (IOException | InstantiationException | IllegalAccessException e) {
            e.printStackTrace();
        } finally {
            if(is != null){
                is.close();
            }
        }
        return datas;
    }

    /**
     * 类型转换
     *
     * @param clazz
     * @param value
     * @return
     */
    private static Object convertType(Class clazz, String value) {
        if(StringUtils.isEmpty(value)){
            return null;
        }
        if (Integer.class.equals(clazz) || int.class.equals(clazz)) {
            if(value.contains(".")){
                return Double.valueOf(value).intValue();
            }
            return Integer.valueOf(value);
        }
        if (Short.class.equals(clazz) || short.class.equals(clazz)) {
            if(value.contains(".")){
                return Double.valueOf(value).shortValue();
            }
            return Short.valueOf(value);
        }
        if (Byte.class.equals(clazz) || byte.class.equals(clazz)) {
            return Byte.valueOf(value);
        }
        if (Character.class.equals(clazz) || char.class.equals(clazz)) {
            return value.charAt(0);
        }
        if (Long.class.equals(clazz) || long.class.equals(clazz)) {
            if(value.contains(".")){
                return Double.valueOf(value).longValue();
            }
            return Long.valueOf(value);
        }
        if (Float.class.equals(clazz) || float.class.equals(clazz)) {
            if(value.contains(".")){
                return Double.valueOf(value).floatValue();
            }
            return Float.valueOf(value);
        }
        if (Double.class.equals(clazz) || double.class.equals(clazz)) {
            return Double.valueOf(value);
        }
        if (Boolean.class.equals(clazz) || boolean.class.equals(clazz)) {
            return Boolean.valueOf(value.toLowerCase());
        }
        if (BigDecimal.class.equals(clazz)) {
            return new BigDecimal(value);
        }
        return value;
    }

    //获取单元格各类型值，返回字符串类型
    private static String getCellStringValue(Cell cell) {
        if (cell == null) {
            return "";
        } else {
            if(cell.getCellType() == Cell.CELL_TYPE_NUMERIC) {
                if (HSSFDateUtil.isCellDateFormatted(cell)) {
                    Date date = cell.getDateCellValue();
                    return DateFormatUtils.format(date, DateUtil.DATEFORMATSECOND);
                } else {
                    return getRealStringValueOfDouble(cell.getNumericCellValue());
                }
            }
            cell.setCellType(Cell.CELL_TYPE_STRING);
            return cell.getStringCellValue().trim();
        }
    }

    /**
     * double转String
     * @param d
     * @return
     */
    private static String getRealStringValueOfDouble(Double d) {
        String doubleStr = d.toString();
        boolean b = doubleStr.contains("E") || doubleStr.contains("e");
        int indexOfPoint = doubleStr.indexOf('.');
        if (b) {
            int indexOfE = doubleStr.indexOf('E');
            BigInteger xs = new BigInteger(doubleStr.substring(indexOfPoint
                    + BigInteger.ONE.intValue(), indexOfE));
            int pow = Integer.valueOf(doubleStr.substring(indexOfE
                    + BigInteger.ONE.intValue()));
            int xsLen = xs.toByteArray().length;
            int scale = xsLen - pow > 0 ? xsLen - pow : 0;
            doubleStr = String.format("%." + scale + "f", d);
        } else {
            Pattern p = Pattern.compile(".0$");
            Matcher m = p.matcher(doubleStr);
            if (m.find()) {
                doubleStr = doubleStr.replace(".0", "");
            }
        }
        return doubleStr;
    }

    // 获取工程全路径下"com.yaspeed.lpsa"的注解了clazz的类
    private Set<Class<?>> getTypesAnnotatedWith(Class<? extends Annotation> clazz){

        String path = this.getClass().getResource("").getPath();
        String pack = path.substring(path.lastIndexOf("classes/")+8).replace("/",".");
        if(pack.endsWith(".")){
            pack = pack.substring(0, pack.length()-1);
        }


        Reflections f = new Reflections(pack);
        return f.getTypesAnnotatedWith(clazz);
    }

    /**
     * 获取被ExcelClass 注解过的类，并转化为Map
     * @return
     */
    private Map<String, Class<?>> annotationExcelClassToMap(){
        Map<String, Class<?>> classMap = new HashMap<>();
        // 获取所有实现了ExcelClass注解的类
        Set<Class<?>> annotatedWithExcelClass = getTypesAnnotatedWith(ExcelClass.class);
        // 处理数据：把annotatedWithExcelClass的ExcelClass的value作为key, class作为value
        for(Class<?> clazz : annotatedWithExcelClass){
            // 获取clazz上的ExcelClass的value值
            String value = clazz.getAnnotation(ExcelClass.class).value();
            classMap.put(value, clazz);
        }
        return classMap;
    }


    private <T> void writeSheet(Workbook workbook, List<T> itemsList, Class<T> clzz, Boolean needHead){
        // 获取T对应的sheetName
        String sheetName = "sheet";
        if(clzz.isAnnotationPresent(ExcelClass.class)){
            // 获取 "类" 上的注解
            sheetName = (clzz.getAnnotation(ExcelClass.class)).sheetName();
        }
        Sheet sheet = workbook.createSheet(sheetName);
        try {
            // 在相应的单元格进行赋值
            Integer rowIndex = 0;
            int count = itemsList.size() / 1000;

            // 获取对象的上的ExcelConfig注解
            List<Field> fields = getClzzOrdeByIndex(clzz);
            // 创建文件题头
            if(needHead){
                // 创建行数据
                Row row = sheet.getRow(rowIndex);
                if (null == row) {
                    row = sheet.createRow(rowIndex);
                }
                for(int i = 0; i < fields.size(); i++){
                    if(fields.get(i).getAnnotation(ExcelConfig.class) != null){
                        // 获取表字段的名称
                        String headValue = fields.get(i).getAnnotation(ExcelConfig.class).value();
                        Cell cell = row.getCell(i);
                        if (null == cell) {
                            cell = row.createCell(i);
                        }
                        if(!StringUtils.isEmpty(headValue)){
                            cell.setCellValue(headValue);
                        }
                    }
                }
                rowIndex = rowIndex+1;
            }

            //  写入数据
            for (int i = 0; i < count + 1; i++) {
                for (Integer j = i * 1000; j< Math.min(itemsList.size(), (i + 1) * 1000); j++) {
                    // 在excel中写入值
                    T item = itemsList.get(j);
                    if (item != null) {
                        rowIndex = writeCell(sheet, item, rowIndex, clzz, null);
                    }
                }
            }
        } catch (Exception e){
            e.printStackTrace();
        }
    }

    private <T> Integer writeCell(Sheet sheet, T item, Integer rowIndex, Class clzz, String relationTypeName) throws NoSuchMethodException, InvocationTargetException, IllegalAccessException {
        // 获取对象的上的ExcelConfig注解
        List<Field> fields = getClzzOrdeByIndex(clzz);
        // 创建行数据
        Row row = sheet.getRow(rowIndex);
        if (null == row) {
            row = sheet.createRow(rowIndex);
        }
        Integer dataPos = 0;
        if(relationTypeName != null){
            // 添加占位符
            Cell cell = row.getCell(dataPos);
            if (null == cell) {
                cell = row.createCell(dataPos);
            }
            cell.setCellValue(relationTypeName);
            dataPos = 1;
        }

        // 处理需要写入excel的数据
        for (int i = 0; i < fields.size(); i++) {
            Field field = fields.get(i);
            // 判断获取的字段必须要有ExcelConfig标识
            if(field != null && field.getAnnotationsByType(ExcelConfig.class) != null){
                // 获取改字段是否是需要经过枚举处理
                Class enumClass = field.getAnnotation(ExcelConfig.class).enumClass();
                Object value = null;
                // 判断是否需要枚举
                if(enumClass != EmptyEums.class){
                    // 获取枚举所需的值
                    String fieldName = field.getAnnotation(ExcelConfig.class).enumFieldName();
                    // 获取枚举的目标字段的值
                    Object needEnumKey = getFieldValueByName(fieldName, item);
                    // 获取枚举的中每个枚举类
                    Object[] objects = enumClass.getEnumConstants();
                    for (Object obj : objects) {
                        Class<?> enumClzz = obj.getClass();
                        // 每个枚举的getKey方法
                        Method getKeyMethod = enumClzz.getDeclaredMethod("getKey");
                        Object enumKey = getKeyMethod.invoke(obj);
                        // 如果枚举的key和需要处理的needEnumKey相同，则获取对应value为field字段的值
                        if(enumKey.toString().equals(needEnumKey.toString())){
                            Method getValueMethod = enumClzz.getDeclaredMethod("getValue");
                            value = getValueMethod.invoke(obj);
                            break;
                        }
                    }
                } else {
                    value = getFieldValueByName(field.getName(), item);
                }
                // 获取注解字段对应的index值
                Integer columnNum = field.getAnnotation(ExcelConfig.class).index() + dataPos;
                Cell cell = row.getCell(columnNum);
                if (null == cell) {
                    cell = row.createCell(columnNum);
                }
                // 判断数据类型
                if (value != null){
                    parseFiledValue(cell, value,null);
                }
            }

        }
        rowIndex = rowIndex+1;
        // 写入关系型数据
        // clazz 集成了relations话，说明存在关系
        if(clzz.getSuperclass() == Relations.class){
            List<Relation> relations = (List<Relation>)getFieldValueByName("relations",item);
            if(relations != null){
                for (Relation relation : relations) {
                    List relationDatas = relation.getData();
                    for (int i = 0; i < relationDatas.size(); i++) {
                        // 判断是否要添加参数名称
                        String subRelationTypeName = "";
                        // 只有在第一行添加参数名称
                        if( i == 0){
                            subRelationTypeName = relation.getRelationTypeName();
                        }
                        rowIndex = writeCell(sheet, relationDatas.get(i), rowIndex, relation.getClazz(), subRelationTypeName);
                    }
                }
            }
        }
        return rowIndex;
    }

    private <T> void parseFiledValue(Cell cell, T paresValue, String dateFormat){
        Class clazz = paresValue.getClass();
        if (paresValue != null){
            if(clazz == Integer.class){
                cell.setCellValue(Integer.parseInt(paresValue.toString()));
            } else if(clazz == Long.class){
                cell.setCellValue(Long.parseLong(paresValue.toString()));
            } else if(clazz == String.class){
                cell.setCellValue(paresValue.toString());
            } else if(clazz == Boolean.class){
                cell.setCellValue(Boolean.valueOf(paresValue.toString()));
            } else if(clazz == Short.class){
                cell.setCellValue(Short.valueOf(paresValue.toString()));
            } else if(clazz == Double.class){
                cell.setCellValue(Double.valueOf(paresValue.toString()));
            } else if(clazz == Date.class){
                if(StringUtils.isEmpty(dateFormat)){
                    dateFormat = DateUtil.DATEFORMATSECOND;
                }
                DateFormat dateformat = new SimpleDateFormat(dateFormat);
                cell.setCellValue( dateformat.format(paresValue));
            } else {
                throw new RuntimeException("不支持的数据类型："+paresValue);
            }
        }
    }

    private Object getFieldValueByName(String fieldName, Object o) {
        try {
            String firstLetter = fieldName.substring(0, 1).toUpperCase();
            String getter = "get" + firstLetter + fieldName.substring(1);
            Method method = o.getClass().getMethod(getter, new Class[] {});
            Object value = method.invoke(o, new Object[] {});
            return value;
        } catch (Exception e) {
            return null;
        }
    }

    /**
     * 根据 ExcelConfig 注解的index 排序
     * @param clazz
     * @return
     */
    private List<Field> getClzzOrdeByIndex(Class clazz){
        Field[] fields=clazz.getDeclaredFields();
        Field[] fieldsHasIndex = new Field[fields.length];
        List<Field> fieldsNoHasIndex = new ArrayList<>();
        // 用于处理有字段的
        Integer pos = 0;
        for (int i = 0; i < fields.length; i++) {
            boolean annotationPresent = fields[i].isAnnotationPresent(ExcelConfig.class);
            if (annotationPresent) {
                // 获取注解值
                Short index = fields[i].getAnnotation(ExcelConfig.class).index();
                if(index != null && index != -1){
                    fieldsHasIndex[pos] = fields[i];
                    pos ++;
                } else {
                    fieldsNoHasIndex.add(fields[i]);
                }
            }
        }
        // 有序的字段往前排
        fieldsNoHasIndex.addAll(0, Arrays.asList(fieldsHasIndex).subList(0, pos));
        // 去除数组中的空元素
        fieldsNoHasIndex.remove(null);
        return fieldsNoHasIndex;
    }

    /**
     * 判断是否为空行
     * @param row
     * @return
     */
    private  boolean isRowEmpty(Row row){
        for (int c = row.getFirstCellNum(); c < row.getLastCellNum(); c++) {
            Cell cell = row.getCell(c);
            if (cell != null && cell.getCellType() != Cell.CELL_TYPE_BLANK)
                return false;
        }
        return true;
    }
}
