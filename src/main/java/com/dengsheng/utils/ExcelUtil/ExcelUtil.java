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
        // ???????????????????????????????????????
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
//                logger.warning("????????????????????????????????????????????????" + e.getMessage());
            }
        }
    }


    public static <T extends Relations> List<T> importExcel(Class<T> clazz, String fileName, Boolean containHead) throws IOException {
        Workbook workbook = null;
        Sheet sheet = null;
        Row row = null;
        ExcelUtil excelUtil = new ExcelUtil(fileName, containHead, clazz);
//        String fileName = "C:/Users/Ds/Desktop/student.xls";
        // ??????excel
        List<T> list = new ArrayList<>();
        try {
            list = excelUtil.readExcel();
        } catch (IOException | InvocationTargetException | NoSuchMethodException e) {
            e.printStackTrace();
        }
        return list;
    }

    private <T extends Relations> List<T> readExcel() throws IOException, InvocationTargetException, NoSuchMethodException {
        //??????
        List<T> datas = new ArrayList<>();

        Workbook workbook = null;
        Sheet sheet = null;
        Row row = null;
        //??????Excel??????
        File excelFile = new File(this.fileName.trim());
        InputStream is = null;
        try {
            is = new FileInputStream(excelFile);
            //??????Excel?????????
            if (excelFile.getName().endsWith("xlsx")) {
                workbook = new XSSFWorkbook(is);
            } else {
                workbook = new HSSFWorkbook(is);
            }
            if (workbook == null) {
                throw new ExcelFileNotFind(this.fileName+"?????????");
            }

            // ?????????ExcelClass?????????map???key????????????valueh???
            Map<String, Class<?>> classMap = annotationExcelClassToMap();

            //??????Excel???????????????sheet?????????????????????sheetname??????????????????????????????
            if(this.clazz.isAnnotationPresent(ExcelClass.class)){
                // ?????? "???" ????????????
                String sheetName = (this.clazz.getAnnotation(ExcelClass.class)).sheetName();
                String value = (this.clazz.getAnnotation(ExcelClass.class)).value();
                // classMap????????????????????????
                classMap.remove(value);
                // ????????????????????????????????????sheet
                sheet = workbook.getSheet(sheetName);
            } else {
                sheet = workbook.getSheetAt(0);
            }
            // ??????????????????
            List<Field> tableFields = getClzzOrdeByIndex(this.clazz);
            // ????????????????????????
            Class<?> analysisClass = null;
            // ????????????????????????????????????????????????????????????????????????????????????????????????
            Class<?> lastAnalysisClass = null;
            List<Field> analysisFields = tableFields;
            // ????????????
            for(int rowNum = 1; rowNum <= sheet.getLastRowNum(); rowNum++) {
                row = sheet.getRow(rowNum);
                // ????????????
                if (isRowEmpty(row)){
                    continue;
                }
                List<Field> needAnalysis = new ArrayList<>();
                String firstCellValue = getCellStringValue(row.getCell(0));
                // ??????????????????????????????
                if(firstCellValue != null && !StringUtils.isEmpty(firstCellValue)){
                    // ?????????????????????
                    if(classMap.keySet().contains(firstCellValue)){
                        analysisClass = classMap.get(firstCellValue);
                        analysisFields = getClzzOrdeByIndex(analysisClass);
                    } else {
                        // ?????????????????????
                        analysisClass = this.clazz;
                        analysisFields = tableFields;
                    }
                }
                // ????????????
//                row = sheet.getRow(rowNum);
                // ??????????????????
                Object instance = new Object();
                // ?????????????????????
                if(analysisClass.equals(this.clazz)){
                    instance = this.clazz.newInstance();
                } else {
                    instance = analysisClass.newInstance();
                }
                // ????????????????????????
                for(int i =0; i < analysisFields.size(); i++){
                    Field field = analysisFields.get(i);
                    Short index = field.getAnnotation(ExcelConfig.class).index();
                    if(index != null && index > -1){
                        String methodName = MethodUtils.setMethodName(field.getName());
                        Method method = analysisClass.getMethod(methodName, field.getType());
                        // ??????????????????????????????
                        Cell cell = null;
                        // ????????? ?????????????????????????????????
                        if(!analysisClass.equals(this.clazz)){
                            // ??????????????????????????? ???????????????????????????
                            cell = row.getCell(index+1);
                        } else {
                            cell = row.getCell(index);
                        }
                        if(cell != null){
                            // ??????????????????
                            cell.setCellType(Cell.CELL_TYPE_STRING);
                            // ????????????????????????
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
                        throw new ExcelNotFiledIndexException(field.getName()+"???index????????????????????????");
                    }
                }
                // ?????????????????????
                if(!analysisClass.equals(this.clazz)){
                    List<Relations> dataRelations = datas.get(datas.size()-1).getRelations();
                    // ????????? relations???
                    if(dataRelations == null){
                        datas.get(datas.size()-1).setRelations(new ArrayList<Relations>());
                        dataRelations = datas.get(datas.size()-1).getRelations();
                    }
                    // ?????????????????????????????????List????????????????????????????????????
                    if(analysisClass.equals(lastAnalysisClass)){
                        // ??????????????????????????????????????????????????????relation????????????
                        Relations dataRelation = dataRelations.get(dataRelations.size()-1);
                        Relation relation = (Relation)dataRelation.getRelations().get(dataRelation.getRelations().size()-1);
                        relation.getData().add(instance);
                    } else {
                        // ???????????????????????????????????????????????????????????????
                        Object finalInstance = instance;
                        Relation relation = new Relation(analysisClass, new ArrayList<Object>(){{add(finalInstance);};}, null, null);
                        Relations relations = new Relations();
                        relations.setRelations(new ArrayList<Relation>(){{add(relation);}});
                        dataRelations.add(relations);
                    }
                } else {
                    // ????????????
                    datas.add((T)instance);
                }
                // ????????????????????????????????????
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
     * ????????????
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

    //???????????????????????????????????????????????????
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
     * double???String
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

    // ????????????????????????"com.yaspeed.lpsa"????????????clazz??????
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
     * ?????????ExcelClass ??????????????????????????????Map
     * @return
     */
    private Map<String, Class<?>> annotationExcelClassToMap(){
        Map<String, Class<?>> classMap = new HashMap<>();
        // ?????????????????????ExcelClass????????????
        Set<Class<?>> annotatedWithExcelClass = getTypesAnnotatedWith(ExcelClass.class);
        // ??????????????????annotatedWithExcelClass???ExcelClass???value??????key, class??????value
        for(Class<?> clazz : annotatedWithExcelClass){
            // ??????clazz??????ExcelClass???value???
            String value = clazz.getAnnotation(ExcelClass.class).value();
            classMap.put(value, clazz);
        }
        return classMap;
    }


    private <T> void writeSheet(Workbook workbook, List<T> itemsList, Class<T> clzz, Boolean needHead){
        // ??????T?????????sheetName
        String sheetName = "sheet";
        if(clzz.isAnnotationPresent(ExcelClass.class)){
            // ?????? "???" ????????????
            sheetName = (clzz.getAnnotation(ExcelClass.class)).sheetName();
        }
        Sheet sheet = workbook.createSheet(sheetName);
        try {
            // ?????????????????????????????????
            Integer rowIndex = 0;
            int count = itemsList.size() / 1000;

            // ?????????????????????ExcelConfig??????
            List<Field> fields = getClzzOrdeByIndex(clzz);
            // ??????????????????
            if(needHead){
                // ???????????????
                Row row = sheet.getRow(rowIndex);
                if (null == row) {
                    row = sheet.createRow(rowIndex);
                }
                for(int i = 0; i < fields.size(); i++){
                    if(fields.get(i).getAnnotation(ExcelConfig.class) != null){
                        // ????????????????????????
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

            //  ????????????
            for (int i = 0; i < count + 1; i++) {
                for (Integer j = i * 1000; j< Math.min(itemsList.size(), (i + 1) * 1000); j++) {
                    // ???excel????????????
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
        // ?????????????????????ExcelConfig??????
        List<Field> fields = getClzzOrdeByIndex(clzz);
        // ???????????????
        Row row = sheet.getRow(rowIndex);
        if (null == row) {
            row = sheet.createRow(rowIndex);
        }
        Integer dataPos = 0;
        if(relationTypeName != null){
            // ???????????????
            Cell cell = row.getCell(dataPos);
            if (null == cell) {
                cell = row.createCell(dataPos);
            }
            cell.setCellValue(relationTypeName);
            dataPos = 1;
        }

        // ??????????????????excel?????????
        for (int i = 0; i < fields.size(); i++) {
            Field field = fields.get(i);
            // ?????????????????????????????????ExcelConfig??????
            if(field != null && field.getAnnotationsByType(ExcelConfig.class) != null){
                // ????????????????????????????????????????????????
                Class enumClass = field.getAnnotation(ExcelConfig.class).enumClass();
                Object value = null;
                // ????????????????????????
                if(enumClass != EmptyEums.class){
                    // ????????????????????????
                    String fieldName = field.getAnnotation(ExcelConfig.class).enumFieldName();
                    // ?????????????????????????????????
                    Object needEnumKey = getFieldValueByName(fieldName, item);
                    // ?????????????????????????????????
                    Object[] objects = enumClass.getEnumConstants();
                    for (Object obj : objects) {
                        Class<?> enumClzz = obj.getClass();
                        // ???????????????getKey??????
                        Method getKeyMethod = enumClzz.getDeclaredMethod("getKey");
                        Object enumKey = getKeyMethod.invoke(obj);
                        // ???????????????key??????????????????needEnumKey????????????????????????value???field????????????
                        if(enumKey.toString().equals(needEnumKey.toString())){
                            Method getValueMethod = enumClzz.getDeclaredMethod("getValue");
                            value = getValueMethod.invoke(obj);
                            break;
                        }
                    }
                } else {
                    value = getFieldValueByName(field.getName(), item);
                }
                // ???????????????????????????index???
                Integer columnNum = field.getAnnotation(ExcelConfig.class).index() + dataPos;
                Cell cell = row.getCell(columnNum);
                if (null == cell) {
                    cell = row.createCell(columnNum);
                }
                // ??????????????????
                if (value != null){
                    parseFiledValue(cell, value,null);
                }
            }

        }
        rowIndex = rowIndex+1;
        // ?????????????????????
        // clazz ?????????relations????????????????????????
        if(clzz.getSuperclass() == Relations.class){
            List<Relation> relations = (List<Relation>)getFieldValueByName("relations",item);
            if(relations != null){
                for (Relation relation : relations) {
                    List relationDatas = relation.getData();
                    for (int i = 0; i < relationDatas.size(); i++) {
                        // ?????????????????????????????????
                        String subRelationTypeName = "";
                        // ????????????????????????????????????
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
                throw new RuntimeException("???????????????????????????"+paresValue);
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
     * ?????? ExcelConfig ?????????index ??????
     * @param clazz
     * @return
     */
    private List<Field> getClzzOrdeByIndex(Class clazz){
        Field[] fields=clazz.getDeclaredFields();
        Field[] fieldsHasIndex = new Field[fields.length];
        List<Field> fieldsNoHasIndex = new ArrayList<>();
        // ????????????????????????
        Integer pos = 0;
        for (int i = 0; i < fields.length; i++) {
            boolean annotationPresent = fields[i].isAnnotationPresent(ExcelConfig.class);
            if (annotationPresent) {
                // ???????????????
                Short index = fields[i].getAnnotation(ExcelConfig.class).index();
                if(index != null && index != -1){
                    fieldsHasIndex[pos] = fields[i];
                    pos ++;
                } else {
                    fieldsNoHasIndex.add(fields[i]);
                }
            }
        }
        // ????????????????????????
        fieldsNoHasIndex.addAll(0, Arrays.asList(fieldsHasIndex).subList(0, pos));
        // ???????????????????????????
        fieldsNoHasIndex.remove(null);
        return fieldsNoHasIndex;
    }

    /**
     * ?????????????????????
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
