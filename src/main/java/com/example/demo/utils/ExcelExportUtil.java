package com.example.demo.utils;

import com.example.demo.annotations.ExcelVo;
import org.apache.commons.lang.StringUtils;
import org.apache.commons.lang.time.DateFormatUtils;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.xssf.usermodel.*;

import java.io.IOException;
import java.io.OutputStream;
import java.lang.reflect.Field;
import java.util.*;
import java.util.regex.Pattern;


/**
 * Excel 导出 工具类
 * @author dqz
 * @param <T>
 */
public class ExcelExportUtil<T> {

    private Class claze;

    public ExcelExportUtil(Class claze) {
        this.claze = claze;
    }


    private static final Pattern PATTERN = Pattern.compile("^//d+(//.//d+)?$");

    /**
     * 2007 版本以上 最大支持1048576行
     */
    public  final static String  EXCEL2007_VERSION = "2007";
    /**
     * 2003 版本 最大支持65536 行
     */
    public  final static String  EXCEL2003_VERSION = "2003";


    public void exportExcel(String title, Collection<T> dataset, OutputStream out, String version) {
        if(StringUtils.isEmpty(version) || EXCEL2003_VERSION.equals(version.trim())){
            exportExcel2003(title,  dataset, out, "yyyy-MM-dd HH:mm:ss");
        }else{
            exportExcel2007(title,  dataset, out, "yyyy-MM-dd HH:mm:ss");
        }
    }
    @SuppressWarnings({ "unchecked", "rawtypes" })
    public void exportExcel2003(String title, Collection<T> dataset, OutputStream out, String pattern) {

        // 声明一个工作薄
        HSSFWorkbook workbook = new HSSFWorkbook();
        // 生成一个表格
        HSSFSheet sheet = workbook.createSheet(title);
        // 设置表格默认列宽度为15个字节
        sheet.setDefaultColumnWidth(20);
        // 生成表格标题头样式
        HSSFCellStyle headerStyle = setExcelTableHeaderStyle(workbook);
        // 生成表格内容样式
        HSSFCellStyle contentStyle = setExcelTableContentStyle(workbook);

        // 产生表格标题行
        HSSFRow row = sheet.createRow(0);
        HSSFCell cellHeader;
        List<Field> fieldList = getFieldList();

        for (int i = 0; i < fieldList.size(); i++) {
            cellHeader = row.createCell(i);
            cellHeader.setCellStyle(headerStyle);
            cellHeader.setCellValue(new HSSFRichTextString(fieldList.get(i).getAnnotation(ExcelVo.class).name()));
        }
        // 遍历集合数据，产生数据行
        Iterator<T> it = dataset.iterator();
        int index = 0;
        T t;
        HSSFCell cell;
        while (it.hasNext()) {
            index++;
            row = sheet.createRow(index);
            t = it.next();
            for (int j = 0; j < fieldList.size(); j++) {
                cell = row.createCell(j);
                cell.setCellStyle(contentStyle);
                Field field2 = fieldList.get(j);
                try{
                    cell.setCellValue(covertAttrType(field2,t));
                }catch (IllegalAccessException e){
                    e.printStackTrace();
                }

            }

        }
        try {
            workbook.write(out);
        } catch (IOException e) {
            e.printStackTrace();
        }finally{
            if(null != out){
                try {
                    out.flush();
                    out.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }
    }

    public void exportExcel2007(String title, Collection<T> dataset, OutputStream out, String pattern) {
        // 声明一个工作薄
        XSSFWorkbook workbook = new XSSFWorkbook();
        // 生成一个表格
        XSSFSheet sheet = workbook.createSheet(title);
        // 设置表格默认列宽度为15个字节
        sheet.setDefaultColumnWidth(20);
        // 生成一个样式
        XSSFCellStyle headerStyle = setExcelTableHeaderStyle(workbook);
        // 生成并设置另一个样式
        XSSFCellStyle contentStyle = setExcelTableContentStyle(workbook);

        // 产生表格标题行
        XSSFRow row = sheet.createRow(0);
        XSSFCell cellHeader;
        List<Field> fieldList = getFieldList();
        for (int i = 0; i < fieldList.size(); i++) {
            cellHeader = row.createCell(i);
            cellHeader.setCellStyle(headerStyle);
            cellHeader.setCellValue(new XSSFRichTextString(fieldList.get(i).getAnnotation(ExcelVo.class).name()));
        }

        // 遍历集合数据，产生数据行
        Iterator<T> it = dataset.iterator();

        int index = 0;
        T t;
        XSSFCell cell;
        while (it.hasNext()) {
            index++;
            row = sheet.createRow(index);
            t = it.next();
            for (int j = 0; j < fieldList.size(); j++) {
                cell = row.createCell(j);
                cell.setCellStyle(contentStyle);
                Field field2 = fieldList.get(j);
                try{
                    cell.setCellValue(covertAttrType(field2,t));
                }catch (IllegalAccessException e){
                    e.printStackTrace();
                }

            }

        }
        try {
            workbook.write(out);
        } catch (IOException e) {
            e.printStackTrace();
        }finally{
            if(null != out){
                try {
                    out.flush();
                    out.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }
    }

    private HSSFCellStyle setExcelTableHeaderStyle(HSSFWorkbook workbook){
        // 生成一个样式
        HSSFCellStyle headerStyle = workbook.createCellStyle();
        // 设置这些样式
        headerStyle.setFillForegroundColor(HSSFColor.GREY_50_PERCENT.index);
        headerStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
        headerStyle.setBorderBottom(HSSFCellStyle.BORDER_THIN);
        headerStyle.setBorderLeft(HSSFCellStyle.BORDER_THIN);
        headerStyle.setBorderRight(HSSFCellStyle.BORDER_THIN);
        headerStyle.setBorderTop(HSSFCellStyle.BORDER_THIN);
        headerStyle.setAlignment(HSSFCellStyle.ALIGN_CENTER);
        // 生成一个字体
        HSSFFont headerFont = workbook.createFont();
        headerFont.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
        headerFont.setFontName("宋体");
        headerFont.setColor(HSSFColor.WHITE.index);
        headerFont.setFontHeightInPoints((short) 11);
        // 把字体应用到当前的样式
        headerStyle.setFont(headerFont);
        return headerStyle;
    }
    private HSSFCellStyle setExcelTableContentStyle(HSSFWorkbook workbook){
        HSSFCellStyle contentStyle = workbook.createCellStyle();
        contentStyle.setFillForegroundColor(HSSFColor.WHITE.index);
        contentStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
        contentStyle.setBorderBottom(HSSFCellStyle.BORDER_THIN);
        contentStyle.setBorderLeft(HSSFCellStyle.BORDER_THIN);
        contentStyle.setBorderRight(HSSFCellStyle.BORDER_THIN);
        contentStyle.setBorderTop(HSSFCellStyle.BORDER_THIN);
        contentStyle.setAlignment(HSSFCellStyle.ALIGN_CENTER);
        contentStyle.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);
        // 生成一个字体
        HSSFFont contentFont = workbook.createFont();
        contentFont.setBoldweight(HSSFFont.BOLDWEIGHT_NORMAL);
        // 把字体应用到当前的样式
        contentStyle.setFont(contentFont);
        return contentStyle;
    }
    private XSSFCellStyle setExcelTableHeaderStyle(XSSFWorkbook workbook){
        XSSFCellStyle headerStyle = workbook.createCellStyle();
        // 设置这些样式
        headerStyle.setFillForegroundColor(new XSSFColor(java.awt.Color.gray));
        headerStyle.setFillPattern(XSSFCellStyle.SOLID_FOREGROUND);
        headerStyle.setBorderBottom(XSSFCellStyle.BORDER_THIN);
        headerStyle.setBorderLeft(XSSFCellStyle.BORDER_THIN);
        headerStyle.setBorderRight(XSSFCellStyle.BORDER_THIN);
        headerStyle.setBorderTop(XSSFCellStyle.BORDER_THIN);
        headerStyle.setAlignment(XSSFCellStyle.ALIGN_CENTER);
        // 生成一个字体
        XSSFFont headerFont = workbook.createFont();
        headerFont.setBoldweight(XSSFFont.BOLDWEIGHT_BOLD);
        headerFont.setFontName("宋体");
        headerFont.setColor(new XSSFColor(java.awt.Color.BLACK));
        headerFont.setFontHeightInPoints((short) 11);
        // 把字体应用到当前的样式
        headerStyle.setFont(headerFont);
        return headerStyle;
    }
    private XSSFCellStyle setExcelTableContentStyle(XSSFWorkbook workbook){
        // 生成并设置一个样式
        XSSFCellStyle contentStyle = workbook.createCellStyle();
        contentStyle.setFillForegroundColor(new XSSFColor(java.awt.Color.WHITE));
        contentStyle.setFillPattern(XSSFCellStyle.SOLID_FOREGROUND);
        contentStyle.setBorderBottom(XSSFCellStyle.BORDER_THIN);
        contentStyle.setBorderLeft(XSSFCellStyle.BORDER_THIN);
        contentStyle.setBorderRight(XSSFCellStyle.BORDER_THIN);
        contentStyle.setBorderTop(XSSFCellStyle.BORDER_THIN);
        contentStyle.setAlignment(XSSFCellStyle.ALIGN_CENTER);
        contentStyle.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
        // 生成另一个字体
        XSSFFont contentFont = workbook.createFont();
        contentFont.setBoldweight(XSSFFont.BOLDWEIGHT_NORMAL);
        // 把字体应用到当前的样式
        contentStyle.setFont(contentFont);
        return contentStyle;
    }
    /**
     * 获取带注解的字段 并且排序
     * @return
     */
    private List<Field> getFieldList() {
        Field[] fields = this.claze.getDeclaredFields();
        // 无序
        List<Field> fieldList = new ArrayList<Field>();
        // 排序后的字段
        List<Field> fieldSortList = new LinkedList<Field>();
        int length = fields.length;
        int sort = 0;
        Field field = null;
        // 获取带注解的字段
        for (int i = 0; i < length; i++) {
            field = fields[i];
            if (field.isAnnotationPresent(ExcelVo.class)) {
                fieldList.add(field);
            }
        }

        length = fieldList.size();

        for (int i = 1; i <= length; i++) {
            for (int j = 0; j < length; j++) {
                field = fieldList.get(j);
                ExcelVo exceVo = field.getAnnotation(ExcelVo.class);
                field.setAccessible(true);
                sort = exceVo.sort();
                if (sort == i) {
                    fieldSortList.add(field);
                    continue;
                }
            }
        }
        return fieldSortList;
    }
    /**
     * 类型转换 转为String
     */
    private String covertAttrType(Field field, T obj) throws IllegalAccessException {
        if (field.get(obj) == null) {
            return "";
        }
        ExcelVo excelVo = field.getAnnotation(ExcelVo.class);
        Class type = field.getType();
        String format = excelVo.dateFormat();
        if (type==Date.class) {
            return DateFormatUtils.format((Date)field.get(obj), format);
        }if(type==Boolean.class){
            return (Boolean)field.get(obj)?"是":"否";
        }else{
            return field.get(obj).toString();
        }
    }

}
