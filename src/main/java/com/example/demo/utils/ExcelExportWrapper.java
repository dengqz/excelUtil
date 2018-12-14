package com.example.demo.utils;


import org.apache.commons.lang.StringUtils;

import javax.servlet.http.HttpServletResponse;
import java.net.URLEncoder;
import java.util.Collection;

/**
 * 包装类
 * @author dqz
 *
 * @param <T>
 */
public class ExcelExportWrapper<T> extends ExcelExportUtil<T> {

    public ExcelExportWrapper(Class claze) {
        super(claze);
    }

    /**
     *
     * <p>
     * 导出带有头部标题行的Excel <br>
     * 时间格式默认：yyyy-MM-dd hh:mm:ss <br>
     * </p>
     * @param fileName 文件名
     * @param title 表格标题
     * @param dataset 数据集合
     * @param response
     * @param version 2003 或者 2007，不传时默认生成2003版本
     */
    public void exportExcel(String fileName, String title,  Collection<T> dataset, HttpServletResponse response,String version) {
        try {
            response.setContentType("application/vnd.ms-excel");
            if(StringUtils.isBlank(version) || EXCEL2003_VERSION.equals(version.trim())){
                response.addHeader("Content-Disposition", "attachment;filename="+ URLEncoder.encode(fileName, "UTF-8") + ".xls");
                exportExcel2003(title,dataset, response.getOutputStream(), "yyyy-MM-dd hh:mm:ss");
            }else{
                response.addHeader("Content-Disposition", "attachment;filename="+ URLEncoder.encode(fileName, "UTF-8") + ".xlsx");
                exportExcel2007(title,dataset, response.getOutputStream(), "yyyy-MM-dd hh:mm:ss");
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
