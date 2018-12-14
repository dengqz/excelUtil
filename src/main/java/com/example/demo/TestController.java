package com.example.demo;

import com.example.demo.utils.ExcelExportUtil;
import com.example.demo.utils.ExcelExportWrapper;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.ResponseBody;
import org.springframework.web.bind.annotation.RestController;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

/**
 * @author dqz
 */
@RestController
@RequestMapping(value = "/api")
public class TestController {
    @RequestMapping("/ping")
    public @ResponseBody String index(){
        return "pong";
    }
    @GetMapping(value = "/getExcel")
    public void getExcel(HttpServletRequest request, HttpServletResponse response) {
        // 准备数据
        List<Student> list = new ArrayList<>();
        for (int i = 0; i < 10; i++) {
            list.add(new Student(111,"张三asdf","男",new Date(),new Date(),new Date()));
            list.add(new Student(111,"李四asd","男",new Date(),new Date(),new Date()));
            list.add(new Student(111,"王五","女",new Date(),new Date(),new Date()));
        }
        String fileName = "excel1";
        ExcelExportWrapper<Student> util = new ExcelExportWrapper<>(Student.class);
        util.exportExcel(fileName, fileName, list, response, ExcelExportUtil.EXCEL2003_VERSION);
    }
}
