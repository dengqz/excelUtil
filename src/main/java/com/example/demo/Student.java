package com.example.demo;

import com.example.demo.annotations.ExcelVo;

import java.util.Date;


/**
 * @author dqz
 */
public class Student {
    @ExcelVo(sort = 1,name = "id")
    private int id;
    @ExcelVo(sort = 3,name = "姓名")
    private String name;
    @ExcelVo(sort = 2,name = "性别")
    private String sex;
    @ExcelVo(sort = 1,name = "b日期",dateFormat = "yyyy-MM-dd hh:mm:ss")
    private Date bDate;
    @ExcelVo(sort = 1,name = "c日期",dateFormat = "yyyy-MM-dd")
    private Date cDate;
    @ExcelVo(sort = 1,name = "d日期",dateFormat = "yyyy/MM/dd")
    private Date dDate;
    @ExcelVo(sort = 6,name="布尔")
    private Boolean isDelete = false;

    public Student(int id, String name, String sex,Date bDate,Date cDate,Date dDate) {
        this.id = id;
        this.name = name;
        this.sex = sex;
        this.bDate=bDate;
        this.cDate=cDate;
        this.dDate=dDate;
    }

    public int getId() {
        return id;
    }

    public void setId(int id) {
        this.id = id;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public String getSex() {
        return sex;
    }

    public void setSex(String sex) {
        this.sex = sex;
    }

    public Date getbDate() {
        return bDate;
    }

    public void setbDate(Date bDate) {
        this.bDate = bDate;
    }

    public Date getcDate() {
        return cDate;
    }

    public void setcDate(Date cDate) {
        this.cDate = cDate;
    }

    public Date getdDate() {
        return dDate;
    }

    public void setdDate(Date dDate) {
        this.dDate = dDate;
    }

    public Boolean getDelete() {
        return isDelete;
    }

    public void setDelete(Boolean delete) {
        isDelete = delete;
    }
}
