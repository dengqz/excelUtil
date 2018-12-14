package com.example.demo;

import com.example.demo.annotations.ExcelVo;

import java.util.Date;

/**
 * 导出实体
 * @author dqz.
 */
public class User {

    private String id;
    @ExcelVo(name = "姓名",sort = 1)
    private String name;
    @ExcelVo(name = "性别",sort = 2)
    private String sex;
    @ExcelVo(name = "生日",sort = 3)
    private Date birthDay;

    public User(String id, String name, String sex, Date birthDay) {
        this.id=id;
        this.name=name;
        this.sex=sex;
        this.birthDay=birthDay;
    }


    public User() {

    }


    public String getId() {
        return id;
    }

    public void setId(String id) {
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

    public Date getBirthDay() {
        return birthDay;
    }

    public void setBirthDay(Date birthDay) {
        this.birthDay = birthDay;
    }
}

