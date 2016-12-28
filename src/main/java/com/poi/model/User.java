package com.poi.model;

import com.poi.dynamics.excel.build.annotation.ExcelCellHeader;

import java.util.Date;

/**
 * Created by lfwang on 2016/12/28.
 */
public class User {

    @ExcelCellHeader("姓名")
    private String name;

    @ExcelCellHeader("年龄")
    private Integer age;

    @ExcelCellHeader("生日")
    private Date birth;

    public User() {}

    public User(String name, Integer age, Date birth) {
        this.name = name;
        this.age = age;
        this.birth = birth;
    }


    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public Integer getAge() {
        return age;
    }

    public void setAge(Integer age) {
        this.age = age;
    }

    public Date getBirth() {
        return birth;
    }

    public void setBirth(Date birth) {
        this.birth = birth;
    }
}
