package com.yiqiang.learn.poi.entity;

import com.yiqiang.learn.poi.annotation.ExcelResources;

/**
 * Title:
 * Description:
 * Create Time: 2018/10/22 1:03
 *
 * @author: YEEChan
 * @version: 1.0
 */
public class Student {

    private int id;

    private String name;

    private String no;

    private String sex;

    public Student() {
    }

    public Student(int id, String name, String no, String sex) {
        this.id = id;
        this.name = name;
        this.no = no;
        this.sex = sex;
    }

    @ExcelResources(title = "学生标识", order = 1)
    public int getId() {
        return id;
    }

    public void setId(int id) {
        this.id = id;
    }

    @ExcelResources(title = "学生姓名", order = 2)
    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    @ExcelResources(title = "学生学号", order = 3)
    public String getNo() {
        return no;
    }

    public void setNo(String no) {
        this.no = no;
    }

    @ExcelResources(title = "学生性别", order = 4)
    public String getSex() {
        return sex;
    }

    public void setSex(String sex) {
        this.sex = sex;
    }

    @Override
    public String toString() {
        return "Student{" +
                "id=" + id +
                ", name='" + name + '\'' +
                ", no='" + no + '\'' +
                ", sex='" + sex + '\'' +
                '}';
    }
}
