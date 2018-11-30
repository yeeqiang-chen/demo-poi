package com.yiqiang.learn.poi;

import com.yiqiang.learn.poi.entity.Student;
import com.yiqiang.learn.poi.entity.User;
import com.yiqiang.learn.poi.util.ExcelTemplate;
import com.yiqiang.learn.poi.util.ExcelUtil;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellRangeAddress;
import org.junit.Test;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * Title:
 * Description:
 * Create Time: 2018/10/22 0:25
 *
 * @author: YEEChan
 * @version: 1.0
 */
public class ExcelTemplateTest {
    @Test
    public void test1() {
        ExcelTemplate et = ExcelTemplate.getInstance()
                .readTemplateByClasspath("/default.xls");
        et.createNewRow();
        et.createCell("1111111");
        et.createCell("aaaaaaaaaaaa");
        et.createCell("a1");
        et.createCell("a2a2");
        et.createNewRow();
        et.createCell("222222");
        et.createCell("bbbbb");
        et.createCell("b");
        et.createCell("dbbb");
        et.createNewRow();
        et.createCell("3333333");
        et.createCell("cccccc");
        et.createCell("a1");
        et.createCell(12333);
        et.createNewRow();
        et.createCell("4444444");
        et.createCell("ddddd");
        et.createCell("a1");
        et.createCell("a2a2");
        et.createNewRow();
        et.createCell("555555");
        et.createCell("eeeeee");
        et.createCell("a1");
        et.createCell(112);
        et.createNewRow();
        et.createCell("555555");
        et.createCell("eeeeee");
        et.createCell("a1");
        et.createCell("a2a2");
        et.createNewRow();
        et.createCell("555555");
        et.createCell("eeeeee");
        et.createCell("a1");
        et.createCell("a2a2");
        et.createNewRow();
        et.createCell("555555");
        et.createCell("eeeeee");
        et.createCell("a1");
        et.createCell("a2a2");
        et.createNewRow();
        et.createCell("555555");
        et.createCell("eeeeee");
        et.createCell("a1");
        et.createCell("a2a2");
        et.createNewRow();
        et.createCell("555555");
        et.createCell("eeeeee");
        et.createCell("a1");
        et.createCell("a2a2");
        et.createNewRow();
        et.createCell("555555");
        et.createCell("eeeeee");
        et.createCell("a1");
        et.createCell("a2a2");
        Map<String,String> datas = new HashMap<String,String>();
        datas.put("title","测试用户信息");
        datas.put("date","2018-10-22 12:30");
        datas.put("dep","测试deptartment");
        et.replaceFinalData(datas);
        et.insertSer();
        et.writeToFile("e:/tmp/test01.xls");
    }

    @Test
    public void testObj2Xls() {
        List<User> users = new ArrayList<User>();
        users.add(new User(1,"aaa","水水水",11));
        users.add(new User(2,"sdf","水水水",11));
        users.add(new User(3,"sdfde","水水水",11));
        users.add(new User(4,"aaa","水水水",11));
        users.add(new User(54,"aaa","水水水",11));
        users.add(new User(16,"aaa","水水水",11));
        ExcelUtil.getInstance().exportObj2ExcelByTemplate(new HashMap<String,String>(),"/user.xls","e:/tmp/tus.xls", users, User.class, true, true);
    }

    @Test
    public void testObj2Xls2() {
        List<Student> stus = new ArrayList<Student>();
        stus.add(new Student(1,"张三","1123123", "男"));
        stus.add(new Student(2,"张三","1123123", "男"));
        stus.add(new Student(3,"张三","1123123", "男"));
        stus.add(new Student(4,"张三","1123123", "男"));
        ExcelUtil.getInstance().exportObj2Excel("e:/tmp/ss1.xls",stus, Student.class, false);
    }

    @Test
    public void testRead01() {
        List<Object> stus = ExcelUtil.getInstance().readExcel2ObjsByPath("e:/tmp/ss1.xls",Student.class);
        for(Object obj:stus) {
            Student stu = (Student)obj;
            System.out.println(stu);
        }
    }

    @Test
    public void testRead02() {
        List<Object> stus = ExcelUtil.getInstance().readExcel2ObjsByPath("e:/tmp/tus.xls",User.class,1,2);
        for(Object obj:stus) {
            User stu = (User)obj;
            System.out.println(stu);
        }
    }

    @Test
    public void test() throws IOException {
        Workbook workbook = WorkbookFactory.create(new File("e:/tmp/test.xlsx"));
        Sheet sheet = workbook.getSheetAt(0);
        List<CellRangeAddress> mergedRegions = sheet.getMergedRegions();
        mergedRegions.stream().forEach(region -> {
            System.out.println(region.formatAsString());
            System.out.println("----------------------------------------------");
            System.out.println("getFirstColumn = "+region.getFirstColumn());
            System.out.println("getFirstRow = "+region.getFirstRow());
            System.out.println("getLastColumn = "+region.getLastColumn());
            System.out.println("getLastRow = "+region.getLastRow());
            System.out.println("getNumberOfCells = "+region.getNumberOfCells());
            System.out.println(sheet.getRow(region.getFirstRow()).getCell(region.getFirstColumn()));
            System.out.println("----------------------------------------------");
        });
//        int numMergedRegions = sheet.getNumMergedRegions();
//        System.out.println(numMergedRegions);
        /*for (int i = 0; i < numMergedRegions; i++) {
            System.out.println("---------------"+i+"---------------");
            CellRangeAddress region = sheet.getMergedRegion(i);
            System.out.println("getFirstColumn = "+region.getFirstColumn());
            System.out.println("getFirstRow = "+region.getFirstRow());
            System.out.println("getLastColumn = "+region.getLastColumn());
            System.out.println("getLastRow = "+region.getLastRow());
            System.out.println("getNumberOfCells = "+region.getNumberOfCells());
            System.out.println(sheet.getRow(region.getFirstRow()).getCell(region.getFirstColumn()));
            System.out.println("--------------"+i+"-------------");
        }*/

    }
}
