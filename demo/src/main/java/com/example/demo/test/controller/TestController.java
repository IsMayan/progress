package com.example.demo.test.controller;

import com.example.demo.Utils.ExportExcel;
import com.example.demo.Utils.ImportExcel;
import com.example.demo.test.entity.TestEntity;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.util.*;

@Controller
@RequestMapping("/test")
public class TestController {

    @Autowired
    ExportExcel exportExcel;

    @Autowired
    ImportExcel importExcel;



    @RequestMapping(value = "/exportExcel")
    public void exportExcel(HttpServletResponse response,TestEntity vo) throws Exception {
        String[] headers = {"姓名", "学校", "年级", "班级"};
        String fileName = "学生表";
        List<TestEntity> list = new ArrayList<>();
        TestEntity test0 = new TestEntity();
        test0.setAge("18");
        test0.setGrade("五年级");
        test0.setSchool("测试小学");
        test0.setStudentName("张三");

        TestEntity test1 = new TestEntity();
        test1.setAge("17");
        test1.setGrade("五年级");
        test1.setSchool("测试小学");
        test1.setStudentName("李四");
        list.add(test0);
        list.add(test1);

        Map<String, Object> studentMap = new HashMap();
        studentMap.put("headers", headers);
        studentMap.put("dataList", list);
        studentMap.put("fileName", fileName);

        List<Map> mapList = new ArrayList();
        mapList.add(studentMap);
        exportExcel.exportMultisheetExcel(fileName, mapList, response);
    }

    @RequestMapping(value = "/readExcel")
    public void readExcel(@RequestParam(value="file",required = false) MultipartFile file,
                          HttpServletResponse response, HttpServletRequest request) throws Exception {

        TestEntity vo = new TestEntity();

        List<?> mapList = importExcel.importExcel(vo,file);
        List<TestEntity> list =(List<TestEntity>)mapList;
        int x = 0;
        for (TestEntity a:list){
            x++;
            System.err.println(x+"："+a);
        }

    }

}
