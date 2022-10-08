package com.example.demo.Test;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.write.metadata.WriteSheet;
import com.deepoove.poi.data.style.CellStyle;
import com.example.demo.Bean.Basic;
import com.example.demo.Bean.Person;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.*;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class FunctionTest {

    public void createNullExcel(){//创建Excel中不同sheet及sheet中的表头，需传入sheet名称及表头基类
        String fileName="D:\\桌面\\temp\\软工205核酸统计.xlsx";
        String[] sheetName={"9.28返校","9.29返校","9.30返校"};
        ExcelWriter excelWriter = EasyExcel.write(fileName, Basic.class).build();//创建写对象
        for (String name:sheetName){
            WriteSheet sheet=EasyExcel.writerSheet(name).build();//创建sheet信息及sheet表头
            excelWriter.write(null,sheet);
        }
        excelWriter.finish();//关闭流
    }
    @Test
    public void fillBasicDate(){//填充基本数据信息
        createNullExcel();
        try {
            String fileName="D:\\桌面\\temp\\软工205核酸统计.xlsx";
            FileInputStream inputStream=new FileInputStream(fileName);
            XSSFWorkbook workbook=new XSSFWorkbook(inputStream);
            Sheet curSheet=workbook.getSheet("9.28返校");

            curSheet.setColumnWidth(1,3500);//设置列宽及水平居中
            XSSFCellStyle cellStyle=workbook.createCellStyle();//居中样式
            cellStyle.setAlignment(HorizontalAlignment.CENTER);

            int i=1;
            Map<Integer,String> map=getBasicDate();
            for (Integer integer: map.keySet()){
                Row curRow=curSheet.createRow(i++);
                Cell classCell=curRow.createCell(0);
                classCell.setCellValue("软工205");//设置班级
                classCell.setCellStyle(cellStyle);
                Cell codeCell = curRow.createCell(1);//设置学号值、并居中
                codeCell.setCellValue(integer);
                codeCell.setCellStyle(cellStyle);
                Cell nameCell=curRow.createCell(2);//姓名
                nameCell.setCellValue(map.get(integer));
                nameCell.setCellStyle(cellStyle);
            }
            inputStream.close();
            FileOutputStream outputStream=new FileOutputStream(fileName);
            workbook.write(outputStream);
            outputStream.close();

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public Map<Integer,String> getBasicDate() throws Exception {//获取基本信息
        Map<Integer,String> map=new HashMap<>();
        String fileName="D:\\桌面\\temp\\软工205通讯录.xlsx";
        File file=new File(fileName);
        FileInputStream inputStream=new FileInputStream(fileName);
        XSSFWorkbook workbook=new XSSFWorkbook(inputStream);
        Sheet sheet=workbook.getSheetAt(0);
        for (int i=1;;i++){
            Row curRow=sheet.getRow(i);
            if (curRow==null) break;
            Integer code=(int)curRow.getCell(0).getNumericCellValue();
            String name=curRow.getCell(1).getStringCellValue();
            map.put(code,name);
        }
        return map;
    }

    @Test
    public void classifyTime(){
        Map<String, List<Person>> map=new HashMap<>();
        SimpleDateFormat simpleDateFormat=new SimpleDateFormat("MM.dd");//时间格式化
        String backTime="";
        List<Person> curList=null;
        List<Person> personList=getTodayPersonInfo();
        for (Person person:personList){
            backTime=simpleDateFormat.format(person.getBackTime());
            if (map.containsKey(backTime)) {
                curList=map.get(backTime);
                curList.add(person);
            }else {
                curList=new ArrayList<>();
                curList.add(person);
                map.put(backTime,curList);
            }
        }
        System.out.println(" ");
    }

    public List<Person> getTodayPersonInfo(){//测试获取当天所有学生提交的信息
        String handExcel="D:\\桌面\\temp\\10.5核酸统计（收集结果）.xlsx";
        List<Person> personList=new ArrayList<>();
        FileInputStream inputStream = null;
        try {
            inputStream=new FileInputStream(handExcel);
            XSSFWorkbook workbook=new XSSFWorkbook(inputStream);
            Sheet sheet=workbook.getSheetAt(0);
            for (int i=1;;i++){//获取每行信息，并封装为对象
                Row curRow=sheet.getRow(i);
                if (curRow==null) break;
                Person curPerson=new Person(curRow.getCell(0).getDateCellValue(),
                        String.valueOf((int)curRow.getCell(1).getNumericCellValue()),
                        curRow.getCell(2).getStringCellValue(),
                        curRow.getCell(3).getDateCellValue(),
                        curRow.getCell(4).getStringCellValue(),
                        curRow.getCell(5).getStringCellValue());
                personList.add(curPerson);
            }
        } catch (Exception e) {
            e.printStackTrace();
        }finally {
            try {
                inputStream.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
        return personList;
    }
}
