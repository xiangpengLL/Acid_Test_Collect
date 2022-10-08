package com.example.demo.ServiceTest;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.write.metadata.WriteSheet;
import com.example.demo.Bean.Basic;
import com.example.demo.Bean.Person;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.*;
import java.net.HttpURLConnection;
import java.net.URL;
import java.text.SimpleDateFormat;
import java.util.*;

public class MainTest {

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

    @Test
    public void time(){
        SimpleDateFormat simpleDateFormat=new SimpleDateFormat("MM.dd");
        System.out.println(simpleDateFormat.format(new Date()));
    }

    @Test
    public void download(){//下载图片
        List<Person> personList=getTodayPersonInfo();
        String imgPosition="D:\\桌面\\temp\\截屏保存";
        String curPersonDir="",curPersonImg="",nowTime=new SimpleDateFormat("YYYY年MM月dd日").format(new Date());
        for (Person person:personList){
            curPersonDir=imgPosition+"\\"+person.getCode()+person.getName();
            File dir=new File(curPersonDir);
            if (!dir.exists()){
                dir.mkdir();
            }
            curPersonImg=curPersonDir+"\\"+nowTime+".png";
            File img=new File(curPersonImg);
            if (!img.exists()) {
                try {
                    img.createNewFile();
                } catch (Exception e) {
                    e.printStackTrace();
                }
            }
            save(img,person.getImgUrl());
        }
    }

    private void save(File img,String downloadUrl){//保存
        HttpURLConnection connection=null;
        InputStream inputStream=null;
        FileOutputStream outputStream=null;
        try {
            URL url=new URL(downloadUrl);
            connection= (HttpURLConnection) url.openConnection();
            connection.setRequestMethod("GET");
            connection.setConnectTimeout(10*1000);
            connection.setReadTimeout(10*1000);
            inputStream=connection.getInputStream();
            outputStream=new FileOutputStream(img);
            byte[] buffer=new byte[1024];
            int len;
            while ((len=inputStream.read(buffer))!=-1){
                outputStream.write(buffer,0,len);
            }
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            if (inputStream!=null){
                try {
                    inputStream.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
            if (connection!=null) connection.disconnect();
            if (outputStream!=null){
                try {
                    outputStream.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }
    }

    @Test
    public void judgeDir(){
        String filePath="D:\\桌面\\temp\\20.xlsx";
        File file=new File(filePath);
        if (file.isFile()) System.out.println("该路径为文件");
        else if (file.isDirectory()) System.out.println("该路径为文件夹");
    }

    @Test
    public void pretreatment(){
        String fileName="软工205核酸统计",holdPosition="D:\\桌面\\temp",position=holdPosition;
        Set<String> set=new HashSet<>();
        set.add("9.28返校");
        set.add("9.29返校");
        set.add("9.30返校");
        position+="\\"+fileName+".xlsx";//新建的excel地址
        ExcelWriter writer = EasyExcel.write(position, Basic.class).build();//创建写对象
        WriteSheet sheet=null;
        for (String sheetName:set){
            sheet = EasyExcel.writerSheet(sheetName).build();
            writer.write(null,sheet);
        }
        writer.finish();
    }
}
