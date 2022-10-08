package com.example.demo.Service;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.write.metadata.WriteSheet;
import com.example.demo.Bean.Basic;
import com.example.demo.Bean.Person;
import jdk.internal.util.xml.impl.Input;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;

import java.io.*;
import java.net.HttpURLConnection;
import java.net.MalformedURLException;
import java.net.URL;
import java.net.URLConnection;
import java.text.SimpleDateFormat;
import java.util.*;

@Service("mainService")
public class MainService {
    /**
     *
     * @param handExcel 收集的excel表格
     * @param holdPosition  统计结果保存位置
     * @param imgPosition   截屏保存位置
     * @return
     */
    public boolean collectImg(String handExcel,String holdPosition,String imgPosition){
        boolean flag=true;
        System.out.println("访问成功");

        File infoExcel=new File(handExcel),holdExcel=new File(holdPosition);
        if (!infoExcel.exists()) return false;

        List<Person> personList=getTodayPersonInfo(handExcel);//获取所有当天所有学生信息

        savePhoto(personList,imgPosition);//保存所有学生核酸检测截屏

        Map<String,List<Person>> map=classifyByBackTime(personList);//数据信息分类

       if (holdExcel.isDirectory()){//统计结果文件的预处理
           holdPosition=pretreatment("软工205核酸统计",holdPosition, map.keySet());
           holdExcel=new File(holdPosition);
           //填充基本信息
           for (String backTime: map.keySet()){
               fillBasicInfoForSheet(backTime,holdPosition,map.get(backTime));
           }
       }

        censusTodayCondition(holdPosition,map);//统计当天核酸情况

        return flag;
    }

    /**
     * 统计当天核酸情况
     * @param filePath  excel路径
     * @param map   统计信息
     */
    private void censusTodayCondition(String filePath,Map<String,List<Person>> map){
        FileInputStream inputStream=null;
        try {
            inputStream=new FileInputStream(filePath);
            XSSFWorkbook workbook=new XSSFWorkbook(inputStream);
            Sheet curSheet=null;
            for (String backTime: map.keySet()){
                curSheet=workbook.getSheet(backTime);
                fillSheet(filePath,curSheet,map.get(backTime),workbook);
            }
        }catch (Exception e){
            e.printStackTrace();
        }finally {
            try {
                if (inputStream!=null) inputStream.close();
            }catch (IOException e){
                e.printStackTrace();
            }
        }
    }

    /**
     * 获取学生学号
     * @param personList
     * @return
     */
    private List<Integer> getPersonCode(List<Person> personList){
        List<Integer> list=new ArrayList<>();
        for (Person person:personList)
            list.add(Integer.valueOf(person.getCode()));
        return list;
    }

    /**
     * 填充sheet中每个人的核酸情况
     * @param filePath  用于输入流操作
     * @param sheet 当前sheet
     * @param personList    属于当前sheet的学生
     * @param workbook  写对象
     */
    private void fillSheet(String filePath,Sheet sheet,List<Person> personList,XSSFWorkbook workbook){
        SimpleDateFormat simpleDateFormat=new SimpleDateFormat("MM.dd日核酸");
        String header= simpleDateFormat.format(new Date());
        FileOutputStream outputStream=null;
        Row firstRow=sheet.getRow(0);
        int nullCol=0;
        while (firstRow.getCell(nullCol)!=null) nullCol++;//最后一列
        CellStyle cellStyle=firstRow.getCell(0).getCellStyle();//添加表头
        Cell firstRowCell = firstRow.createCell(nullCol);
        firstRowCell.setCellValue(header);
        firstRowCell.setCellStyle(cellStyle);

        XSSFCellStyle centerStyle=workbook.createCellStyle();//居中样式
        centerStyle.setAlignment(HorizontalAlignment.CENTER);

        int lastRow=sheet.getLastRowNum();
        Row curRow=null;
        List<Integer> codeList=getPersonCode(personList);
        for (int i=1;i<=lastRow;i++){
            curRow=sheet.getRow(i);
            Integer cellCode=(int)(curRow.getCell(1).getNumericCellValue());
            Cell curRowCell = curRow.createCell(nullCol);
            curRowCell.setCellStyle(centerStyle);
            if (codeList.contains(cellCode)) curRowCell.setCellValue("√");
        }
        try {
            outputStream=new FileOutputStream(filePath);
            workbook.write(outputStream);
            outputStream.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    /**
     * 根据返校时间填充基本信息
     * @param sheetName 返校时间为sheet名
     * @param filePath  excel路径
     * @param personList    返校时间内的所有学生
     */
    private void fillBasicInfoForSheet(String sheetName,String filePath,List<Person> personList){
        FileInputStream inputStream=null;
        FileOutputStream outputStream=null;
        try {
            inputStream=new FileInputStream(filePath);
            XSSFWorkbook workbook=new XSSFWorkbook(inputStream);
            Sheet sheet=workbook.getSheet(sheetName);
            XSSFCellStyle cellStyle=workbook.createCellStyle();//居中样式
            cellStyle.setAlignment(HorizontalAlignment.CENTER);
            for (int i=0;i<personList.size();i++){
                Row curRow= sheet.createRow(i+1);
                Cell cell1=curRow.createCell(0);
                cell1.setCellValue("软工205");
                cell1.setCellStyle(cellStyle);
                Cell cell2=curRow.createCell(1);
                cell2.setCellValue(Integer.valueOf(personList.get(i).getCode()));
                cell2.setCellStyle(cellStyle);
                Cell cell3=curRow.createCell(2);
                cell3.setCellValue(personList.get(i).getName());
                cell3.setCellStyle(cellStyle);
            }
            sheet.setColumnWidth(1,3500);//列宽
            outputStream=new FileOutputStream(filePath);
            workbook.write(outputStream);
        }catch (Exception e){
            e.printStackTrace();
        }finally {
            try {
                outputStream.close();
                inputStream.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }

    /**
     * 根据返校时间对学生进行分类
     * @param personList    所以学生信息
     * @return
     */
    private Map<String,List<Person>> classifyByBackTime(List<Person> personList){
        Map<String,List<Person>> map=new HashMap<>();
        SimpleDateFormat simpleDateFormat=new SimpleDateFormat("MM.dd");//时间格式化
        String backTime="";
        List<Person> curList=null;
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
        return map;
    }

    /**
     * 对统计结果的保存位置进行预处理
     * @param fileName  统计结果文件名
     * @param holdPosition  统计结果路径
     * @param set   sheet名
     * @return  最终存储的位置
     */
    private String pretreatment(String fileName, String holdPosition, Set<String> set){
        String position=holdPosition;
        position+="\\"+fileName+".xlsx";//新建的excel地址
        ExcelWriter writer = EasyExcel.write(position, Basic.class).build();
        WriteSheet sheet=null;
        for (String sheetName:set){//创建不同sheet
            sheet = EasyExcel.writerSheet(sheetName).build();
            writer.write(null,sheet);
        }
        writer.finish();
        return position;
    }

    /**
     *
     * @param handExcel 当天收集的学生信息
     * @return
     */
    private List<Person> getTodayPersonInfo(String handExcel){
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

    /**
     *
     * @param personList    所有学生信息
     * @param imgPosition   图片保存路径
     * @return
     */
    private void savePhoto(List<Person> personList, String imgPosition){//保存截屏
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
            System.out.println("正在下载"+person.getName()+"核酸截屏……");
            save(img,person.getImgUrl());
            System.out.println("下载完成");
        }
    }

    /**
     *
     * @param img   图片文件
     * @param downloadUrl   图片url
     */
    private void save(File img,String downloadUrl){
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
}
